using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace LoggingSys
{
    internal enum MessageType
    {
        Error, Warning, Success, Debug, System, CriticalError, Empty
    }

#pragma warning disable CS8524
    internal class Output
    {
        /// <summary>
        /// Raw path string for creating logs.
        /// </summary>
        private readonly static string _logPathRaw = "Logs/{Date}";

        /// <summary>
        /// Handles writing to the log as well as printing to the console with proper formatting.
        /// </summary>
        /// <param name="message">Message to be logged.</param>
        /// <param name="messageType">The type of message being logged.</param>
        /// <param name="printToConsole">Boolean value representing if the log goes to the Console.</param>
        /// <param name="addToLog">Boolean value representing if the log goes to the File.</param>
        internal async static Task Print(string message, MessageType messageType = MessageType.Empty, bool printToConsole = true, bool addToLog = true,[CallerLineNumber] int line = 0, [CallerMemberName] string caller = "", [CallerFilePath] string callerFilePath = "")
        {
            //Format the message for when we print it.
            message = message.FormatForPrint(messageType, line, caller, callerFilePath);

            //If it's to the console, then get the console ready to color code it.
            if (printToConsole)
                PrintToConsole(message, messageType);

            if (addToLog) 
                await WriteLog(message);
        }

        /// <summary>
        /// Sets the Console Color based on the message type, then prints the desired message to the console.
        /// </summary>
        /// <param name="message">Message to be printed.</param>
        /// <param name="messageType">The type of message it is.</param>
        internal static void PrintToConsole(string message, MessageType messageType)
        {
            //Set the console color based on the message type.
            Console.ForegroundColor = messageType switch
            {
                MessageType.Success => ConsoleColor.Green,
                MessageType.Warning => ConsoleColor.Yellow,
                MessageType.Error => ConsoleColor.Red,
                MessageType.CriticalError => ConsoleColor.DarkRed,
                MessageType.Debug => ConsoleColor.Blue,
                MessageType.System => ConsoleColor.Cyan,
                MessageType.Empty => ConsoleColor.Gray
            };
            Console.WriteLine(message + "\n"); //print the result.
            //Set the console back to default.
            Console.ForegroundColor = ConsoleColor.Gray;
        }

        /// <summary>
        /// Handles printing End-Of-Log for this session.
        /// </summary>
        internal static async Task PrintLogEnd()
        {
            Console.WriteLine("Ending log...");

            //Generates the first part of the message
            string message = $"Log End: {DateTime.Now}.\n";

            //Write the end log
            await WriteLog(message);
        }

        /// <summary>
        /// Writes the desired message to the log file.
        /// </summary>
        /// <param name="message">Message to be printed.</param>
        private static async Task WriteLog(string message)
        {
            //This will be used later to determine if we are creating a new file for the first time. The purpose is to set the first line of the log file.
            bool logCreated = false;
            //This will be the path of the log we are writing to currently.
            string truePath = _logPathRaw.Replace("{Date}", $"{DateTime.Now.DayOfWeek}-{DateTime.Now.Day}-{DateTime.Now:MMMM}-{DateTime.Now.Year}.log");
            //This will be the path for the log folder itself, where we will store all future logs.
            string logFolderPath = _logPathRaw.Replace("/{Date}", "");

            try
            {
                //If we need to create the directory, then create that,
                if (!Directory.Exists(logFolderPath))
                    Directory.CreateDirectory(logFolderPath);

                //If the file needed doesn't already exist,
                if (!File.Exists(truePath))
                {
                    var creationStream = File.Create(truePath);
                    creationStream.Flush();
                    creationStream.Close();
                } //Create it.

                //Write the log.
                await using (StreamWriter writer = new(truePath, true))
                {
                    //We now use the variable from before to set the first line of a new log.
                    if (logCreated)
                        writer.Write("==================================Start of Log==================================");
                    writer.WriteLine(message);
                    writer.Flush();
                    writer.Close();
                }
            }
            catch (Exception e)
            {
                await Print("There was an unknown error when writing the logs!\n\n" + e.Message, MessageType.CriticalError, addToLog: false);
            }
        }
    }

    internal static class Extensions
    {
        /// <summary>
        /// Structures the message for the log format.
        /// </summary>
        /// <param name="message">Event message</param>
        /// <param name="messageType">Message type</param>
        /// <param name="line">Line the event was triggered</param>
        /// <param name="caller">Caller of the triggering event</param>
        /// <returns>The formatted message</returns>
        internal static string FormatForPrint(this string message, MessageType type, int line, string caller, string callerFilePath)
        {
            message = $"{caller} ({callerFilePath})\t[Line: {line}]\t[{DateTime.Now}]:\n{message}";

            if (type != MessageType.Empty)
                message = $"[{type}]\t{message}";

            return message;
        }

        /// <summary>
        /// Loops through and performs actions for each element in an IEnumerable<T>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="items"></param>
        /// <param name="action"></param>
        internal static void Each<T>(this IEnumerable<T> items, Action<T> action)
        {
            foreach (var item in items)
                action(item);
        }

        /// <summary>
        /// Removes system specific information from the data.
        /// </summary>
        /// <param name="original">The original string value that will be trimmed.</param>
        /// <returns>The trimmed value</returns>
        internal static string TrimM8(this string original) =>
            original.Replace("M8", "");

        /// <summary>
        /// Checks if <see langword="T"/> is null
        /// </summary>
        /// <typeparam name="T">The given type.</typeparam>
        /// <param name="field">The variable we wish to check for being null.</param>
        /// <returns>
        /// <para><see langword="True"/> if the value is default of <see langword="T"/> or null.</para>
        /// <para><see langword="False"/> if the value is not default of <see langword="T"/> or null.</para>
        /// </returns>
        public static bool DefaultOrNull<T>(this T field)
        {
            if ((field == null) || field.Equals(default(T))) return true;
            if (field.GetType() == typeof(string))
                if (field.Equals(string.Empty)) return true;
            return false;
        }

        public static string TrimSpacing(this string text)
        {
            if (!text.DefaultOrNull())
                text = Regex.Replace(text, @"\s+", "");
            return text;
        }
    }
}
