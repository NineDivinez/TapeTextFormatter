using LoggingSys;
using TapeTextFormatter;
using OfficeOpenXml;
using System.Data;
using Configuration;
using FileSaving;

namespace Main
{
#pragma warning disable CS8602
#pragma warning disable CS8618
#pragma warning disable CS8600
    internal class Program
    {
        private readonly static ConfigReader _config = new();
        private readonly static ExcelSaver _excelSaver = new();
        private readonly static TextSaver _textSaver = new();

        /// <summary>
        /// Contains the directory (NOT A FILE) for the output.
        /// </summary>
        private static string outputDirectory;
        /// <summary>
        /// Contains the directory (NOT A FILE) for the input.
        /// </summary>
        private static string inputDirectory;

        internal static void Main(string[] args)
        {
            outputDirectory = _config.GetDestination(ConfigReader.Destinations.OutputForLexington);
            //outputDirectory = _config.GetDestination(ConfigReader.Destinations.OutputForChaska);
            if (outputDirectory.DefaultOrNull())
                outputDirectory = _config.GetDestination(ConfigReader.Destinations.DefaultOutput);

            #region Desired List Construction
            //Get the desired list from the user.
            Logging.Print("Please paste in the list of tape names.\nEnter \"Input\" if you want placed a file with the info in the Input folder.", MessageType.System).GetAwaiter().GetResult();
            //If the input is invalid, go back to recapture it.
            AwaitInput:
            List<string> desiredListOfTapeNames = AwaitEntries();

            if (File.Exists(desiredListOfTapeNames[0]) || Directory.Exists(desiredListOfTapeNames[0]) || desiredListOfTapeNames[0].ToLower() == "input")
            {
                var fileAttributes = File.GetAttributes(desiredListOfTapeNames[0]);
                if (fileAttributes.HasFlag(FileAttributes.Directory) && desiredListOfTapeNames[0].ToLower() != "input")
                {
                    Logging.Print("Please enter a file path, not a directory!", MessageType.Warning).GetAwaiter().GetResult();
                    goto AwaitInput;
                }
                //This will be the final destination we are reading. Set it to first entry and then check if it needs to change.
                string inputFileDestination = desiredListOfTapeNames[0];
                
                //Read the file in the input folder.
                if (desiredListOfTapeNames[0].ToLower() == "input") 
                {
                    //Finds all files in the folder, then filters out any that are text files.
                    inputDirectory = _config.GetDestination(ConfigReader.Destinations.InputFolder);
                    inputFileDestination = Directory.GetFiles(inputDirectory).First(file => file.EndsWith(".txt"));
                    //If we did not find anything, inform the user to try again.
                    if (inputFileDestination.DefaultOrNull())
                    {
                        //Print summary of what went wrong.
                        Logging.Print($"Input folder either does not exist or is empty. Please ensure this is not the case and try again.", MessageType.Warning).GetAwaiter().GetResult();
                        //Silent debug log
                        Logging.Print($"User Input: {string.Join(", ", desiredListOfTapeNames)}\nExtracted: {inputFileDestination}", MessageType.Debug, printToConsole: false).GetAwaiter().GetResult();
                        goto AwaitInput;
                    }
                }

                try
                {
                    desiredListOfTapeNames = File.ReadAllLines(inputFileDestination).ToList();
                }
                catch(Exception ex)
                {
                    Logging.Print("Unknown error when processing. Going back to recapture input.", MessageType.Warning).GetAwaiter().GetResult();
                    Logging.Print(ex.Message, MessageType.Error, printToConsole: false).GetAwaiter().GetResult(); //Adds the error message silently to the log for review.
                    goto AwaitInput;
                }
            }
            else
            {
                if (desiredListOfTapeNames.Any(val => val.ContainsNoSpecialCharacters()))
                {
                    Logging.Print("These should not have special characters. Please try the entry again.", MessageType.Warning).GetAwaiter().GetResult();
                    goto AwaitInput;
                }
            }
            #endregion

            //Sorts the list in alphabetical order.
            desiredListOfTapeNames.Sort();//desiredListOfTapeNames.Each(val => val = val.TrimM8()); //Commented out temporarily since this is not a needed feature yet. Just shows where we need it.

            //Silent log print
            Logging.Print("Found based on user input: " + string.Join(", ", desiredListOfTapeNames), MessageType.System, printToConsole: false).GetAwaiter().GetResult();

            #region Possible List Construction
            //Get the Excel Sheet from the user.
            Logging.Print("Please paste the directory to the Excel Spreadsheet.", MessageType.System).GetAwaiter().GetResult();
        ExcelSheetEntry:
            string excelDestination = Console.ReadLine();

            //Log print
            Logging.Print("User entered: " + excelDestination, MessageType.System, false).GetAwaiter().GetResult();

            List<TapeData> unfilteredTapeDataList;

            //Verifies the entry is valid

            excelDestination = excelDestination.Replace("\"", "");
            if (File.Exists(excelDestination) || excelDestination.ToLower() == "input")
            {
                //Allows the user to specify "Whatever is in the input folder"
                if (excelDestination.ToLower() == "input")
                {
                    //Finds all files in the folder, then filters out any that are text files.
                    string inputFileDestination = Directory.GetFiles(excelDestination).First(file => !file.EndsWith(".txt"));
                    //If we did not find anything, inform the user to try again.
                    if (inputFileDestination.DefaultOrNull())
                    {
                        Logging.Print($"Input folder either does not exist or is empty. Please ensure this is not the case and try again.", MessageType.Warning).GetAwaiter().GetResult();
                        goto ExcelSheetEntry;
                    }
                    excelDestination = inputFileDestination;
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package;

                //Determine our format
                if (excelDestination.Contains(".csv")) //If it's a CVS file, open it and convert it to an Excel sheet
                    package = ConvertCsvToExcel(excelDestination);
                else package = new(excelDestination); //Otherwise just open it as an excel sheet.

                //Create all the tape data objects
                ExcelWorksheet sheet = package.Workbook.Worksheets[0];
                /*TODO: Take the columns we want to search for from the config.*/
                unfilteredTapeDataList = ExtractData(sheet, _config.GetColumns());

                //Sort the list alphabetically
                unfilteredTapeDataList.Sort((x, y) => string.Compare(x.name, y.name));

                //Logging the extracted data for debugging
                Logging.Print("Tape names found: " + string.Join(", ", unfilteredTapeDataList), MessageType.Debug, false).GetAwaiter().GetResult();
                #endregion

                #region Filter Lists
                //Filter out the ones we don't need
                List<TapeData> filteredList = new();
                foreach (var candidate in unfilteredTapeDataList)
                {
                    desiredListOfTapeNames.Each(entry =>
                    {
                        if (entry.ToLower() == candidate.name.ToLower())
                            filteredList.Add(candidate);
                    });
                }

                //Logging the filtered list for debugging
                Logging.Print("Filtered list: " + string.Join(", ", filteredList), MessageType.Debug, false).GetAwaiter().GetResult();

                using (ExcelPackage output = new())
                {
                    ExcelWorksheet worksheet = output.Workbook.Worksheets.Add("Sheet1");
                    int columnIndex = 1;
                    foreach (var tape in filteredList)
                    {
                        worksheet.Cells[$"A{columnIndex}"].Value = tape.name;
                        worksheet.Cells[$"B{columnIndex}"].Value = tape.returnDate;
                        worksheet.Cells[$"C{columnIndex}"].Value = tape.description;
                        columnIndex++;
                    }

                    _excelSaver.SaveFileWhenReady(output, $"{outputDirectory}/{DateTime.Now.ToString("MM-dd-yyyy")}.xlsx");
                }

                string textFileLocation = $"{outputDirectory}/{DateTime.Now.ToString("MM-dd-yyyy")}.txt";
                _textSaver.SaveFileWhenReady(new(textFileLocation), textFileLocation, filteredList.JoinSpecificField("\n", tape => tape.name));
            }
            else
            {
                Logging.Print("No such file exists. Please try again.", MessageType.Warning).GetAwaiter().GetResult();
                goto ExcelSheetEntry;
            }
            #endregion

            Logging.Print("All finished. Check the Logging folder for the results :)", MessageType.System, true, false).GetAwaiter().GetResult();
            Logging.PrintLogEnd().GetAwaiter().GetResult();
        }
#pragma warning disable CS8619
        private static List<TapeData> ExtractData(ExcelWorksheet sheet, params string[] Columns)
        {
            List<string>[] data = new List<string>[Columns.Length];
            for (int i = 0; i < data.Length; i++)
                data[i] = new();

            List<TapeData> extractedData = new();

            for (int i = 0; i < Columns.Length; i++)
            {
                var rowRange = sheet.Cells[$"{Columns[i]}:{Columns[i]}"];
                List<string> foundValue = rowRange.Select(cell => cell.Value?.ToString()).ToList();

                foundValue.Each(val =>
                {
                    if (!val.DefaultOrNull())
                    {
                        val = val.Trim();
                        if (val.Length > 0)
                            data[i].Add(val);
                    }
                });
            }

            for (int i = 0; i < data[0].Count; i++)
                extractedData.Add(new(data[0][i].TrimSpacing(), data[1][i].TrimSpacing(), data[2][i].TrimSpacing()));

            return extractedData;
        }


#pragma warning disable CS8600
        /// <summary>
        /// Waits for the user to finish entering multiple entries.
        /// </summary>
        /// <returns>A List of entries provided by the user.</returns>
        private static List<string> AwaitEntries()
        {
            //Logging.PrintToConsole("Double press 'Return' when completed.", MessageType.System);
            List<string> inputs = new();
            while (true)
            {
                string entry = Console.ReadLine();

                if (string.IsNullOrEmpty(entry))
                    break;

                inputs.Add(entry);

                if (entry.ToLower() == "input" || entry.Contains(":\\"))
                    break;
            }
            return inputs; ;
        }


#pragma warning disable CS8600
        /// <summary>
        /// Converts the data from a text file to an excel sheet.
        /// </summary>
        /// <param name="filePath">The file path of the target.</param>
        /// <returns>The full Excel Sheet</returns>
        private static ExcelPackage ConvertCsvToExcel(string filePath)
        {
            DataTable dataTable = new();

            using (StreamReader reader = new(filePath))
            {
                string headerLine = reader.ReadLine();
                string[] headers = headerLine.Split(',');

                foreach (string header in headers)
                {
                    dataTable.Columns.Add(header);
                }

                while (!reader.EndOfStream)
                {
                    string dataLine = reader.ReadLine();
                    string[] dataValues = dataLine.Split(',');

                    // Adjust the number of columns if needed
                    if (dataValues.Length > dataTable.Columns.Count)
                    {
                        int diff = dataValues.Length - dataTable.Columns.Count;
                        for (int i = 0; i < diff; i++)
                        {
                            dataTable.Columns.Add(dataValues[i]);
                        }
                    }

                    dataTable.Rows.Add(dataValues);
                }
            }

            ExcelPackage package = new();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
            worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

            return package;
        }
    }
}