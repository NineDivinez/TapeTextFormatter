using OfficeOpenXml;
using LoggingSys;

namespace FileSaving
{
    internal class ExcelSaver
    {
        FileSystemWatcher watcher = new();

        public async void SaveFileWhenReady(ExcelPackage excelSheet, string saveDestination)
        {
                Retry:
            //Try to save it, if it's not catch the error
            try
            {
                excelSheet.SaveAs(saveDestination);
                await Logging.Print($"File saved to {saveDestination}.", MessageType.Success);
            }
            catch (InvalidOperationException exc)
            {
                watcher.Filter = "*.xlsx";
                watcher.Path = Directory.GetParent(saveDestination).FullName;

                await Logging.Print($"Unable to {saveDestination} file, as it is already in use. Please close the file so changes can be made.\nPress any key to cancel.", MessageType.Warning);
                await Logging.Print(exc.Message, MessageType.Debug, printToConsole: false);

                watcher.WaitForChanged(WatcherChangeTypes.Deleted);

                goto Retry;
            }
        }
    }
    internal class TextSaver
    {

        FileSystemWatcher watcher = new();

        public async void SaveFileWhenReady(StreamWriter writer, string saveDestination, string content)
        {
                Retry:
            //Try to save it, if it's not catch the error
            try
            {
                writer.Write(content);
                writer.Close();
                Logging.Print($"File saved to {saveDestination}.", MessageType.Success).GetAwaiter().GetResult();
            }
            #region Subscribe for when the file is available
            catch (IOException exc)
            {
                watcher.Filter = "*.txt";
                watcher.Path = Directory.GetParent(saveDestination).FullName;

                await Logging.Print($"Unable to save {saveDestination}, as it is already in use. Please close the file so changes can be made.", MessageType.Warning);
                await Logging.Print(exc.Message, MessageType.Debug);

                watcher.WaitForChanged(WatcherChangeTypes.Deleted);

                goto Retry;
            }
            #endregion
        }
    }
}
