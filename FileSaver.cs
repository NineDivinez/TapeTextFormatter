using OfficeOpenXml;
using LoggingSys;

namespace FileSaving
{
    internal class ExcelSaver
    {
        private ExcelPackage _package;
        private string _packageDestination;


        FileSystemWatcher watcher = new();

        public void SaveFileWhenReady(ExcelPackage excelSheet, string saveDestination)
        {
            //Try to save it, if it's not catch the error
            try
            {
                excelSheet.SaveAs(saveDestination);
                Logging.Print($"File saved to {saveDestination}.", MessageType.Success).GetAwaiter().GetResult();
            }
            #region Subscribe for when the file is available
            catch (IOException exc)
            {
                _package = excelSheet;
                _packageDestination = saveDestination;

                watcher.Filter = saveDestination;
                watcher.NotifyFilter = NotifyFilters.FileName;

                watcher.Changed += SaveFile;

                watcher.EnableRaisingEvents = true;

                Logging.Print($"Unable to {saveDestination} file, as it is already in use. Please close the file so changes can be made.", MessageType.Warning).GetAwaiter().GetResult();
                Logging.Print(exc.Message, MessageType.Debug).GetAwaiter().GetResult();
            }
            #endregion
        }
        private void SaveFile(object sender, FileSystemEventArgs e)
        {
            try
            {
                _package.Save(_packageDestination);
                watcher.EnableRaisingEvents = false;
            }
            catch(Exception ex)
            {
                Logging.Print("Error when saving file! Flagged as no longer in use, but could not save!", MessageType.CriticalError).GetAwaiter().GetResult();
                Logging.Print($"File Destination: {_packageDestination}", MessageType.Debug, printToConsole: false).GetAwaiter().GetResult();
            }
        }
    }
    internal class TextSaver
    {
        private StreamWriter _writer;
        private string _saveDestination;
        private string _content;

        FileSystemWatcher watcher = new();

        public void SaveFileWhenReady(StreamWriter writer, string saveDestination, string content)
        {
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
                _writer = writer;
                _saveDestination = saveDestination;
                _content = content;

                watcher.Filter = saveDestination;
                watcher.NotifyFilter = NotifyFilters.FileName;

                watcher.Changed += SaveFile;

                watcher.EnableRaisingEvents = true;

                Logging.Print($"Unable to save {saveDestination}, as it is already in use. Please close the file so changes can be made.", MessageType.Warning).GetAwaiter().GetResult();
                Logging.Print(exc.Message, MessageType.Debug).GetAwaiter().GetResult();
            }
            #endregion
        }
        private void SaveFile(object sender, FileSystemEventArgs e)
        {
            try
            {
                _writer.Write(_content);
                _writer.Close();
                Logging.Print($"File saved to {_saveDestination}.", MessageType.Success).GetAwaiter().GetResult();
                watcher.EnableRaisingEvents = false;
            }
            catch (Exception ex)
            {
                Logging.Print("Error when saving file! Flagged as no longer in use, but could not save!", MessageType.CriticalError).GetAwaiter().GetResult();
                Logging.Print($"File Destination: {_saveDestination}", MessageType.Debug, printToConsole: false).GetAwaiter().GetResult();
            }
        }
    }
}
