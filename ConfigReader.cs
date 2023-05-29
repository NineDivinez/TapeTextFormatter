using IniParser;
using IniParser.Model;
using LoggingSys;

namespace Configuration
{
    internal class ConfigReader
    {
        /// <summary>
        /// Contains all the data read from the config file.
        /// </summary>
        private IniData configData;
        /// <summary>
        /// The location of the config file.
        /// </summary>
        private readonly string configSection = "ConfigData";

        /// <summary>
        /// All possible destinations for finding the specified folder from the config file.
        /// </summary>
        public enum Destinations
        {
            InputFolder, OutputForLexington, OutputForChaska, Default
        }

        /// <summary>
        /// Object that handles loading data from the configuration file.
        /// </summary>
        public ConfigReader()
        {
            FileIniDataParser parser = new();
            configData = parser.ReadFile("Config.ini");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="optionNeeded"></param>
        /// <returns></returns>
        internal string GetDestination(Destinations optionNeeded) =>
            configData[configSection][optionNeeded.ToString()];

        /// <summary>
        /// Gets the columns needed for reading all the data in the excel sheets and returns it as a string array.
        /// </summary>
        /// <returns>string array of all columns containing the data we need.</returns>
        internal string[] GetColumns()
        {
            string columnsToRead = configData[configSection]["ColumnsToRead"]; //Reads the data
            string[] columns = columnsToRead.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries); //Splits it in to an array.

            //Checking for errors
            if (columns.Length == 0) Logging.Print("Columns not found when reading config!", MessageType.CriticalError).GetAwaiter().GetResult();

            return columns; //Returns the final product.
        }
    }
}