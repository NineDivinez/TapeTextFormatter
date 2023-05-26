using LoggingSys;
using System.Windows.Forms;
using TapeTextFormatter;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Data;

namespace Main
{
    internal class Program
    {
        //[STAThreadAttribute]
        internal static void Main(string[] args)
        {
            /*TODO: Add support for pasting a text file directory instead of the list itself AND being able to add the text file in the Input folder.*/
            //Get the desired list from the user.
            Output.Print("Please paste in the list of tape names we need.", MessageType.System).GetAwaiter().GetResult();
            List<string> desiredList = AwaitEntries();
            desiredList.Sort();
            desiredList.Each(val => val = val.TrimM8());

            //Log print
            Output.Print("User entered: " + string.Join(", ", desiredList), MessageType.System, false).GetAwaiter().GetResult();

            //Get the Excel Sheet from the user.
            Output.Print("Please paste the directory to the Excel Spreadsheet.", MessageType.System).GetAwaiter().GetResult();
        Input:
            string excelDestination = Console.ReadLine();

            //Log print
            Output.Print("User entered: " + excelDestination, MessageType.System, false).GetAwaiter().GetResult();

            List<TapeData> possibleData;
            if (File.Exists(excelDestination) || excelDestination == "Input")
            {
                if (excelDestination == "Input")
                {
                    string inputFileDestination = Directory.GetFiles(excelDestination).First().ToString();
                    if (inputFileDestination.DefaultOrNull())
                    {
                        Output.Print($"Input folder either does not exist or is empty. Please ensure this is not the case and try again.", MessageType.Warning).GetAwaiter().GetResult();
                        goto Input;
                    }
                    excelDestination = Directory.GetFiles(excelDestination).First();
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package;

                //Determine our format
                if (excelDestination.Contains(".csv")) //If it's a CVS file, open it and convert it to an Excel sheet
                    package = ConvertCsvToExcel(excelDestination);
                else package = new(excelDestination); //Otherwise just open it as an excel sheet.

                //Create all the tape data objects
                ExcelWorksheet sheet = package.Workbook.Worksheets[0];
                possibleData = ExtractData(sheet, "A", "B", "D");

                //Sort the list alphabetically
                possibleData.Sort((x, y) => string.Compare(x.tapeName, y.tapeName));

                Output.Print("Tape names found: " + string.Join(", ", possibleData), MessageType.Debug, false).GetAwaiter().GetResult();

                //Filter out the ones we don't need
                List<TapeData> filteredList = new();
                foreach (var candidate in possibleData)
                {
                    desiredList.Each(entry =>
                    {
                        if (entry.ToLower().Equals(candidate.tapeName.ToLower()))
                            filteredList.Add(candidate);
                    });
                }

                Output.Print("Filtered list: " + string.Join(", ", filteredList), MessageType.Debug, false).GetAwaiter().GetResult();

                using (ExcelPackage output = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = output.Workbook.Worksheets.Add("Sheet1");
                    int columnIndex = 1;
                    foreach (var tape in filteredList)
                    {
                        worksheet.Cells[$"A{columnIndex}"].Value = tape.tapeName;
                        worksheet.Cells[$"B{columnIndex}"].Value = tape.tapeReturnDate;
                        worksheet.Cells[$"C{columnIndex}"].Value = tape.tapeDescription;
                        columnIndex++;
                    }

                    output.SaveAs($"Output/{DateTime.Now.ToString("MM-dd-yyyy")}.xlsx");
                }

                using (StreamWriter writer = new($"Output/{DateTime.Now.ToString("MM-dd-yyyy")}.txt"))
                {
                    foreach (var tape in filteredList)
                        writer.WriteLine(tape.tapeName);
                }
            }
            else
            {
                Output.Print("No such file exists. Please try again.", MessageType.Warning).GetAwaiter().GetResult();
                goto Input;
            }

            Output.Print("All finished. Check the output folder for the results :)", MessageType.System, true, false).GetAwaiter().GetResult();
            Output.PrintLogEnd().GetAwaiter().GetResult();
        }


        private static List<TapeData> ExtractData(ExcelWorksheet sheet, params string[] Cells)
        {
            List<string>[] data = new List<string>[Cells.Length];
            List<TapeData> extractedData = new();

            for (int i = 0; i < Cells.Length; i++)
            {
                if (sheet.Columns.Any(val => val.Hidden))
                {
                    Output.Print("One or more of the columns are hidden!", MessageType.Warning).GetAwaiter().GetResult();
                    return null;
                }

                var columnARange = sheet.Cells[$"{Cells[i]}:{Cells[i]}"];
                data[i] = columnARange.Select(cell => cell.Value?.ToString()).ToList();
            }
            
            foreach (var column in  data)
            {
                column.RemoveAt(0);
                column.RemoveAt(0);
                column.RemoveAt(column.Count -1);
                column.RemoveAt(column.Count -1);
            }

            TapeData current = null;
            for (int i = 0; i < data[0].Count; i++ )
                extractedData.Add(new(data[0][i].TrimSpacing(), data[1][i].TrimSpacing(), data[2][i].TrimSpacing()));

            return extractedData;
        }

        /// <summary>
        /// Waits for the user to finish entering multiple entries.
        /// </summary>
        /// <returns>A List of entries provided by the user.</returns>
        private static List<string> AwaitEntries()
        {
            Output.PrintToConsole("Double press 'Return' when completed.", MessageType.System);
            List<string> inputs = new();
            while (true)
            {
#pragma warning disable CS8600
                string entry = Console.ReadLine();
#pragma warning restore CS8600

                if (string.IsNullOrEmpty(entry))
                    break;

                inputs.Add(entry);
            }
            return inputs; ;
        }

        /// <summary>
        /// Converts the data from a text file to an excel sheet.
        /// </summary>
        /// <param name="filePath">The file path of the target.</param>
        /// <returns>The full Excel Sheet</returns>
        private static ExcelPackage ConvertCsvToExcel(string filePath)
        {
            DataTable dataTable = new DataTable();

            using (StreamReader reader = new StreamReader(filePath))
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
                            string newColumnName = $"Column {dataTable.Columns.Count + i + 1}";
                            dataTable.Columns.Add(newColumnName);
                        }
                    }

                    dataTable.Rows.Add(dataValues);
                }
            }

            ExcelPackage package = new ExcelPackage();
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
            worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

            return package;
        }

        private static char IntToLetter(int number)
        {
            if (number < 1 || number > 26)
            {
                throw new ArgumentException("Invalid number. Number must be between 1 and 26.");
            }

            char letter = (char)('A' + (number - 1));
            return letter;
        }
    }
}