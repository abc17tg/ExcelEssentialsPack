using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ScintillaNET.Style;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportTableToExcel
{
    public static class WorksheetFromTxtCreator
    {
        public static async Task ImportTextFileToExcelLegacy(Excel.Worksheet worksheet, string filePath, char delimiter)
        {
            try
            {
                // Create a connection string to the text file using Power Query
                string connectionString = $"TEXT;{filePath}";
                int columnCount;
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string firstLine = reader.ReadLine();
                    columnCount = !string.IsNullOrEmpty(firstLine) ? firstLine.Split(delimiter).Length : 0;
                    if (columnCount < 1)
                        throw new Exception("File empty!");
                }

                try
                {
                    Excel.QueryTable queryTable = worksheet.QueryTables.Add(Connection: connectionString, Destination: worksheet.Cells[1, 1]);
                    queryTable.TextFileParseType = Excel.XlTextParsingType.xlDelimited;
                    queryTable.TextFileOtherDelimiter = delimiter.ToString();
                    queryTable.FieldNames = true;
                    queryTable.HasAutoFormat = false;
                    queryTable.PreserveFormatting = true;
                    var typesObjectsArray = new object[columnCount];
                    for (int i = 0; i < columnCount; i++)
                        typesObjectsArray[i] = Excel.XlColumnDataType.xlTextFormat;
                    queryTable.TextFileColumnDataTypes = typesObjectsArray;
                    queryTable.TextFileTextQualifier = Excel.XlTextQualifier.xlTextQualifierNone;
                    queryTable.Refresh();
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message, ex);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                var app = worksheet.Application;
                app.DisplayAlerts = false;
                worksheet.Delete();
                app.DisplayAlerts = true;
            }
        }

        public static async Task ImportTextFileToExcelWhateverItCan(Excel.Worksheet worksheet, string filePath, char delimiter)
        {
            try
            {
                DataTable dataTable = await ReadFileIntoDataTableWhateverItCan(filePath, delimiter);
                UtilsExcel.PasteDataTableToRange(dataTable, worksheet.Cells[1, 1]);
                if (!File.Exists(filePath))
                {
                    if (File.Exists(Path.ChangeExtension(filePath, "txt")))
                    {
                        filePath = Path.ChangeExtension(filePath, "txt");
                    }
                    else if (File.Exists(Path.ChangeExtension(filePath, "csv")))
                    {
                        filePath = Path.ChangeExtension(filePath, "csv");
                    }
                    else
                        return;
                }

                int linesInFile = File.ReadLines(filePath).Count();
                int dataRowsInFile = linesInFile > 0 ? linesInFile - 1 : 0;

                int rowsInTable = dataTable.Rows.Count;
                int missingRows = dataRowsInFile - rowsInTable;

                string message = $"File: {filePath}\nData Rows in File: {dataRowsInFile}\nRows Loaded in Table: {rowsInTable}";

                if (missingRows > 0)
                {
                    MessageBox.Show(message + $"\n\n⚠️ {missingRows} rows are missing!", "Import Stats", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show(message + "\n\n✅ All rows imported successfully.", "Import Stats", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                var app = worksheet.Application;
                app.DisplayAlerts = false;
                worksheet.Delete();
                app.DisplayAlerts = true;
            }
        }

        public static async Task ImportTextFileToExcel(Excel.Worksheet worksheet, string filePath, char delimiter)
        {
            try
            {
                DataTable dataTable = await ReadFileIntoDataTable(filePath, delimiter);
                UtilsExcel.PasteDataTableToRange(dataTable, worksheet.Cells[1, 1]);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                var app = worksheet.Application;
                app.DisplayAlerts = false;
                worksheet.Delete();
                app.DisplayAlerts = true;
            }
        }

        public static async Task ImportTextFileToExcelAdv(Excel.Worksheet worksheet, string filePath, char delimiter)
        {
            try
            {
                DataTable dataTable = await ReadFileIntoDataTableAdv(filePath, delimiter.ToString());
                UtilsExcel.PasteDataTableToRange(dataTable, worksheet.Cells[1, 1]);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Import Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                var app = worksheet.Application;
                app.DisplayAlerts = false;
                worksheet.Delete();
                app.DisplayAlerts = true;
            }
        }

        public static async Task<DataTable> ReadFileIntoDataTableWhateverItCan(string filePath, char delimiter)
        {
            DataTable dataTable = new DataTable();
            List<object[]> allData = new List<object[]>();

            using (StreamReader sr = new StreamReader(filePath, true))
            {
                // Read the first line to create column headers
                string headersLine = await sr.ReadLineAsync();
                if (headersLine == null) return dataTable; // Empty file, return empty DataTable

                string[] headers = headersLine.Split(delimiter)
                                              .Select(h => h.StartsWith("=\"") && h.EndsWith("\"") ? h.Substring(2, h.Length - 3) : h)
                                              .ToArray();

                PopulateDataTableWithColumnHeaders(dataTable, headers);

                // Read all lines into memory
                string line;
                while ((line = await sr.ReadLineAsync()) != null)
                {
                    string[] fields = line.Split(delimiter).Select(p => string.IsNullOrWhiteSpace(p) ? "" : p).ToArray();
                    allData.Add(fields);
                }
            }

            int expectedColCount = dataTable.Columns.Count;

            // Filter allData to keep only rows that match the expected column count
            allData = allData.Where(row => row.Length == expectedColCount).ToList();
            if (allData.Count == 0)
                return dataTable;

            var quotedColumnsToClean = new List<int>();
            if (allData.Count > 0)
            {
                int rowCount = allData.Count;

                for (int col = 0; col < expectedColCount; col++)
                {
                    string firstVal = allData[0][col]?.ToString() ?? "";
                    string lastVal = allData[rowCount - 1][col]?.ToString() ?? "";

                    // Check first or last row for the pattern to avoid unnecessary loops
                    if ((firstVal.StartsWith("=\"") && firstVal.EndsWith("\"")) ||
                        (lastVal.StartsWith("=\"") && lastVal.EndsWith("\"")))
                    {
                        // Verify if all non-null cells in this column start with ="
                        bool isConsistent = allData.All(row =>
                        {
                            string val = row[col]?.ToString();
                            return string.IsNullOrEmpty(val) || val.StartsWith("=\"");
                        });

                        if (isConsistent)
                        {
                            quotedColumnsToClean.Add(col);
                        }
                    }
                }

                // Strip the quotes only from the verified columns
                if (quotedColumnsToClean.Any())
                {
                    allData.AsParallel().ForAll(row =>
                    {
                        foreach (int col in quotedColumnsToClean)
                        {
                            string val = row[col]?.ToString();
                            if (!string.IsNullOrEmpty(val) && val.StartsWith("=\"") && val.EndsWith("\""))
                            {
                                row[col] = val.Substring(2, val.Length - 3);
                            }
                        }
                    });
                }
            }

            Dictionary<int, Type> columnTypes = new Dictionary<int, Type>();
            // Determine the most appropriate data type for each column
            columnTypes = Enumerable.Range(0, dataTable.Columns.Count).AsParallel().WithDegreeOfParallelism(dataTable.Columns.Count).Select(colIndex =>
            {
                if (quotedColumnsToClean.Contains(colIndex))
                    return new { ColIndex = colIndex, Type = typeof(string) };

                bool allNumbers = allData.All(row => !(row[colIndex].ToString().Length > 1 && row[colIndex].ToString().StartsWith("0") && !row[colIndex].ToString().Contains(".")) && double.TryParse(row[colIndex].ToString(), out _));
                if (allNumbers)
                    return new { ColIndex = colIndex, Type = typeof(double) };
                else
                    return new { ColIndex = colIndex, Type = typeof(string) };
            }).ToDictionary(col => col.ColIndex, col => col.Type);

            // Set the data types for the columns
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                string cleanedName = Regex.Replace(dataTable.Columns[i].ColumnName, @"[^a-zA-Z]+$", string.Empty).ToUpper();
                if (
                        !(
                            cleanedName.EndsWith(" ID", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_ID", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_CD", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_CODE", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith(" CODE", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_KY", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_KEY", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith(" KEY", StringComparison.OrdinalIgnoreCase)
                        ) &&
                        columnTypes.TryGetValue(i, out Type dataType)
                   )
                    dataTable.Columns[i].DataType = dataType;
            }

            // Populate the DataTable with data
            foreach (var rowData in allData)
            {
                dataTable.Rows.Add(rowData);
            }

            return dataTable;
        }

        public static async Task<DataTable> ReadFileIntoDataTable(string filePath, char delimiter)
        {
            DataTable dataTable = new DataTable();
            List<object[]> allData = new List<object[]>();

            using (StreamReader sr = new StreamReader(filePath, true))
            {
                // Read the first line to create column headers
                string headersLine = await sr.ReadLineAsync();
                if (headersLine == null) return dataTable; // Empty file, return empty DataTable

                string[] headers = headersLine.Split(delimiter)
                                              .Select(h => h.StartsWith("=\"") && h.EndsWith("\"") ? h.Substring(2, h.Length - 3) : h)
                                              .ToArray();

                PopulateDataTableWithColumnHeaders(dataTable, headers);

                // Read all lines into memory
                string line;
                while ((line = await sr.ReadLineAsync()) != null)
                {
                    string[] fields = line.Split(delimiter).Select(p => string.IsNullOrWhiteSpace(p) ? "" : p).ToArray();
                    allData.Add(fields);
                }
            }

            int expectedColCount = dataTable.Columns.Count;

            // Find rows where the column count does not match the header count
            var badRows = allData
                .Select((row, index) => new { RowLength = row.Length, LineNumber = index + 2 }) // +2 accounts for 0-based index and the Header line
                .Where(x => x.RowLength != expectedColCount)
                .Select(x => x.LineNumber)
                .ToList();

            if (badRows.Any())
            {
                throw new InvalidDataException($"Row lengths do not match on line(s): {string.Join(", ", badRows)}. Expected {expectedColCount} columns but found mismatches.");
            }

            var quotedColumnsToClean = new List<int>();
            if (allData.Count > 0)
            {
                int rowCount = allData.Count;

                for (int col = 0; col < expectedColCount; col++)
                {
                    string firstVal = allData[0][col]?.ToString() ?? "";
                    string lastVal = allData[rowCount - 1][col]?.ToString() ?? "";

                    // Check first or last row for the pattern to avoid unnecessary loops
                    if ((firstVal.StartsWith("=\"") && firstVal.EndsWith("\"")) ||
                        (lastVal.StartsWith("=\"") && lastVal.EndsWith("\"")))
                    {
                        // Verify if all non-null cells in this column start with ="
                        bool isConsistent = allData.All(row =>
                        {
                            string val = row[col]?.ToString();
                            return string.IsNullOrEmpty(val) || val.StartsWith("=\"");
                        });

                        if (isConsistent)
                        {
                            quotedColumnsToClean.Add(col);
                        }
                    }
                }

                // Strip the quotes only from the verified columns
                if (quotedColumnsToClean.Any())
                {
                    allData.AsParallel().ForAll(row =>
                    {
                        foreach (int col in quotedColumnsToClean)
                        {
                            string val = row[col]?.ToString();
                            if (!string.IsNullOrEmpty(val) && val.StartsWith("=\"") && val.EndsWith("\""))
                            {
                                row[col] = val.Substring(2, val.Length - 3);
                            }
                        }
                    });
                }
            }

            Dictionary<int, Type> columnTypes = new Dictionary<int, Type>();
            // Determine the most appropriate data type for each column
            columnTypes = Enumerable.Range(0, dataTable.Columns.Count).AsParallel().WithDegreeOfParallelism(dataTable.Columns.Count).Select(colIndex =>
            {
                if (quotedColumnsToClean.Contains(colIndex))
                    return new { ColIndex = colIndex, Type = typeof(string) };

                bool allNumbers = allData.All(row => !(row[colIndex].ToString().Length > 1 && row[colIndex].ToString().StartsWith("0") && !row[colIndex].ToString().Contains(".")) && double.TryParse(row[colIndex].ToString(), out _));
                if (allNumbers)
                    return new { ColIndex = colIndex, Type = typeof(double) };
                else
                    return new { ColIndex = colIndex, Type = typeof(string) };
            }).ToDictionary(col => col.ColIndex, col => col.Type);

            // Set the data types for the columns
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                string cleanedName = Regex.Replace(dataTable.Columns[i].ColumnName, @"[^a-zA-Z]+$", string.Empty).ToUpper();
                if (
                        !(
                            cleanedName.EndsWith(" ID", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_ID", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_CD", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_CODE", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith(" CODE", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_KY", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("_KEY", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith(" KEY", StringComparison.OrdinalIgnoreCase)
                        ) &&
                        columnTypes.TryGetValue(i, out Type dataType)
                   )
                    dataTable.Columns[i].DataType = dataType;
            }

            // Populate the DataTable with data
            foreach (var rowData in allData)
            {
                dataTable.Rows.Add(rowData);
            }

            return dataTable;
        }

        private static void PopulateDataTableWithColumnHeaders(DataTable dataTable, string[] headers)
        {
            Dictionary<string, int> headerCounts = new Dictionary<string, int>();

            foreach (string header in headers)
            {
                string originalHeader = header;
                string colName = originalHeader;
                if (headerCounts.TryGetValue(originalHeader, out int count))
                {
                    count++;
                    colName = $"{originalHeader}{{[{count}]}}";
                    headerCounts[originalHeader] = count;
                }
                else
                {
                    headerCounts[originalHeader] = 1;
                }
                dataTable.Columns.Add(colName, typeof(string)); // Initially set as string
            }
        }

        public static DataTable ReadFileIntoDataTableTest(string filePath, char delimiter)
        {
            DataTable dataTable = new DataTable();
            List<object[]> allData = new List<object[]>();

            using (StreamReader sr = new StreamReader(filePath, true))
            {
                // Read the first line to create column headers
                string[] headers = sr.ReadLine().Split(delimiter);
                foreach (string header in headers)
                {
                    dataTable.Columns.Add(header, typeof(string)); // Initially set as string
                }

                // Read all lines into memory
                while (!sr.EndOfStream)
                {
                    string[] fields = sr.ReadLine().Split(delimiter).Select(p =>
                    {
                        if (string.IsNullOrWhiteSpace(p))
                            return string.Empty;

                        // Check if the value is quoted
                        if (p.StartsWith("\"") && p.EndsWith("\""))
                        {
                            string innerValue = p.Substring(1, p.Length - 2);

                            // Check if all quotes inside are doubled
                            if (!innerValue.Replace("\"\"", "").Contains("\""))
                            {
                                // Replace doubled quotes with single quotes
                                return innerValue.Replace("\"\"", "\"");
                            }
                            else
                                return innerValue;
                        }
                        else
                            return p; // Return unmodified if not wrapped in quotes
                    }).ToArray();

                    allData.Add(fields);
                }

            }

            Dictionary<int, Type> columnTypes = new Dictionary<int, Type>();
            // Determine the most appropriate data type for each column
            columnTypes = Enumerable.Range(0, dataTable.Columns.Count).AsParallel().WithDegreeOfParallelism(dataTable.Columns.Count).Select(colIndex =>
            {
                bool allNumbers = allData.All(row => !(row[colIndex].ToString().Length > 1 && row[colIndex].ToString().StartsWith("0") && !row[colIndex].ToString().Contains(".")) && double.TryParse(row[colIndex].ToString(), out _));
                if (allNumbers)
                    return new { ColIndex = colIndex, Type = typeof(double) };
                else
                    return new { ColIndex = colIndex, Type = typeof(string) };
            }).ToDictionary(col => col.ColIndex, col => col.Type);


            // Set the data types for the columns
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                string cleanedName = Regex.Replace(dataTable.Columns[i].ColumnName, @"[^a-zA-Z]+$", string.Empty);
                if (
                        !(
                            cleanedName.EndsWith("ID", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("CODE", StringComparison.OrdinalIgnoreCase) ||
                            cleanedName.EndsWith("KEY", StringComparison.OrdinalIgnoreCase)
                        ) &&
                        columnTypes.TryGetValue(i, out Type dataType)
                   )
                    dataTable.Columns[i].DataType = dataType;
            }

            // Populate the DataTable with data
            foreach (var rowData in allData)
            {
                dataTable.Rows.Add(rowData);
            }

            return dataTable;
        }

        public static async Task<DataTable> ReadFileIntoDataTableAdv(string filePath, string delimiter)
        {
            DataTable dataTable = new DataTable();
            List<object[]> allData = new List<object[]>();

            using (StreamReader sr = new StreamReader(filePath, true))
            {
                string headersLine = await sr.ReadLineAsync();
                if (headersLine == null) return dataTable; // Empty file, return empty DataTable

                List<string> headers = new List<string>();

                // Regex to match quoted and unquoted fields
                var matches = Regex.Matches(headersLine, $@"""([^""]*(?:""""[^""]*)*)""|([^{Regex.Escape(delimiter)}""]+)|(?<={Regex.Escape(delimiter)})$|(?<={Regex.Escape(delimiter)})(?={Regex.Escape(delimiter)})|^{Regex.Escape(delimiter)}|{Regex.Escape(delimiter)}$");

                foreach (Match match in matches)
                {
                    string header = match.Value;

                    if (header.StartsWith("\"") && header.EndsWith("\""))
                    {
                        // Remove surrounding quotes and replace doubled quotes
                        header = header.Substring(1, header.Length - 2).Replace("\"\"", "\"");
                    }

                    headers.Add(header);
                }

                Dictionary<string, int> headerCounts = new Dictionary<string, int>();
                PopulateDataTableWithColumnHeaders(dataTable, headers.ToArray());

                string line;
                while ((line = await sr.ReadLineAsync()) != null)
                {
                    List<string> fields = new List<string>();

                    matches = Regex.Matches(line, $@"""([^""]*(?:""""[^""]*)*)""|([^{Regex.Escape(delimiter)}""]+)|(?<={Regex.Escape(delimiter)})$|(?<={Regex.Escape(delimiter)})(?={Regex.Escape(delimiter)})");

                    foreach (Match match in matches)
                    {
                        string field;

                        if (match.Groups[1].Success)
                        {
                            // Quoted field, remove surrounding quotes and replace doubled quotes
                            field = match.Groups[1].Value.Replace("\"\"", "\"");
                        }
                        else if (match.Groups[2].Success)
                        {
                            // Unquoted field
                            field = match.Groups[2].Value;
                        }
                        else
                        {
                            // Empty field
                            field = string.Empty;
                        }

                        fields.Add(field);
                    }

                    // Add the parsed fields to allData
                    allData.Add(fields.ToArray());
                }
            }

            int expectedColCount = dataTable.Columns.Count;

            // Find rows where the column count does not match the header count
            var badRows = allData
                .Select((row, index) => new { RowLength = row.Length, LineNumber = index + 2 }) // +2 accounts for 0-based index and the Header line
                .Where(x => x.RowLength != expectedColCount)
                .Select(x => x.LineNumber)
                .ToList();

            if (badRows.Any())
            {
                throw new InvalidDataException($"Row lengths do not match on line(s): {string.Join(", ", badRows)}. Expected {expectedColCount} columns but found mismatches.");
            }

            Dictionary<int, Type> columnTypes = new Dictionary<int, Type>();
            // Determine the most appropriate data type for each column
            columnTypes = Enumerable.Range(0, dataTable.Columns.Count).AsParallel().WithDegreeOfParallelism(dataTable.Columns.Count).Select(colIndex =>
            {
                bool allNumbers = allData.All(row => !(row[colIndex].ToString().Length > 1 && row[colIndex].ToString().StartsWith("0") && !row[colIndex].ToString().Contains(".")) && double.TryParse(row[colIndex].ToString(), out _));
                if (allNumbers)
                    return new { ColIndex = colIndex, Type = typeof(double) };
                else
                    return new { ColIndex = colIndex, Type = typeof(string) };
            }).ToDictionary(col => col.ColIndex, col => col.Type);


            // Set the data types for the columns
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                string colNameUpper = dataTable.Columns[i].ColumnName.ToUpper();
                if (
                        !(
                            colNameUpper.EndsWith(" ID", StringComparison.OrdinalIgnoreCase) ||
                            colNameUpper.EndsWith("_ID", StringComparison.OrdinalIgnoreCase) ||
                            colNameUpper.EndsWith("_CD", StringComparison.OrdinalIgnoreCase) ||
                            colNameUpper.EndsWith("_CODE", StringComparison.OrdinalIgnoreCase) ||
                            colNameUpper.EndsWith(" CODE", StringComparison.OrdinalIgnoreCase) ||
                            colNameUpper.EndsWith("_KY", StringComparison.OrdinalIgnoreCase) ||
                            colNameUpper.EndsWith("_KEY", StringComparison.OrdinalIgnoreCase) ||
                            colNameUpper.EndsWith(" KEY", StringComparison.OrdinalIgnoreCase)
                        ) &&
                        columnTypes.TryGetValue(i, out Type dataType)
                   )
                    dataTable.Columns[i].DataType = dataType;
            }

            // Populate the DataTable with data
            foreach (var rowData in allData)
            {
                dataTable.Rows.Add(rowData);
            }

            return dataTable;
        }

        public static string StripDuplicatedColumnSuffix(string columnName)
        {
            // Remove trailing {n} where n is digits
            return Regex.Replace(columnName, @"\{\[\d+\]\}$", string.Empty);
        }

        /*public static Excel.Range CreateExcelWorkbookFromTextFile(Excel.Worksheet worksheet, string filePath, string delimiter = "\t")
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found at path: {filePath}");
            }

            using (StreamReader reader = new StreamReader(filePath))
            {
                List<List<string>> rows = new List<List<string>>();

                string line;

                while ((line = reader.ReadLine()) != null)
                {
                    string[] values = line.Split(new string[] { delimiter }, StringSplitOptions.None);

                    rows.Add(values.ToList());
                }

                int rowCount = rows.Count;
                int columnCount = rows.Max(row => row.Count);

                Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, columnCount]];
                range.Value2 = rows.Select(row => row.ToArray()).ToArray();
                range.NumberFormat = "@";

                return range;
            }
        }*/
    }
}


