using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportTableToExcel
{
    public static class WorksheetFromTxtCreator
    {
        public static void CreateExcelWorkbookFromTextFileQueryTable(Excel.Worksheet worksheet, string filePath, char delimiter)
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

        public static void ImportTextFileToExcel(Excel.Worksheet worksheet, string filePath, char delimiter)
        {
            Task<DataTable> taskReadFile = Task.Run(() => ReadFileIntoDataTable(filePath, delimiter));
            taskReadFile.ContinueWith(t =>
            {
                DataTable dataTable = t.Result;
                UtilsExcel.PasteDataTableToRange(dataTable, worksheet.Cells[1, 1]);
            });
        }


        public static DataTable ReadFileIntoDataTable(string filePath, char delimiter)
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
                    string[] fields = sr.ReadLine().Split(delimiter).Select(p => string.IsNullOrWhiteSpace(p) ? "" : p).ToArray();
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
                if (
                        !(
                            dataTable.Columns[i].ColumnName.EndsWith("ID", StringComparison.OrdinalIgnoreCase) || 
                            dataTable.Columns[i].ColumnName.EndsWith("CODE", StringComparison.OrdinalIgnoreCase) || 
                            dataTable.Columns[i].ColumnName.EndsWith("KEY", StringComparison.OrdinalIgnoreCase)
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


