using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAddInByMarcinOlszewski;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data;
using Microsoft.SqlServer.TransactSql.ScriptDom;
using System.Security.Cryptography;
using ColorMine.ColorSpaces;
using System.Drawing;
using System.Text;
using Microsoft.Data.Sqlite;

public static class Utils
{
    public static List<string> TextExt = new List<string> { ".txt", ".csv" };
    public static List<string> ExcelExt = new List<string> { ".xlsx", ".xlsb", ".xlsm", ".xltm", ".xls", ".xlt" };
    public static Random rand = new Random();

    public static char DetermineTableDelimiter(string filePath)
    {
        char[] delimiters = { '\t', ',', ';', '|' }; // set the candidate delimiters

        using (StreamReader reader = new StreamReader(filePath))
        {
            char[] foundDelimiters = new char[10];
            for (int i = 0; i < 10; i++) // read the first 10 lines of the file
            {
                string line = reader.ReadLine();
                if (line != null)
                {
                    char delimiter = delimiters.FirstOrDefault(d => line.Contains(d));
                    if (delimiter != default(char))
                    {
                        foundDelimiters[i] = delimiter; // return the first delimiter found
                    }
                }
                else
                    break;
            }
            if (foundDelimiters.Where(p => p != default(char)).ToArray().Length > 0 && foundDelimiters.Where(p => p != default(char)).All(p => p.Equals(foundDelimiters[0])))
                return foundDelimiters[0];
        }
        return default(char); // no delimiter found
    }

    

    public static bool IsSQLQueryValid(string sql, out List<string> errors)
    {
        errors = new List<string>();

        TSql140Parser parser = new TSql140Parser(false);
        TSqlFragment fragment;
        IList<ParseError> parseErrors;

        using (TextReader reader = new StringReader(sql))
        {
            fragment = parser.Parse(reader, out parseErrors);
            if (parseErrors != null && parseErrors.Count > 0)
            {
                errors = parseErrors.Select(e => e.Message).ToList();
                return false;
            }
        }
        return true;
    }

    public static T Clamp<T>(this T val, T min, T max) where T : IComparable<T>
    {
        if (val.CompareTo(min) < 0) return min;
        else if (val.CompareTo(max) > 0) return max;
        else return val;
    }

    public static string GetUniqueString(List<string> existingStrings, string baseString)
    {
        baseString = baseString.Substring(0, Math.Min(31, baseString.Length));
        string newString = baseString;
        int i = 1;

        while (existingStrings.Contains(newString))
        {
            string suffix = $" ({i})";
            int baseStringLength = 31 - suffix.Length;

            // Cut characters from the base string to make room for the suffix
            newString = baseString.Substring(0, Math.Min(baseStringLength, baseString.Length)) + suffix;
            i++;
        }

        return newString;
    }

    public static Dictionary<string, int> GetCounts(DataTable dt, string searchWord = "")
    {
        var counts = new Dictionary<string, int>();

        foreach (DataColumn column in dt.Columns)
        {
            int count = 0;
            foreach (DataRow row in dt.Rows)
            {
                var cellValue = row[column].ToString();
                if (string.IsNullOrEmpty(searchWord))
                {
                    if (!string.IsNullOrEmpty(cellValue))
                        count++;
                }
                else
                {
                    if (cellValue.Contains(searchWord))
                        count++;
                }
            }
            counts[column.ColumnName] = count;
        }

        return counts;
    }

    public static void SaveDataTableToTxt(DataTable dt, string filePath = "", string initialName = "", string initialDirectory = "")
    {
        if (string.IsNullOrWhiteSpace(filePath))
            filePath = FileManager.GetPathByDialog(initialName, initialDirectory);

        if (string.IsNullOrWhiteSpace(filePath))
            return;

        StringBuilder sb = new StringBuilder();

        // Add column headers
        string[] columnNames = dt.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray();
        sb.AppendLine(string.Join(",", columnNames));

        // Add rows
        foreach (DataRow row in dt.Rows)
        {
            string[] fields = row.ItemArray.Select(field => field.ToString()).ToArray();
            sb.AppendLine(string.Join(",", fields));
        }

        // Write to file
        File.WriteAllText(filePath, sb.ToString());
    }

    public static void SuperShuffle<T>(this IList<T> list)
    {
        RNGCryptoServiceProvider provider = new RNGCryptoServiceProvider();
        int n = list.Count;
        while (n > 1)
        {
            byte[] box = new byte[1];
            do provider.GetBytes(box);
            while (!(box[0] < n * (Byte.MaxValue / n)));
            int k = (box[0] % n);
            n--;
            T value = list[k];
            list[k] = list[n];
            list[n] = value;
        }
    }


    public static void Shuffle<T>(this IList<T> list)
    {
        int n = list.Count;
        while (n > 1)
        {
            n--;
            int k = rand.Next(n + 1);
            T value = list[k];
            list[k] = list[n];
            list[n] = value;
        }
    }

    public static List<Color> GenerateColorPalette(int numberOfColors)
    {
        List<Color> colorPalette = new List<Color>();
        Random rand = new Random();
        double hue = 0.0;
        double saturation = 0.5; // Set saturation to middle
        double lightness = 0.85;  // Set lightness to middle
        double saturationIncrement = 0.1;
        double lightnessIncrement = 0.1;


        int numberOfColorsLeft = numberOfColors;
        while (numberOfColorsLeft > 0)
        {
            hue = 160.0;
            double hueIncrement = Math.Min(350.0 / Math.Min(Math.Max(numberOfColorsLeft, 0), 350), 50);
            for (int i = 0; i <= Math.Min(numberOfColorsLeft - 1, 350); i++, hue += hueIncrement)
            {
                colorPalette.Add(HSLToRGB(hue % 350, saturation, lightness));
            }
            numberOfColorsLeft = numberOfColors - colorPalette.Count;
            if (numberOfColorsLeft > 0 && lightness > 0.66)
            {
                lightness -= lightnessIncrement;
                hue = 0;
            }
            else if (numberOfColorsLeft > 0 && saturation < 0.99)
            {
                saturation += saturationIncrement;
                hue = 0;
            }
            else
                for (int i = 0; i < numberOfColorsLeft; i++)
                    colorPalette.Add((Color)colorPalette[i % colorPalette.Count]);
        }
        colorPalette = colorPalette.OrderBy(c => c.GetHue()).ToList();
        return colorPalette.Take(numberOfColors).ToList();
    }

    public static Color HSLToRGB(double h, double s, double l)
    {
        var hsl = new Hsl { H = h, S = s, L = l };
        var rgb = hsl.To<Rgb>();
        return rgb.ToSystemColor();
    }

    public static Color DarkenColor(this Color color, float howMuchDarker01)
    {
        int r = (int)(color.R * (1 - howMuchDarker01));
        int g = (int)(color.G * (1 - howMuchDarker01));
        int b = (int)(color.B * (1 - howMuchDarker01));
        return Color.FromArgb(color.A, r, g, b);
    }

    public static void MoveFormToCursor(Form form)
    {
        // Set the form's location to the cursor's position
        form.Location = Cursor.Position;

        // Get the working area of the screen that contains the form
        Rectangle screenWorkingArea = Screen.GetWorkingArea(form);

        // Check if the form is completely visible in the screen's working area
        if (!screenWorkingArea.Contains(form.Bounds))
        {
            // If not, adjust the form's location

            // If the form's right edge is out of the screen, move it to the left
            if (form.Right > screenWorkingArea.Right)
            {
                form.Left = screenWorkingArea.Right - form.Width;
            }

            // If the form's bottom edge is out of the screen, move it up
            if (form.Bottom > screenWorkingArea.Bottom)
            {
                form.Top = screenWorkingArea.Bottom - form.Height;
            }

            // If the form's left edge is out of the screen, move it to the right
            if (form.Left < screenWorkingArea.Left)
            {
                form.Left = screenWorkingArea.Left;
            }

            // If the form's top edge is out of the screen, move it down
            if (form.Top < screenWorkingArea.Top)
            {
                form.Top = screenWorkingArea.Top;
            }
        }
    }
    

    public static void ExecuteSqlQueryAndDisplayResults(DataGridView dataGridView, string sqlQuery)
    {
        try
        {
            // Convert DataGridView data to DataTable
            DataTable dt = (DataTable)dataGridView.DataSource;
            SQLitePCL.raw.SetProvider(new SQLitePCL.SQLite3Provider_e_sqlite3());
            using (SqliteConnection conn = new SqliteConnection("Data Source=:memory:"))
            {
                conn.Open();

                // Create table schema in SQLite in-memory database
                string createTableQuery = string.Join(", ", dt.Columns.Cast<DataColumn>().Select(c => $"[{c.ColumnName}] {GetSQLiteDataType(c.DataType)}"));
                string createTableCommand = $"CREATE TABLE DataTable ({createTableQuery})";
                using (SqliteCommand cmd = new SqliteCommand(createTableCommand, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                // Insert data from DataTable into SQLite in-memory database
                using (SqliteTransaction transaction = conn.BeginTransaction())
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        string insertValues = string.Join(", ", row.ItemArray.Select(v => $"'{v}'"));
                        string insertCommand = $"INSERT INTO DataTable VALUES ({insertValues})";
                        using (SqliteCommand cmd = new SqliteCommand(insertCommand, conn, transaction))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }

                // Execute SQL query and load results back into DataTable
                dt.Clear();
                using (SqliteCommand cmd = new SqliteCommand(sqlQuery, conn))
                {
                    SqliteDataReader reader = cmd.ExecuteReader();
                    dt.Load(reader);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }
    }

    public static string GetSQLiteDataType(Type type)
    {
        if (type == typeof(Int32))
        {
            return "INT";
        }
        else if (type == typeof(string))
        {
            return "TEXT";
        }
        else if (type == typeof(Double))
        {
            return "REAL";
        }
        else if (type == typeof(Boolean))
        {
            return "INTEGER";
        }
        else
        {
            return "BLOB";
        }
    }

    public static Color Invert(this Color color)
    {
        return Color.FromArgb(255 - color.R, 255 - color.G, 255 - color.B);
    }

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

}

