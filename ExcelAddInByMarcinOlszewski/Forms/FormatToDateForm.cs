using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelEssentials.Scripts;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEssentials.Forms
{
    public partial class FormatToDateForm : Form
    {
        private Excel.Application m_app;
        private Excel.Range m_rng;
        private DataTable m_dataTable;

        public FormatToDateForm(Excel.Application app)
        {
            InitializeComponent();
            m_app = app;
            predefinedFormatsComboBox.Items.AddRange(DateFormatDetector.DateFormats.ToArray());
            checkButton.Enabled = Fetch();
        }

        private void fetchButton_Click(object sender, EventArgs e)
        {
            checkButton.Enabled = Fetch();
        }

        private bool Fetch()
        {
            m_rng = m_app.ActiveWindow.RangeSelection.GetUsableRange();

            if (!m_rng.Valid() || m_rng.Columns.Count != 1)
            {
                m_rng = null;
                m_dataTable = null;
                rngDataGridView.Rows.Clear();
                countLabel.Text = $"Cells count: NaN";
                return false;
            }

            m_dataTable = m_rng.GetDataTable2(false);
            if (m_dataTable != null && m_dataTable.Rows.Count > 0)
            {
                m_dataTable.Columns[0].ColumnName = "Input value";
                m_dataTable.Columns.Add("Formatted date", typeof(DateTime));
                rngDataGridView.DataSource = m_dataTable;
                countLabel.Text = $"Cells count: {m_dataTable.Rows.Count}";
            }
            else
            {
                rngDataGridView.DataSource = null;
                countLabel.Text = $"Cells count: NaN";
                return false;
            }

            return true;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            if (IsParsed())
            {
                m_dataTable.Columns.Remove(m_dataTable.Columns[0]);
                UtilsExcel.PasteDataTableToRange(m_dataTable, m_app.ActiveWindow.RangeSelection, false);
            }

            Close();
        }

        private bool IsParsed()
        {
            foreach (DataRow row in m_dataTable.Rows)
            {
                // Check if the second column contains a non-DBNull value
                if (row[1] != DBNull.Value)
                {
                    return true; // Return true if any value is not DBNull
                }
            }
            return false; // Return false if all values are DBNull
        }

        private void checkButton_Click(object sender, EventArgs e)
        {
            ParseDates();
        }

        private void ParseDates()
        {
            string format = formatTextBox?.Text; // Get the format from the TextBox

            if (string.IsNullOrWhiteSpace(format))
            {
                string detectedFormat = DateFormatDetector.DetectDateFormat(m_dataTable.Rows.Cast<DataRow>().FirstOrDefault(row => row[0] != DBNull.Value)?[0]?.ToString());
                if (detectedFormat != null)
                {
                    predefinedFormatsComboBox.SelectedText = detectedFormat;
                    formatTextBox.Text = detectedFormat;
                    format = detectedFormat;
                }
                else
                {
                    MessageBox.Show("Please enter a valid date format.");
                    parsedDateLabel.Text = string.Empty;
                    return;
                }

            }

            // Loop through each row to parse the dates
            foreach (DataRow row in m_dataTable.Rows)
            {
                string inputValue = row[0]?.ToString();

                if (string.IsNullOrWhiteSpace(inputValue))
                {
                    row[1] = DBNull.Value; // Set to DBNull if input is empty or null
                    continue;
                }

                try
                {
                    // Try to parse the date with the given format
                    DateTime parsedDate = DateTime.ParseExact(inputValue, format, System.Globalization.CultureInfo.InvariantCulture);
                    row[1] = parsedDate; // Assign the parsed date to the second column
                }
                catch (FormatException)
                {
                    // If parsing fails, reset the column and show a message
                    foreach (DataRow resetRow in m_dataTable.Rows)
                    {
                        resetRow[1] = DBNull.Value; // Set entire column to DBNull
                    }

                    parsedDateLabel.Text = string.Empty;
                    MessageBox.Show($"Parsing failed for value: '{inputValue}'. Please check the format and try again.", "Parsing Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return; // Exit function on failure
                }
            }

            parsedDateLabel.Text = $"Current date: {DateTime.Now.ToString(format)}\n{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)}";
            // Update the DataGridView to reflect changes in the DataTable
            rngDataGridView.DataSource = m_dataTable;
        }

        private void inputFormatButton_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = Path.Combine(FileManager.ResourcesPath, "DateTimeFormatCheatSheet.txt");
                DataTable dt = Utils.ReadTabDelimitedFile(filePath);
                DataTableForm dataTableForm = new DataTableForm(dt, "No query", m_app);
                dataTableForm.Show();
            }
            catch (Exception) { }
        }

        private void predefinedFormatsComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            formatTextBox.Text = predefinedFormatsComboBox?.SelectedText ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(formatTextBox.Text))
                ParseDates();
        }

        private void formatTextBox_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
