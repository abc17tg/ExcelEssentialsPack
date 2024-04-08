using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInByMarcinOlszewski.Forms
{
    public partial class DataTableForm : Form
    {
        public DataTable DataTable;
        public string Query;
        Excel.Application ExcelApp;

        public DataTableForm(DataTable dataTable, string query, Excel.Application app)
        {
            InitializeComponent();
            DataTable = dataTable;
            Query = query;
            ExcelApp = app;
            queryRichTextBox.Text = query;
            dataGridView.AutoGenerateColumns = true;
            dataGridView.DataSource = DataTable;
            dataGridView.RowPostPaint += dataGridView_RowPostPaint;
            var nfi = new CultureInfo("en-US", false).NumberFormat;
            nfi.NumberGroupSeparator = " ";
            dataTableDimentionsLabel.Text = $"Rows: {(DataTable.Rows.Count + (headersCheckBox.Checked ? 1 : 0)).ToString("N0", nfi)}\nColumns: {DataTable.Columns.Count.ToString("N0", nfi)}";
        }

        private void dataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void pasteButton_Click(object sender, EventArgs e)
        {
            Paste();
        }

        public void Paste()
        {
            Excel.Range rng = ExcelApp.ActiveWindow.RangeSelection;
            if (rng.Valid())
            {
                UtilsExcel.PasteDataTableToRange(DataTable, rng, headersCheckBox.Checked);
            }
            else
            {
                var result = MessageBox.Show("No selection to paste", "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                if (result == DialogResult.Retry)
                    Paste();
            }
        }

        private void queryRichTextBox_DoubleClick(object sender, EventArgs e)
        {
            queryRichTextBox.SendToBack();
            queryRichTextBox.Visible = false;
        }

        private void queryLabel_Click(object sender, EventArgs e)
        {
            queryRichTextBox.BringToFront();
            queryRichTextBox.Visible = true;
        }

        private void DataTableForm_Load(object sender, EventArgs e)
        {
            /*dataGridView.DataSource = DataTable;
            dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.ColumnHeader);*/
        }
    }
}
