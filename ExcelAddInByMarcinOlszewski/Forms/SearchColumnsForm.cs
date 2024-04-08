using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInByMarcinOlszewski.Forms
{
    public partial class SearchColumnsForm : Form
    {
        private Excel.Application m_app;
        private Excel.Worksheet m_worksheet;
        public Excel.Range Rng;
        public DataTable DataTable;

        public SearchColumnsForm(Excel.Application app, string title = null)
        {
            InitializeComponent();
            m_app = app;
            m_worksheet = app.ActiveSheet;
            if (!string.IsNullOrEmpty(title))
                this.Text = title;
            m_app.SheetActivate += SelectionChanged;
            SelectionChanged(m_worksheet);
        }

        public void SelectionChanged(object sender)
        {
            if (m_worksheet != null && m_worksheet == (Excel.Worksheet)sender)
                return;

            m_worksheet = (Excel.Worksheet)sender;
            Rng = m_worksheet.UsedRange;
            DataTable = Rng.GetDataTable();
            var dataDic = Utils.GetCounts(DataTable, searchTextBox.Text);
            // Clear the DataGridView
            dataGridView.Rows.Clear();
            var cb = new CheckBox();
            cb.CheckState = CheckState.Unchecked;
            cb.Text = "";

            // Populate the DataGridView with dictionary data
            foreach (var item in dataDic)
            {
                dataGridView.Rows.Add(item.Key, item.Value.ToString(), cb, new Button());
            }

        }

        private void searchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (searchContentsCheckBox.Checked)
                return;

            var dataDic = Utils.GetCounts(DataTable, searchTextBox.Text);
            // Clear the DataGridView
            dataGridView.Rows.Clear();
            var cb = new CheckBox();
            cb.CheckState = CheckState.Unchecked;
            cb.Text = "";

            // Populate the DataGridView with dictionary data
            foreach (var item in dataDic)
            {
                dataGridView.Rows.Add(item.Key, item.Value.ToString(), cb, new Button());
            }
        }

        private void ColumnSelect(object sender, EventArgs e)
        {

        }

        private void okBtn_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            m_worksheet.Activate();
            Close();
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            m_app.SheetActivate -= SelectionChanged;
            DialogResult = DialogResult.Cancel;
            m_worksheet.Activate();
            Close();
        }

        ~SearchColumnsForm()
        {
            m_app.SheetActivate -= SelectionChanged;
            Close();
        }

        private void searchBtn_Click(object sender, EventArgs e)
        {
            var dataDic = Utils.GetCounts(DataTable, searchTextBox.Text);
            // Clear the DataGridView
            dataGridView.Rows.Clear();
            var cb = new CheckBox();
            cb.CheckState = CheckState.Unchecked;
            cb.Text = "";

            // Populate the DataGridView with dictionary data
            foreach (var item in dataDic)
            {
                dataGridView.Rows.Add(item.Key, item.Value.ToString(), cb, new Button());
            }
        }

        private void searchContentsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            searchBtn.Enabled = searchContentsCheckBox.Checked;
        }
    }
}
