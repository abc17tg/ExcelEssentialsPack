using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelEssentials.Scripts;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEssentials.Forms
{
    public partial class SearchColumnsForm : Form
    {
        private Excel.Application m_app;
        private Excel.Range m_rng;
        private DataTable m_dataTable;
        private bool m_isPivotTable = false;
        private bool UseDataTable => useDataTableCheckBox.Checked;

        public SearchColumnsForm(Excel.Application app, string title = null)
        {
            InitializeComponent();
            m_app = app;
            if (!string.IsNullOrWhiteSpace(title))
                this.Text = title;
            dataGridView.Columns[1].Visible = countsCheckBox.Checked;
            this.Cursor = Cursors.WaitCursor;
            Fetch();
            this.Cursor = Cursors.Default;
            useDataTableToolTip.SetToolTip(useDataTableCheckBox, "Fetch much longer but search faster. (after checking fetch again)");
            useDataTableToolTip.SetToolTip(searchContentsCheckBox, "Search contents of table not headers)");
            useDataTableToolTip.SetToolTip(countsCheckBox, "Count elements in column");
        }

        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (e.ColumnIndex == 0 && e.RowIndex >= 0)
            {
                ColumnSelect(dataGridView[e.ColumnIndex, e.RowIndex].Value.ToString());
            }

            if (e.ColumnIndex == 3 && e.RowIndex >= 0)
            {
                Excel.Range headerCell = ((Excel.Range)m_rng.Rows[1]).Cells.Cast<Excel.Range>().FirstOrDefault(p => dataGridView[0, e.RowIndex].Value.ToString() == (p.Value2?.ToString() ?? $"Column{p.Column - m_rng.Column + 1}"));
                ClearCells(headerCell.Column - m_rng.Column + 1);
            }
            this.Cursor = Cursors.Default;
        }


        private void Form_Load(object sender, EventArgs e)
        {
            Utils.MoveFormToCursor(this);
        }

        private void ClearCells(int column)
        {
            Excel.Range col;

            if (m_isPivotTable)
            {
                var pc = m_rng.PivotTable.TableRange1.Cells[1, column] as Excel.Range;
                string addr = pc.PivotTable.SourceData;
                addr = m_rng.Application.ConvertFormula(addr, Excel.XlReferenceStyle.xlR1C1, Excel.XlReferenceStyle.xlA1);
                Excel.Range sourceData = m_rng.Application.Range[addr];
                col = sourceData.Resize[1, sourceData.Rows.Count];
                col = col.Find(What: pc.PivotCell.PivotField.SourceName, LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlWhole, MatchCase: true);

                if (col.Valid())
                {
                    col = m_rng.Application.Intersect(col.EntireColumn, sourceData);
                }
                else
                {
                    MessageBox.Show("Error occured while clearing");
                    return;
                }
            }
            else
                col = m_rng.Columns[column];

            if (!searchContentsCheckBox.Checked)
            {
                col.Offset[1, 0].Resize[col.Cells.Count - 1, 1].ClearContents();
            }
            else
            {
                using (new ExcelExecutionBlock(m_app))
                {
                    foreach (var cell in col.Offset[1, 0].Resize[col.Cells.Count - 1, 1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Cast<Excel.Range>())
                    {
                        string value = cell.Value2?.ToString() ?? string.Empty;
                        if (value.Contains(searchTextBox.Text, StringComparison.OrdinalIgnoreCase))
                            cell.ClearContents();
                    }
                }
            }

            if (m_isPivotTable)
            {
                m_rng.PivotTable.RefreshTable();
            }
        }

        private void searchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (searchContentsCheckBox.Checked)
                return;

            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                dataGridView.Rows[i].Visible = (dataGridView[0, i].Value.ToString().Contains(searchTextBox.Text, StringComparison.OrdinalIgnoreCase));
            }
        }

        private void ColumnSelect(string columnName)
        {
            int columnIndex = -1;
            if (UseDataTable)
                columnIndex = m_dataTable.Columns[columnName].Ordinal + 1;
            else
                columnIndex = ((Excel.Range)m_rng.Rows[1]).Cells.Cast<Excel.Range>().FirstOrDefault(p => columnName == (p.Value2?.ToString() ?? $"Column{p.Column - m_rng.Column + 1}")).Column - m_rng.Column + 1;
            m_rng.Columns[columnIndex].Select();
        }

        private void okBtn_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        ~SearchColumnsForm()
        {
            Close();
        }

        private void searchBtn_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            SortedDictionary<string, long> dataDic = new SortedDictionary<string, long>();

            if (string.IsNullOrEmpty(searchTextBox.Text))
                dataDic = GetCounts(m_rng);
            else if (UseDataTable)
                dataDic = Utils.GetCounts(m_dataTable, searchTextBox.Text);
            else
                dataDic = UtilsExcel.GetCounts(m_rng, searchTextBox.Text);

            if (dataDic.Count != dataGridView.Rows.Count)
            {
                MessageBox.Show("Error");
                this.Cursor = Cursors.Default;
                return;
            }

            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                long val = dataDic[dataGridView[0, i].Value.ToString()];
                dataGridView[1, i].Value = val.ToString();

                if (string.IsNullOrEmpty(searchTextBox.Text))
                    dataGridView.Rows[i].Visible = true;
                else
                    dataGridView.Rows[i].Visible = (val > 0);
            }
            this.Cursor = Cursors.Default;
        }

        private void searchContentsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            searchBtn.Enabled = searchContentsCheckBox.Checked;
        }

        private void fetchBtn_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            Fetch();
            this.Cursor = Cursors.Default;
        }

        private bool Fetch()
        {
            searchContentsCheckBox.Checked = false;
            searchTextBox.Text = string.Empty;

            m_rng = m_app.ActiveWindow.RangeSelection.GetUsableRange();
            m_isPivotTable = m_rng.IsPivotCell();
            if (m_rng.Columns.Count < 2)
            {
                if (m_rng.IsPivotCell())
                {
                    m_rng = m_rng.PivotCell.PivotTable.TableRange1;
                    m_isPivotTable = true;
                }
                else
                    m_rng = m_rng.CurrentRegion;
            }
            else
                m_isPivotTable = false;

            if (!m_rng.Valid())
            {
                m_rng = null;
                m_dataTable = null;
                dataGridView.Rows.Clear();
                return false;
            }

            SortedDictionary<string, long> dataDic = new SortedDictionary<string, long>();

            if (UseDataTable)
            {
                m_dataTable = m_rng.GetDataTable2();
                dataDic = Utils.GetCounts(m_dataTable);
            }
            else
                dataDic = GetCounts(m_rng);

            dataGridView.Rows.Clear();
            foreach (var item in dataDic)
            {
                dataGridView.Rows.Add(item.Key, item.Value.ToString(), false, "Clear");
            }

            return true;
        }

        private SortedDictionary<string, long> GetCounts(Excel.Range rng)
        {
            var counts = new SortedDictionary<string, long>();
            object[,] cellValues = (object[,])rng.Value2;
            int column = rng.Column;

            // Synchronize access to the shared counts dictionary
            object lockObject = new object();
            Parallel.For(1, cellValues.GetLength(1) + 1, i =>
            {
                string columnName;
                long count = 0;

                columnName = cellValues[1, i].ToString() ?? $"Column{i - column + 1}";
                for (int j = 2; j < cellValues.GetLength(0) + 1; j++)
                {
                    if (!string.IsNullOrEmpty(cellValues[j, i]?.ToString()))
                        ++count;
                }

                // Synchronize access to the counts dictionary
                lock (lockObject)
                {
                    counts[columnName] = count;
                }
            });

            return counts;
        }


        private void dataGridView_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView.IsCurrentCellDirty)
            {
                dataGridView.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dataGridView_OnCellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2 && e.RowIndex >= 0)
            {
                string columnName = dataGridView[0, e.RowIndex].Value.ToString();
                int columnIndex = -1;
                if (UseDataTable)
                    columnIndex = m_dataTable.Columns[columnName].Ordinal + 1;
                else
                    columnIndex = ((Excel.Range)m_rng.Rows[1]).Cells.Cast<Excel.Range>().FirstOrDefault(p => columnName == (p.Value2?.ToString() ?? $"Column{p.Column - m_rng.Column + 1}")).Column - m_rng.Column + 1;
                Excel.Range col = ((Excel.Range)m_rng.Columns[columnIndex]).Offset[1, 0].Resize[m_rng.Rows.Count - 1, 1];

                if (searchContentsCheckBox.Checked)
                {
                    if ((bool)dataGridView[e.ColumnIndex, e.RowIndex].Value)
                        UtilsExcel.ColorRange(col, UtilsExcel.RangeType.Cells, Color.Yellow, searchTextBox.Text);
                    else
                        UtilsExcel.ColorRange(col, UtilsExcel.RangeType.Cells, Color.Transparent, searchTextBox.Text);
                }
                else
                {
                    if ((bool)dataGridView[e.ColumnIndex, e.RowIndex].Value)
                        UtilsExcel.ColorRange(col, UtilsExcel.RangeType.Colums, Color.Yellow, string.Empty);
                    else
                        UtilsExcel.ColorRange(col, UtilsExcel.RangeType.Colums, Color.Transparent, string.Empty);
                }
            }
        }

        private void searchTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (searchContentsCheckBox.Checked)
                    searchBtn.PerformClick();
            }
        }

        private void countsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            dataGridView.Columns[1].Visible = countsCheckBox.Checked;
            if (countsCheckBox.Checked)
            {
                SortedDictionary<string, long> dataDic = GetCounts(m_rng);
                dataGridView.Rows.Clear();
                foreach (var item in dataDic)
                {
                    dataGridView.Rows.Add(item.Key, item.Value.ToString(), false, "Clear");
                }
            }            
            this.Cursor = Cursors.Default;
        }
    }
}
