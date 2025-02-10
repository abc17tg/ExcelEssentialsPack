using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelEssentials.Scripts;
using ScintillaNET;
using Excel = Microsoft.Office.Interop.Excel;
using WTC = ImportTableToExcel.WorksheetFromTxtCreator;

namespace ExcelEssentials.Forms
{
    public partial class SqlExcelTableForm : Form
    {
        public DataTable Table
        {
            get
            {
                return m_table;
            }
            set
            {
                if (value != null)
                {
                    m_table = value;
                    dataGridView.DataSource = m_table;
                    var nfi = new CultureInfo("en-US", false).NumberFormat;
                    nfi.NumberGroupSeparator = " ";
                    dimentionsLabel.Text = $"Rows: {(m_table.Rows.Count + 1).ToString("N0", nfi)} | Columns: {m_table.Columns.Count.ToString("N0", nfi)}";
                }
                else
                {
                    if (m_table == null)
                        m_table = new DataTable();
                    dimentionsLabel.Text = "Rows: null | Columns: null";
                }
            }
        }
        private DataTable m_table;
        private Excel.Application m_excelApp;

        public SqlExcelTableForm(Excel.Application application)
        {
            ScintillaFix.CopyNativeFolderIfNotExistOrDifferentFixForScintillaBug();
            InitializeComponent();
            UtilsScintilla.SetupSqlEditor(sqlEditorScintilla);
            ContextMenu cm = new ContextMenu();

            MenuItem copyCMI = new MenuItem("Copy", (o, e) => { (o as Scintilla).Copy(); });
            MenuItem pasteCMI = new MenuItem("Paste", (o, e) => { (o as Scintilla).Paste(); });
            MenuItem commentCMI = new MenuItem("Comment", (o, e) => { commentBtn.PerformClick(); });
            MenuItem pasteRangeCMI = new MenuItem("Paste range", (o, e) => { PasteRng(); });
            MenuItem runSelectionCMI = new MenuItem("Run selected", (o, e) => { runSelectionButton.PerformClick(); });
            MenuItem pasteClipboardRangeCMI = new MenuItem("Paste rng from clipboard", (o, e) => { PasteFromClipboard(); });
            cm.MenuItems.Add(pasteCMI);
            cm.MenuItems.Add(copyCMI);
            cm.MenuItems.Add(commentCMI);
            cm.MenuItems.Add(pasteRangeCMI);
            cm.MenuItems.Add(pasteClipboardRangeCMI);
            cm.MenuItems.Add(runSelectionCMI);
            sqlEditorScintilla.ContextMenu = cm;

            m_excelApp = application;
            Excel.Range rng = m_excelApp.ActiveWindow.RangeSelection.GetUsableRange();
            if (rng.Valid())
                Table = rng.GetDataTable(dataHasHeadersCheckBox.Checked);
        }


        public SqlExcelTableForm(Excel.Application application, DataTable dataTable)
        {
            InitializeComponent();

            UtilsScintilla.SetupSqlEditor(sqlEditorScintilla);
            ContextMenu cm = new ContextMenu();

            MenuItem copyCMI = new MenuItem("Copy", (o, e) => { (o as Scintilla).Copy(); });
            MenuItem pasteCMI = new MenuItem("Paste", (o, e) => { (o as Scintilla).Paste(); });
            MenuItem commentCMI = new MenuItem("Comment", (o, e) => { commentBtn.PerformClick(); });
            MenuItem pasteRangeCMI = new MenuItem("Paste range", (o, e) => { PasteRng(); });
            MenuItem runSelectionCMI = new MenuItem("Run selected", (o, e) => { runSelectionButton.PerformClick(); });
            MenuItem pasteClipboardRangeCMI = new MenuItem("Paste rng from clipboard", (o, e) => { PasteFromClipboard(); });
            cm.MenuItems.Add(pasteCMI);
            cm.MenuItems.Add(copyCMI);
            cm.MenuItems.Add(commentCMI);
            cm.MenuItems.Add(pasteRangeCMI);
            cm.MenuItems.Add(pasteClipboardRangeCMI);
            cm.MenuItems.Add(runSelectionCMI);
            sqlEditorScintilla.ContextMenu = cm;

            m_excelApp = application;
            Table = dataTable;
        }

        private void PasteRng()
        {
            Excel.Range rng = m_excelApp.ActiveWindow.RangeSelection;
            if (rng.Valid())
                sqlEditorScintilla.ReplaceSelection(UtilsExcel.FormatRangeToSqlPattern(rng));
        }

        private void loadButton_Click(object sender, EventArgs e)
        {
            Excel.Range rng = m_excelApp.ActiveWindow.RangeSelection;
            if (rng.Valid())
                Table = rng.GetDataTable(dataHasHeadersCheckBox.Checked);
        }

        private void runBtn_Click(object sender, EventArgs e)
        {
            Utils.ExecuteSqlQueryAndDisplayResults(dataGridView, sqlEditorScintilla.Text);
        }

        private void runSelectionBtn_Click(object sender, EventArgs e)
        {
            Utils.ExecuteSqlQueryAndDisplayResults(dataGridView, sqlEditorScintilla.SelectedText);
        }

        private void commentBtn_Click(object sender, EventArgs e)
        {
            UtilsScintilla.Comment(sqlEditorScintilla);
        }

        private void sqlEditorScintilla_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
                UtilsScintilla.IndentAfterReturn(sqlEditorScintilla);
        }

        private void sqlEditorScintilla_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Control && e.Shift && (e.KeyCode == Keys.Divide || e.KeyCode == Keys.Oem2))
            {
                UtilsScintilla.Comment(sqlEditorScintilla);
                e.SuppressKeyPress = true;
            }

            if (e.Alt)
            {
                if (e.KeyCode == Keys.Up)
                {
                    UtilsScintilla.MoveLineUp(sqlEditorScintilla);
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.Down)
                {
                    UtilsScintilla.MoveLineDown(sqlEditorScintilla);
                    e.Handled = true;
                }
            }
        }

        private void wrapIntoBlockBtn_Click(object sender, EventArgs e)
        {
            UtilsScintilla.WrapIntoSqlBlock(sqlEditorScintilla);
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void validateBtn_Click(object sender, EventArgs e)
        {
            /*List<string> errors = new List<string>();
            if (!Utils.IsSQLQueryValid(sqlEditorScintilla.Text, out errors))
                MessageBox.Show(string.Join(Environment.NewLine, errors));*/
        }

        private void PasteFromClipboard()
        {
            string text = Clipboard.GetText(TextDataFormat.Text);
            if (string.IsNullOrWhiteSpace(text))
                return;

            List<char> DelimiterChars = new List<char> { ' ', @"'"[0], '(', ')', ',', '.', '\t', '\n', '\r', ';', '|' };
            text = $"({string.Join(", ", text.Split(DelimiterChars.ToArray(), StringSplitOptions.RemoveEmptyEntries).Select(p => $"\'{p.Trim()}\'").ToArray())})";
            sqlEditorScintilla.ReplaceSelection(text);
        }

        private void validateSelectionBtn_Click(object sender, EventArgs e)
        {
            /*List<string> errors = new List<string>();
            if (!Utils.IsSQLQueryValid(sqlEditorScintilla.SelectedText, out errors))
                MessageBox.Show(string.Join(Environment.NewLine, errors))*/;
        }

        private void loadFromFilebutton_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            FileDropForm form = new FileDropForm(Utils.TextExt.Concat(Utils.ExcelExt).ToList());
            form.Show();
            form.FormClosed += (s, _) =>
            {
                filePath = form.FilePath;
                if (Utils.TextExt.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
                {
                    char delimiter = Utils.DetermineTableDelimiter(filePath);
                    DataTable dt = WTC.ReadFileIntoDataTable(filePath, delimiter);
                    if (dt != null)
                        Table = dt;
                    else
                        MessageBox.Show("Error when loading a file to data table", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
        }

        private void pasteButton_Click(object sender, EventArgs e)
        {
            Paste();
        }

        public void Paste()
        {
            Excel.Range rng = m_excelApp.ActiveWindow.RangeSelection;
            if (rng.Valid())
            {
                UtilsExcel.PasteDataTableToRange(Table, rng, dataHasHeadersCheckBox.Checked);
            }
            else
            {
                var result = MessageBox.Show("No selection to paste", "Error", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                if (result == DialogResult.Retry)
                    Paste();
            }
        }
    }
}
