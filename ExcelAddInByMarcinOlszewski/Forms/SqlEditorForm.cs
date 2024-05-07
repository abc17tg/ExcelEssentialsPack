using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAddInByMarcinOlszewski.Forms;
using ExcelAddInByMarcinOlszewski.Scripts;
using ScintillaNET;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInByMarcinOlszewski
{
    public partial class SqlEditorForm : Form
    {
        public static string FormTitle = "SQL Editor";
        public string Query;
        public int RunningQueries = 0;
        public Excel.Application App;
        public SqlServerManager.ServerType ServerType;
        public SqlConn SqlConn;
        public static string DefaultSheetName = "Sql Query";
        public static string NewSheetName;
        public bool PasteHeaders => headersCheckBox.Checked;
        public bool PasteToSelection => pasteResultsToSelectionCheckBox.Checked;

        private List<string> m_objectsListBoxAllItemsList = new List<string>();
        private List<string> m_objectsTablesListBoxAllItemsList = new List<string>();
        private List<string> m_objectsListBoxSelectedItemsList = new List<string>();
        private Dictionary<string, SqlConn> m_connDic;
        private Dictionary<string, string> m_queriesDic;
        private static readonly string m_sheetNameTextBoxPlaceholder = "Worksheet name";

        public const Int32 WM_SYSCOMMAND = 0x112;
        public const Int32 MF_BYPOSITION = 0x400;
        public const Int32 ToggleTopMostMenuItem = 1000;

        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll")]
        private static extern bool InsertMenu(IntPtr hMenu, Int32 wPosition, Int32 wFlags, Int32 wIDNewItem, string lpNewItem);


        public SqlEditorForm(Excel.Application app)
        {
            InitializeComponent();
            App = app;
            TopMost = true;
            /*app.WindowActivate += (_, w) => this.TopMost = true;
            app.WindowDeactivate += (_, w) => this.TopMost = false;*/

            serverTypeComboBox.Items.AddRange(Directory.EnumerateDirectories(FileManager.SqlQueriesPath).Select(p => Path.GetFileName(p)).ToArray());

            UtilsScintilla.SetupSqlEditor(sqlEditorScintilla);

            serverTypeComboBox.ContextMenuStrip = new ContextMenuStrip();
            serverTypeComboBox.ContextMenuStrip.Items.Add("Add Server Connection").Click += (sender, e) => SqlServerManager.AddSqlConnection();

            serverComboBox.ContextMenuStrip = new ContextMenuStrip();
            serverComboBox.ContextMenuStrip.Items.Add("Add Server Connection").Click += (sender, e) => SqlServerManager.AddSqlConnection();

            sheetNameTextBox.Enter += (s, e) =>
            {
                if (sheetNameTextBox.Text == m_sheetNameTextBoxPlaceholder)
                    sheetNameTextBox.Text = "";
            };

            searchTextBox.Enter += (s, e) =>
            {
                if (searchTextBox.Text == "Search")
                    searchTextBox.Text = "";
                else
                    searchTextBox.SelectAll();
            };

            searchTextBox.Leave += (s, e) =>
            {
                if (searchTextBox.Text == "")
                    searchTextBox.Text = "Search";
            };

            ContextMenu cm = new ContextMenu();

            MenuItem copyCMI = new MenuItem("Copy", (o, e) => { (o as Scintilla).Copy(); });
            MenuItem pasteCMI = new MenuItem("Paste", (o, e) => { (o as Scintilla).Paste(); });
            MenuItem fetchCMI = new MenuItem("Fetch", (o, e) => { fetchBtn.PerformClick(); });
            MenuItem commentCMI = new MenuItem("Comment", (o, e) => { commentBtn.PerformClick(); });
            MenuItem pasteRangeCMI = new MenuItem("Paste range", (o, e) => { pasteRngBtn.PerformClick(); });
            MenuItem runSelectionCMI = new MenuItem("Run selected", (o, e) => { runSelectionBtn.PerformClick(); });
            MenuItem pasteClipboardRangeCMI = new MenuItem("Paste rng from clipboard", (o, e) => { PasteFromClipboard(); });
            cm.MenuItems.Add(pasteCMI);
            cm.MenuItems.Add(copyCMI);
            cm.MenuItems.Add(fetchCMI);
            cm.MenuItems.Add(commentCMI);
            cm.MenuItems.Add(pasteRangeCMI);
            cm.MenuItems.Add(pasteClipboardRangeCMI);
            cm.MenuItems.Add(runSelectionCMI);
            sqlEditorScintilla.ContextMenu = cm;
        }

        protected override void WndProc(ref Message msg)
        {
            if (msg.Msg == WM_SYSCOMMAND)
            {
                switch (msg.WParam.ToInt32())
                {
                    case ToggleTopMostMenuItem:
                        ToggleTopMost();
                        return;
                    default:
                        break;
                }
            }
            base.WndProc(ref msg);
        }

        private void SqlEditorForm_Load(object sender, EventArgs e)
        {
            IntPtr MenuHandle = GetSystemMenu(this.Handle, false);
            InsertMenu(MenuHandle, 5, MF_BYPOSITION, ToggleTopMostMenuItem, "Pin/Unpin this window");
        }

        private void ToggleTopMost()
        {
            this.TopMost = !this.TopMost;
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void validateBtn_Click(object sender, EventArgs e)
        {
            List<string> errors = new List<string>();
            if (!Utils.IsSQLQueryValid(sqlEditorScintilla.Text, out errors))
                MessageBox.Show(string.Join(Environment.NewLine, errors));
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

        private void pasteRngBtn_Click(object sender, EventArgs e)
        {
            Excel.Range rng = App.ActiveWindow.RangeSelection;
            if (rng.Valid())
                sqlEditorScintilla.ReplaceSelection(UtilsExcel.FormatRangeToSqlPattern(rng));
        }

        private void pasteRngFilterBtn_Click(object sender, EventArgs e)
        {
            Excel.Range rng = App.ActiveWindow.RangeSelection;
            if (!rng.Valid())
                return;

            string rngText = UtilsExcel.GenerateSqlFilterFromExcelSelection(rng);
            if (!string.IsNullOrEmpty(rngText))
                sqlEditorScintilla.ReplaceSelection(rngText);
        }

        private void validateSelectionBtn_Click(object sender, EventArgs e)
        {
            List<string> errors = new List<string>();
            if (!Utils.IsSQLQueryValid(sqlEditorScintilla.SelectedText, out errors))
                MessageBox.Show(string.Join(Environment.NewLine, errors));
        }

        private void runBtn_Click(object sender, EventArgs e)
        {
            Query = sqlEditorScintilla.Text;
            Run(Query);
        }

        private void runSelectionBtn_Click(object sender, EventArgs e)
        {
            Query = sqlEditorScintilla.SelectedText;
            Run(Query);
        }

        private void Run(string query)
        {
            if (new List<string> { query, SqlConn?.ConnectionString(), serverComboBox.SelectedItem.ToString(), serverTypeComboBox.SelectedItem.ToString() }.Any(p => string.IsNullOrWhiteSpace(p)))
            {
                MessageBox.Show("Missing server selections or query", "Run error");
                return;
            }

            List<string> errors = new List<string>();
            if (!Utils.IsSQLQueryValid(query, out errors))
            {
                var result = MessageBox.Show(string.Join(Environment.NewLine, errors), "Syntax error", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                if (result != DialogResult.OK)
                    return;
            }

            Task runQuery = null;
            Task<SqlResult> runQueryWithResult = null;

            if (pasteToDataTableCheckBox.Checked)
                runQueryWithResult = new Task<SqlResult>(() => SqlServerManager.GetDataFromServer(query, SqlConn));
            else if (!pasteResultsToSelectionCheckBox.Checked)
                runQuery = new Task(() => SqlServerManager.GetDataFromServerToNewSheet(query, SqlConn, PasteHeaders, NewSheetName == m_sheetNameTextBoxPlaceholder ? DefaultSheetName : NewSheetName));
            else
                runQuery = new Task(() => SqlServerManager.GetDataFromServerToSelection(query, SqlConn, App.ActiveWindow.RangeSelection, PasteHeaders));

            if (++RunningQueries == 1)
                Text = $"{FormTitle} [{RunningQueries}] running queries";
            else
                Text = Regex.Replace(Text, @"\s\[\d+\]\srunning\squeries", $" [{RunningQueries}] running queries");


            if (pasteToDataTableCheckBox.Checked)
            {
                runQueryWithResult.GetAwaiter().OnCompleted(() =>
                {
                    if (this == null)
                        return;

                    --RunningQueries;
                    string text;
                    if (RunningQueries <= 0)
                        text = FormTitle;
                    else
                        text = Regex.Replace(Text, @"\s\[\d+\]\srunning\squeries", $" [{RunningQueries}] running queries");

                    this.Invoke(new Action(() =>
                    {
                        if (this != null)
                            this.Text = text;
                        SqlResult sqlResult = runQueryWithResult.Result;
                        if (sqlResult.HasErrors)
                            return;
                        DataTableForm form = new DataTableForm(sqlResult.DataTable, query, App);
                        form.Show();
                        form.Activate();
                    }));
                });

                runQueryWithResult.Start();
            }
            else
            {
                runQuery.GetAwaiter().OnCompleted(() =>
                {
                    if (this == null || this.IsDisposed)
                        return;

                    --RunningQueries;
                    string text;
                    if (RunningQueries <= 0)
                        text = FormTitle;
                    else
                        text = Regex.Replace(Text, @"\s\[\d+\]\srunning\squeries", $" [{RunningQueries}] running queries");

                    this.Invoke(new Action(() =>
                    {
                        if (this == null || this.IsDisposed)
                            return;
                        this.Text = text;
                    }));
                    //#if DEBUG
                    MessageBox.Show($"{query}\n\nFinished", $"{NewSheetName} query finished", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    //#endif
                });

                runQuery.Start();
            }
        }

        private void commentBtn_Click(object sender, EventArgs e)
        {
            UtilsScintilla.Comment(sqlEditorScintilla);
        }

        private void testConnBtn_Click(object sender, EventArgs e)
        {
            SqlConn sqlConn;
            bool result = m_connDic.TryGetValue(m_connDic.Keys.First(p => p == serverComboBox.SelectedItem.ToString()) ?? "", out sqlConn);
            if (result)
            {
                if (sqlConn.Test())
                    MessageBox.Show("Connection success!");
                else
                    MessageBox.Show("Connection failed!");
            }
            else
                MessageBox.Show("Connection failed!");
        }

        private void saveQueryBtn_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                AddExtension = true,
                DefaultExt = "sql",
                Title = "Save query",
                ValidateNames = true,
                FileName = "Untitled.sql",
                Filter = "SQL queries | *.sql",
                OverwritePrompt = true
            };

            switch (ServerType)
            {
                case SqlServerManager.ServerType.SqlServer:
                    saveFileDialog.InitialDirectory = FileManager.SqlServerQueriesPath;
                    break;
                case SqlServerManager.ServerType.Oracle:
                    saveFileDialog.InitialDirectory = FileManager.OracleQueriesPath;
                    break;
                case SqlServerManager.ServerType.Excel:
                    saveFileDialog.InitialDirectory = FileManager.ExcelQueriesPath;
                    break;
                default:
                    break;
            }

            var result = saveFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                using (var sw = File.CreateText(filePath))
                {
                    sw.Write(sqlEditorScintilla.Text);
                }
                savedQueriesComboBox.Items.Clear();
                savedQueriesComboBox.Items.AddRange(m_queriesDic.Keys.Select(p => Path.GetFileName(p)).ToArray());
            }
        }

        private void serverTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlServerManager.ServerType serverType = (SqlServerManager.ServerType)Enum.Parse(typeof(SqlServerManager.ServerType), serverTypeComboBox.SelectedItem.ToString());
            if (!Enum.IsDefined(typeof(SqlServerManager.ServerType), serverType))
                return;

            m_objectsTablesListBoxAllItemsList.Clear();

            switch (serverType)
            {
                case SqlServerManager.ServerType.SqlServer:
                    this.TopMost = false;
                    m_connDic = FileManager.GetSqlServerConnectionValues();
                    m_queriesDic = FileManager.GetSqlServerQueries();
                    this.TopMost = true;
                    break;
                case SqlServerManager.ServerType.Oracle:
                    this.TopMost = false;
                    m_connDic = FileManager.GetOracleConnectionValues();
                    m_queriesDic = FileManager.GetOracleQueries();
                    this.TopMost = true;
                    break;
                case SqlServerManager.ServerType.Excel:
                    this.TopMost = false;
                    m_connDic = null;
                    m_queriesDic = FileManager.GetExcelQueries();
                    this.TopMost = true;
                    break;
                default:
                    return;
            }
            ServerType = serverType;
            savedQueriesComboBox.Items.Clear();
            try
            {
                savedQueriesComboBox.Items.AddRange(m_queriesDic.Keys.Select(p => Path.GetFileName(p)).ToArray());
            }
            catch { }

            serverComboBox.Items.Clear();
            try
            {
                if (ServerType != SqlServerManager.ServerType.Excel)
                    serverComboBox.Items.AddRange(m_connDic.Keys.ToArray());
            }
            catch { }
        }

        private void savedQueriesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(sqlEditorScintilla.Text) && !(sqlEditorScintilla.Text.Trim() == "SELECT * FROM"))
            {
                DialogResult result;
                result = MessageBox.Show("That will load query and erase current one.\nDo you want to paste it below?", "Load query warning",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                switch (result)
                {
                    case DialogResult.Yes:
                        sqlEditorScintilla.Text = sqlEditorScintilla.Text.TrimEnd('\n', '\r', '\t', ' ');
                        int position = sqlEditorScintilla.Lines.Last().Position;
                        sqlEditorScintilla.AppendText($"\n\n{new string('-', 50)}\n\n{m_queriesDic[m_queriesDic.Keys.First(p => p.Contains(savedQueriesComboBox.SelectedItem.ToString()))]}");
                        sqlEditorScintilla.GotoPosition(position);
                        break;
                    case DialogResult.No:
                        sqlEditorScintilla.Text = m_queriesDic[m_queriesDic.Keys.First(p => p.Contains(savedQueriesComboBox.SelectedItem.ToString()))];
                        break;
                    case DialogResult.Cancel:
                    case DialogResult.None:
                    default:
                        return;
                }
            }
            else
                sqlEditorScintilla.Text = m_queriesDic[m_queriesDic.Keys.First(p => p.Contains(savedQueriesComboBox.SelectedItem.ToString()))];
        }

        private void serverComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var result = m_connDic.TryGetValue((sender as ComboBox).SelectedItem.ToString(), out SqlConn);
            m_objectsTablesListBoxAllItemsList.Clear();
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

        private void fetchBtn_Click(object sender, EventArgs e)
        {
            bool tables = false;
            string query;

            m_objectsListBoxAllItemsList.Clear();
            m_objectsListBoxSelectedItemsList.Clear();

            if (string.IsNullOrWhiteSpace(sqlEditorScintilla.SelectedText))
            {
                tables = true;
                listObjectsTypeLabel.Text = "Tables";
            }
            else
            {
                listObjectsTypeLabel.Text = "Columns";
            }

            objectsListBox.Items.Clear();
            objectsListBox.Items.Add("Fetching...");
            objectsListBox.Update();

            SqlConn sqlConn;
            try
            {
                bool result = m_connDic.TryGetValue(m_connDic.Keys.FirstOrDefault(p => p.Contains(serverComboBox.SelectedItem.ToString())), out sqlConn);
                if (result)
                    result = sqlConn.Test();
                if (!result)
                    MessageBox.Show("Connection failed!");
            }
            catch
            {
                MessageBox.Show("Connection failed!");
                objectsListBox.Items.Clear();
                return;
            }
            if (tables)
            {
                if (tables && m_objectsTablesListBoxAllItemsList != null && m_objectsTablesListBoxAllItemsList.Count > 0)
                {
                    objectsListBox.Items.Clear();
                    objectsListBox.Items.AddRange(m_objectsTablesListBoxAllItemsList.ToArray());
                    m_objectsListBoxAllItemsList = m_objectsTablesListBoxAllItemsList;
                    m_objectsListBoxSelectedItemsList.Clear();
                    return;
                }
                switch (sqlConn.Type)
                {
                    case SqlServerManager.ServerType.SqlServer:
                        /*query = "CREATE TABLE #AllTables (Database_Schema_Table NVARCHAR(MAX));DECLARE @sql NVARCHAR(MAX) = N'';DECLARE @dbName NVARCHAR(128);DECLARE dbCursor CURSOR FOR SELECT [name] FROM sys.databases WHERE state = 0 AND [name] NOT IN ('master', 'tempdb', 'model', 'msdb');OPEN dbCursor;FETCH NEXT FROM dbCursor INTO @dbName;WHILE @@FETCH_STATUS = 0 BEGIN SET @sql = N'USE [' + @dbName + ']; INSERT INTO #AllTables SELECT ''' + @dbName + '.'' + SCHEMA_NAME(schema_id) + ''.'' + [name] FROM sys.tables t WHERE EXISTS (SELECT 1 FROM ' + QUOTENAME(@dbName) + '.sys.partitions p WHERE p.object_id = t.object_id AND p.rows > 0);'; BEGIN TRY EXEC sp_executesql @sql; END TRY BEGIN CATCH PRINT 'Error accessing database ' + @dbName + ': ' + ERROR_MESSAGE(); END CATCH; FETCH NEXT FROM dbCursor INTO @dbName;END CLOSE dbCursor;DEALLOCATE dbCursor;SELECT * FROM #AllTables ORDER BY Database_Schema_Table;DROP TABLE #AllTables;";*/
                        query = "CREATE TABLE #AllTables (Database_Schema_Object NVARCHAR(MAX)); DECLARE @sql NVARCHAR(MAX) = N''; DECLARE @dbName NVARCHAR(128); DECLARE dbCursor CURSOR FOR SELECT [name] FROM sys.databases WHERE state = 0 AND [name] NOT IN ('master', 'tempdb', 'model', 'msdb'); OPEN dbCursor; FETCH NEXT FROM dbCursor INTO @dbName; WHILE @@FETCH_STATUS = 0 BEGIN SET @sql = N'USE [' + @dbName + ']; INSERT INTO #AllTables SELECT ''' + @dbName + '.'' + SCHEMA_NAME(schema_id) + ''.'' + [name] FROM sys.tables t WHERE EXISTS (SELECT 1 FROM ' + QUOTENAME(@dbName) + '.sys.partitions p WHERE p.object_id = t.object_id AND p.rows > 0) UNION ALL SELECT ''' + @dbName + '.'' + SCHEMA_NAME(schema_id) + ''.'' + [name] FROM sys.views v;'; BEGIN TRY EXEC sp_executesql @sql; END TRY BEGIN CATCH PRINT 'Error accessing database ' + @dbName + ': ' + ERROR_MESSAGE(); END CATCH; FETCH NEXT FROM dbCursor INTO @dbName; END CLOSE dbCursor; DEALLOCATE dbCursor; SELECT * FROM #AllTables ORDER BY Database_Schema_Object; DROP TABLE #AllTables;";
                        break;
                    case SqlServerManager.ServerType.Oracle:
                        /*query = "SELECT TABLE_NAME FROM ALL_TABLES WHERE ( ( TABLE_NAME LIKE '%S4%' OR TABLE_NAME LIKE '%EMEA%' OR TABLE_NAME LIKE '%MDM%' OR TABLE_NAME LIKE '%ECDER%' OR TABLE_NAME LIKE '%ECLIPSE%' OR TABLE_NAME LIKE '%POLARIS%' OR TABLE_NAME LIKE '%SIQP%' OR TABLE_NAME LIKE '%EDW%' OR TABLE_NAME LIKE '%MCW%' OR TABLE_NAME LIKE '%OMEGA%' OR TABLE_NAME LIKE '%SAP%' ) AND ( TABLE_NAME NOT LIKE '%AMS%' AND TABLE_NAME NOT LIKE '%APJ%' AND TABLE_NAME NOT LIKE '%TEMP%' AND TABLE_NAME NOT LIKE '%TEST%' AND TABLE_NAME NOT LIKE '%OLD%' AND TABLE_NAME NOT LIKE '%JAPAN%' AND TABLE_NAME NOT LIKE '%US_%' AND TABLE_NAME NOT LIKE '%AP_%' )) ORDER BY TABLE_NAME";*/
                        query = "SELECT TABLE_NAME FROM ALL_TABLES WHERE NUM_ROWS > 0 ORDER BY TABLE_NAME";
                        break;
                    case SqlServerManager.ServerType.Excel:
                        objectsListBox.Items.Clear();
                        return;
                    default:
                        objectsListBox.Items.Clear();
                        return;
                }
            }
            else
                query = $"SELECT * FROM {sqlEditorScintilla.SelectedText.Trim()} WHERE 1=0";


            var sqlResult = SqlServerManager.GetDataFromServer(query, sqlConn, 40);
            objectsListBox.Items.Clear();
            if (!sqlResult.HasErrors)
            {
                if (tables)
                    objectsListBox.Items.AddRange(sqlResult.DataTable.AsEnumerable().Select(row => row.Field<string>(0)).Distinct().ToArray() ?? new string[1]);
                else
                    objectsListBox.Items.AddRange(sqlResult.DataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).Distinct().ToArray());
            }

            if (objectsListBox.Items.Count > 0)
            {
                m_objectsListBoxAllItemsList.AddRange(objectsListBox.Items.Cast<string>().ToList());
                m_objectsTablesListBoxAllItemsList = m_objectsListBoxAllItemsList;
            }
        }
        private void transferToQueryBtn_Click(object sender, EventArgs e)
        {
            string text = string.Empty;
            foreach (var obj in objectsListBox.SelectedItems)
            {
                if (!obj.ToString().Contains(" "))
                    text += $", {obj.ToString()}";
                else
                    text += $", [{obj.ToString()}]";
            }

            int lastWordRange = Math.Max(sqlEditorScintilla.WordStartPosition(sqlEditorScintilla.SelectionStart, true) - 10, 0);
            string lastText = sqlEditorScintilla.GetTextRange(lastWordRange, Math.Min(sqlEditorScintilla.SelectionStart, 10)).TrimEnd('\t', '\n', '\r', ' ');
            if (listObjectsTypeLabel.Text == "Tables" ||
                string.IsNullOrWhiteSpace(lastText) ||
                lastText.EndsWith("select", true, System.Globalization.CultureInfo.InvariantCulture) ||
                UtilsScintilla.SqlKeywords.Split(' ').Any(p => lastText.EndsWith(p, true, System.Globalization.CultureInfo.InvariantCulture)) || lastText.EndsWith("("))
            {
                if ((new char[] {' ','\t'}).ToList().Contains((char)sqlEditorScintilla.GetCharAt(sqlEditorScintilla.SelectionStart - 1)))
                    sqlEditorScintilla.ReplaceSelection(text?.TrimStart(',', ' ') ?? "");
                else
                    sqlEditorScintilla.ReplaceSelection(text?.TrimStart(',') ?? "");
            }
            else
                sqlEditorScintilla.ReplaceSelection(text ?? "");
        }

        private void wrapIntoBlockBtn_Click(object sender, EventArgs e)
        {
            UtilsScintilla.WrapIntoSqlBlock(sqlEditorScintilla);
        }

        private void openInNotepadBtn_Click(object sender, EventArgs e)
        {
            if (sqlEditorScintilla.Text.Length > 1)
                FileManager.OpenStringWithNotepad(sqlEditorScintilla.Text);
        }

        private void sheetNameTextBox_Leave(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(sheetNameTextBox.Text))
                NewSheetName = sheetNameTextBox.Text;
            else
                NewSheetName = m_sheetNameTextBoxPlaceholder;
        }

        private void pasteResultsToSelectionCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (pasteResultsToSelectionCheckBox.Checked)
            {
                sheetNameTextBox.Enabled = false;
                fillSheetNameBtn.Enabled = false;
            }
            else
            {
                sheetNameTextBox.Enabled = true;
                fillSheetNameBtn.Enabled = true;
            }
        }
        private void pasteToDataTableCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (pasteToDataTableCheckBox.Checked)
            {
                sheetNameTextBox.Enabled = false;
                fillSheetNameBtn.Enabled = false;
                pasteResultsToSelectionCheckBox.Enabled = false;
                headersCheckBox.Enabled = false;
            }
            else
            {
                if (pasteResultsToSelectionCheckBox.Checked)
                {
                    sheetNameTextBox.Enabled = false;
                    fillSheetNameBtn.Enabled = false;
                }
                else
                {
                    sheetNameTextBox.Enabled = true;
                    fillSheetNameBtn.Enabled = true;
                }
                pasteResultsToSelectionCheckBox.Enabled = true;
                headersCheckBox.Enabled = true;
            }
        }

        private void fillSheetNameBtn_Click(object sender, EventArgs e)
        {
            string word = sqlEditorScintilla.GetWordFromPosition(sqlEditorScintilla.CurrentPosition);
            if (!string.IsNullOrWhiteSpace(word))
            {
                sheetNameTextBox.Text = word;
                sheetNameTextBox.Focus();
            }
        }

        private void objectsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // If Ctrl is not pressed, clear the selected items list first
            if ((Control.ModifierKeys & Keys.Control) == 0)
            {
                m_objectsListBoxSelectedItemsList.Clear();
            }

            // Add the newly selected items
            foreach (string item in objectsListBox.SelectedItems)
            {
                if (!m_objectsListBoxSelectedItemsList.Contains(item))
                {
                    m_objectsListBoxSelectedItemsList.Add(item);
                }
            }
        }

        private void searchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;

                if (m_objectsListBoxAllItemsList == null || m_objectsListBoxAllItemsList.Count < 1)
                    return;

                // Clear the ListBox
                objectsListBox.Items.Clear();

                // Filter the items and add them to the ListBox
                var filteredItems = m_objectsListBoxAllItemsList.Where(item => item.IndexOf(searchTextBox.Text, StringComparison.OrdinalIgnoreCase) >= 0 || m_objectsListBoxSelectedItemsList.Contains(item)).ToList();

                objectsListBox.Items.AddRange(filteredItems.ToArray());
                objectsListBox.Update();

                // Reselect the previously selected items
                for (int i = 0; i < objectsListBox.Items.Count; i++)
                {
                    var item = objectsListBox.Items[i].ToString();
                    if (m_objectsListBoxSelectedItemsList.Contains(item))
                    {
                        objectsListBox.SetSelected(i, true);
                    }
                }
            }
        }

        private void clearEditorLabel_Click(object sender, EventArgs e)
        {
            sqlEditorScintilla.ClearAll();
        }

        private void sqlEditorScintilla_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string word = sqlEditorScintilla.GetWordFromPosition(sqlEditorScintilla.CurrentPosition);
            if (!string.IsNullOrWhiteSpace(word))
            {
                sheetNameTextBox.Text = word;
                sheetNameTextBox.Focus();
            }
        }

        private void objectsListBox_DoubleClick(object sender, EventArgs e)
        {
            if (objectsListBox.SelectedItems.Count > 0)
                transferToQueryBtn.PerformClick();
        }

        private void SqlEditorForm_MouseClick(object sender, MouseEventArgs e)
        {

        }
    }
}
