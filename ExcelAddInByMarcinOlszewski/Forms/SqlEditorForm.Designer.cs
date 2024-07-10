using System.Windows.Forms;

namespace ExcelAddInByMarcinOlszewski
{
    partial class SqlEditorForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SqlEditorForm));
            this.sqlEditorScintilla = new ScintillaNET.Scintilla();
            this.transferToQueryBtn = new System.Windows.Forms.Button();
            this.fetchBtn = new System.Windows.Forms.Button();
            this.objectsListBox = new System.Windows.Forms.ListBox();
            this.listObjectsTypeLabel = new System.Windows.Forms.Label();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.pasteResultsToSelectionCheckBox = new System.Windows.Forms.CheckBox();
            this.sheetNameTextBox = new System.Windows.Forms.TextBox();
            this.fillSheetNameBtn = new System.Windows.Forms.Button();
            this.headersCheckBox = new System.Windows.Forms.CheckBox();
            this.clearEditorLabel = new System.Windows.Forms.Label();
            this.mainTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.upperTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.parametersTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.worksheetTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.pasteToDataTableCheckBox = new System.Windows.Forms.CheckBox();
            this.objectsTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.objectsButtonsTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.buttonsTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.validateSelectionBtn = new System.Windows.Forms.Button();
            this.pasteRngBtn = new System.Windows.Forms.Button();
            this.commentBtn = new System.Windows.Forms.Button();
            this.serverTypeComboBox = new System.Windows.Forms.ComboBox();
            this.serverComboBox = new System.Windows.Forms.ComboBox();
            this.testConnBtn = new System.Windows.Forms.Button();
            this.openInNotepadBtn = new System.Windows.Forms.Button();
            this.runSelectionBtn = new System.Windows.Forms.Button();
            this.validateBtn = new System.Windows.Forms.Button();
            this.pasteRngFilterBtn = new System.Windows.Forms.Button();
            this.wrapIntoBlockBtn = new System.Windows.Forms.Button();
            this.savedQueriesComboBox = new System.Windows.Forms.ComboBox();
            this.saveQueryBtn = new System.Windows.Forms.Button();
            this.runBtn = new System.Windows.Forms.Button();
            this.mainTableLayoutPanel.SuspendLayout();
            this.upperTableLayoutPanel.SuspendLayout();
            this.parametersTableLayoutPanel.SuspendLayout();
            this.worksheetTableLayoutPanel.SuspendLayout();
            this.objectsTableLayoutPanel.SuspendLayout();
            this.objectsButtonsTableLayoutPanel.SuspendLayout();
            this.buttonsTableLayoutPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // sqlEditorScintilla
            // 
            this.sqlEditorScintilla.AllowDrop = true;
            this.sqlEditorScintilla.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.sqlEditorScintilla.CaretForeColor = System.Drawing.Color.White;
            this.sqlEditorScintilla.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sqlEditorScintilla.Lexer = ScintillaNET.Lexer.Sql;
            this.sqlEditorScintilla.Location = new System.Drawing.Point(3, 48);
            this.sqlEditorScintilla.Name = "sqlEditorScintilla";
            this.sqlEditorScintilla.Size = new System.Drawing.Size(728, 616);
            this.sqlEditorScintilla.TabIndex = 9;
            this.sqlEditorScintilla.Text = "SELECT * FROM";
            this.sqlEditorScintilla.UseTabs = true;
            this.sqlEditorScintilla.WrapIndentMode = ScintillaNET.WrapIndentMode.Indent;
            this.sqlEditorScintilla.WrapMode = ScintillaNET.WrapMode.Word;
            this.sqlEditorScintilla.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.sqlEditorScintilla_KeyPress);
            this.sqlEditorScintilla.KeyUp += new System.Windows.Forms.KeyEventHandler(this.sqlEditorScintilla_KeyUp);
            this.sqlEditorScintilla.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.sqlEditorScintilla_MouseDoubleClick);
            // 
            // transferToQueryBtn
            // 
            this.transferToQueryBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.transferToQueryBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.transferToQueryBtn.Location = new System.Drawing.Point(3, 3);
            this.transferToQueryBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.transferToQueryBtn.Name = "transferToQueryBtn";
            this.transferToQueryBtn.Size = new System.Drawing.Size(52, 26);
            this.transferToQueryBtn.TabIndex = 19;
            this.transferToQueryBtn.Text = "←";
            this.transferToQueryBtn.UseVisualStyleBackColor = true;
            this.transferToQueryBtn.Click += new System.EventHandler(this.transferToQueryBtn_Click);
            // 
            // fetchBtn
            // 
            this.fetchBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fetchBtn.Location = new System.Drawing.Point(61, 3);
            this.fetchBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.fetchBtn.Name = "fetchBtn";
            this.fetchBtn.Size = new System.Drawing.Size(169, 26);
            this.fetchBtn.TabIndex = 18;
            this.fetchBtn.Text = "Fetch";
            this.fetchBtn.UseVisualStyleBackColor = true;
            this.fetchBtn.Click += new System.EventHandler(this.fetchBtn_Click);
            // 
            // objectsListBox
            // 
            this.objectsListBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.objectsListBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.objectsListBox.FormattingEnabled = true;
            this.objectsListBox.HorizontalScrollbar = true;
            this.objectsListBox.ItemHeight = 12;
            this.objectsListBox.Location = new System.Drawing.Point(3, 30);
            this.objectsListBox.Name = "objectsListBox";
            this.objectsListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.objectsListBox.Size = new System.Drawing.Size(233, 545);
            this.objectsListBox.TabIndex = 20;
            this.objectsListBox.SelectedIndexChanged += new System.EventHandler(this.objectsListBox_SelectedIndexChanged);
            this.objectsListBox.DoubleClick += new System.EventHandler(this.objectsListBox_DoubleClick);
            // 
            // listObjectsTypeLabel
            // 
            this.listObjectsTypeLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listObjectsTypeLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listObjectsTypeLabel.Location = new System.Drawing.Point(737, 0);
            this.listObjectsTypeLabel.Name = "listObjectsTypeLabel";
            this.listObjectsTypeLabel.Size = new System.Drawing.Size(239, 45);
            this.listObjectsTypeLabel.TabIndex = 21;
            this.listObjectsTypeLabel.Text = "Objects";
            this.listObjectsTypeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // searchTextBox
            // 
            this.searchTextBox.AcceptsReturn = true;
            this.searchTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.searchTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.searchTextBox.Location = new System.Drawing.Point(3, 3);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(233, 21);
            this.searchTextBox.TabIndex = 24;
            this.searchTextBox.Text = "Search";
            this.searchTextBox.WordWrap = false;
            this.searchTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.searchTextBox_KeyDown);
            // 
            // pasteResultsToSelectionCheckBox
            // 
            this.pasteResultsToSelectionCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pasteResultsToSelectionCheckBox.Location = new System.Drawing.Point(315, 3);
            this.pasteResultsToSelectionCheckBox.Name = "pasteResultsToSelectionCheckBox";
            this.pasteResultsToSelectionCheckBox.Size = new System.Drawing.Size(186, 33);
            this.pasteResultsToSelectionCheckBox.TabIndex = 25;
            this.pasteResultsToSelectionCheckBox.Text = "Paste result to selection";
            this.pasteResultsToSelectionCheckBox.UseVisualStyleBackColor = true;
            this.pasteResultsToSelectionCheckBox.CheckedChanged += new System.EventHandler(this.pasteResultsToSelectionCheckBox_CheckedChanged);
            // 
            // sheetNameTextBox
            // 
            this.sheetNameTextBox.AllowDrop = true;
            this.sheetNameTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sheetNameTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.sheetNameTextBox.Location = new System.Drawing.Point(3, 3);
            this.sheetNameTextBox.Name = "sheetNameTextBox";
            this.sheetNameTextBox.Size = new System.Drawing.Size(147, 21);
            this.sheetNameTextBox.TabIndex = 26;
            this.sheetNameTextBox.Text = "Worksheet name";
            this.sheetNameTextBox.Leave += new System.EventHandler(this.sheetNameTextBox_Leave);
            // 
            // fillSheetNameBtn
            // 
            this.fillSheetNameBtn.Location = new System.Drawing.Point(156, 3);
            this.fillSheetNameBtn.Name = "fillSheetNameBtn";
            this.fillSheetNameBtn.Size = new System.Drawing.Size(27, 23);
            this.fillSheetNameBtn.TabIndex = 27;
            this.fillSheetNameBtn.Text = "▲";
            this.fillSheetNameBtn.UseVisualStyleBackColor = true;
            this.fillSheetNameBtn.Click += new System.EventHandler(this.fillSheetNameBtn_Click);
            // 
            // headersCheckBox
            // 
            this.headersCheckBox.Checked = true;
            this.headersCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.headersCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.headersCheckBox.Location = new System.Drawing.Point(3, 3);
            this.headersCheckBox.Name = "headersCheckBox";
            this.headersCheckBox.Size = new System.Drawing.Size(114, 33);
            this.headersCheckBox.TabIndex = 28;
            this.headersCheckBox.Text = "Headers";
            this.headersCheckBox.UseVisualStyleBackColor = true;
            // 
            // clearEditorLabel
            // 
            this.clearEditorLabel.AutoSize = true;
            this.clearEditorLabel.BackColor = System.Drawing.Color.White;
            this.clearEditorLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.clearEditorLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.clearEditorLabel.ForeColor = System.Drawing.Color.IndianRed;
            this.clearEditorLabel.Location = new System.Drawing.Point(699, 0);
            this.clearEditorLabel.Name = "clearEditorLabel";
            this.clearEditorLabel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.clearEditorLabel.Size = new System.Drawing.Size(26, 39);
            this.clearEditorLabel.TabIndex = 29;
            this.clearEditorLabel.Text = "❌";
            this.clearEditorLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.clearEditorLabel.Click += new System.EventHandler(this.clearEditorLabel_Click);
            // 
            // mainTableLayoutPanel
            // 
            this.mainTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.mainTableLayoutPanel.ColumnCount = 1;
            this.mainTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainTableLayoutPanel.Controls.Add(this.upperTableLayoutPanel, 0, 0);
            this.mainTableLayoutPanel.Controls.Add(this.buttonsTableLayoutPanel, 0, 1);
            this.mainTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainTableLayoutPanel.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
            this.mainTableLayoutPanel.Location = new System.Drawing.Point(0, 0);
            this.mainTableLayoutPanel.Name = "mainTableLayoutPanel";
            this.mainTableLayoutPanel.RowCount = 2;
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 65F));
            this.mainTableLayoutPanel.Size = new System.Drawing.Size(985, 738);
            this.mainTableLayoutPanel.TabIndex = 30;
            // 
            // upperTableLayoutPanel
            // 
            this.upperTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.upperTableLayoutPanel.ColumnCount = 2;
            this.upperTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 75F));
            this.upperTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.upperTableLayoutPanel.Controls.Add(this.listObjectsTypeLabel, 1, 0);
            this.upperTableLayoutPanel.Controls.Add(this.parametersTableLayoutPanel, 0, 0);
            this.upperTableLayoutPanel.Controls.Add(this.sqlEditorScintilla, 0, 1);
            this.upperTableLayoutPanel.Controls.Add(this.objectsTableLayoutPanel, 1, 1);
            this.upperTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.upperTableLayoutPanel.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
            this.upperTableLayoutPanel.Location = new System.Drawing.Point(3, 3);
            this.upperTableLayoutPanel.Name = "upperTableLayoutPanel";
            this.upperTableLayoutPanel.RowCount = 2;
            this.upperTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45F));
            this.upperTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.upperTableLayoutPanel.Size = new System.Drawing.Size(979, 667);
            this.upperTableLayoutPanel.TabIndex = 32;
            // 
            // parametersTableLayoutPanel
            // 
            this.parametersTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.parametersTableLayoutPanel.ColumnCount = 5;
            this.parametersTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 17.34695F));
            this.parametersTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.55102F));
            this.parametersTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.55102F));
            this.parametersTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 27.55102F));
            this.parametersTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.parametersTableLayoutPanel.Controls.Add(this.headersCheckBox, 0, 0);
            this.parametersTableLayoutPanel.Controls.Add(this.worksheetTableLayoutPanel, 1, 0);
            this.parametersTableLayoutPanel.Controls.Add(this.pasteResultsToSelectionCheckBox, 2, 0);
            this.parametersTableLayoutPanel.Controls.Add(this.pasteToDataTableCheckBox, 3, 0);
            this.parametersTableLayoutPanel.Controls.Add(this.clearEditorLabel, 4, 0);
            this.parametersTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.parametersTableLayoutPanel.Location = new System.Drawing.Point(3, 3);
            this.parametersTableLayoutPanel.MinimumSize = new System.Drawing.Size(0, 35);
            this.parametersTableLayoutPanel.Name = "parametersTableLayoutPanel";
            this.parametersTableLayoutPanel.RowCount = 1;
            this.parametersTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 39F));
            this.parametersTableLayoutPanel.Size = new System.Drawing.Size(728, 39);
            this.parametersTableLayoutPanel.TabIndex = 31;
            // 
            // worksheetTableLayoutPanel
            // 
            this.worksheetTableLayoutPanel.AutoSize = true;
            this.worksheetTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.worksheetTableLayoutPanel.ColumnCount = 2;
            this.worksheetTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.worksheetTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.worksheetTableLayoutPanel.Controls.Add(this.sheetNameTextBox, 0, 0);
            this.worksheetTableLayoutPanel.Controls.Add(this.fillSheetNameBtn, 1, 0);
            this.worksheetTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.worksheetTableLayoutPanel.Location = new System.Drawing.Point(123, 3);
            this.worksheetTableLayoutPanel.Name = "worksheetTableLayoutPanel";
            this.worksheetTableLayoutPanel.RowCount = 1;
            this.worksheetTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.worksheetTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 33F));
            this.worksheetTableLayoutPanel.Size = new System.Drawing.Size(186, 33);
            this.worksheetTableLayoutPanel.TabIndex = 28;
            // 
            // pasteToDataTableCheckBox
            // 
            this.pasteToDataTableCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pasteToDataTableCheckBox.Location = new System.Drawing.Point(507, 3);
            this.pasteToDataTableCheckBox.Name = "pasteToDataTableCheckBox";
            this.pasteToDataTableCheckBox.Size = new System.Drawing.Size(186, 33);
            this.pasteToDataTableCheckBox.TabIndex = 33;
            this.pasteToDataTableCheckBox.Text = "Paste result to data table";
            this.pasteToDataTableCheckBox.UseVisualStyleBackColor = true;
            this.pasteToDataTableCheckBox.CheckedChanged += new System.EventHandler(this.pasteToDataTableCheckBox_CheckedChanged);
            // 
            // objectsTableLayoutPanel
            // 
            this.objectsTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.objectsTableLayoutPanel.ColumnCount = 1;
            this.objectsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.objectsTableLayoutPanel.Controls.Add(this.searchTextBox, 0, 0);
            this.objectsTableLayoutPanel.Controls.Add(this.objectsListBox, 0, 1);
            this.objectsTableLayoutPanel.Controls.Add(this.objectsButtonsTableLayoutPanel, 0, 2);
            this.objectsTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.objectsTableLayoutPanel.Location = new System.Drawing.Point(737, 48);
            this.objectsTableLayoutPanel.Name = "objectsTableLayoutPanel";
            this.objectsTableLayoutPanel.RowCount = 3;
            this.objectsTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.objectsTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.objectsTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.objectsTableLayoutPanel.Size = new System.Drawing.Size(239, 616);
            this.objectsTableLayoutPanel.TabIndex = 32;
            // 
            // objectsButtonsTableLayoutPanel
            // 
            this.objectsButtonsTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.objectsButtonsTableLayoutPanel.ColumnCount = 2;
            this.objectsButtonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.objectsButtonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 75F));
            this.objectsButtonsTableLayoutPanel.Controls.Add(this.transferToQueryBtn, 0, 0);
            this.objectsButtonsTableLayoutPanel.Controls.Add(this.fetchBtn, 1, 0);
            this.objectsButtonsTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.objectsButtonsTableLayoutPanel.Location = new System.Drawing.Point(3, 581);
            this.objectsButtonsTableLayoutPanel.Name = "objectsButtonsTableLayoutPanel";
            this.objectsButtonsTableLayoutPanel.RowCount = 1;
            this.objectsButtonsTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.objectsButtonsTableLayoutPanel.Size = new System.Drawing.Size(233, 32);
            this.objectsButtonsTableLayoutPanel.TabIndex = 31;
            // 
            // buttonsTableLayoutPanel
            // 
            this.buttonsTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.buttonsTableLayoutPanel.ColumnCount = 8;
            this.buttonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.buttonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.buttonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.buttonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.buttonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.buttonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.buttonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.buttonsTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11F));
            this.buttonsTableLayoutPanel.Controls.Add(this.validateSelectionBtn, 0, 0);
            this.buttonsTableLayoutPanel.Controls.Add(this.pasteRngBtn, 1, 0);
            this.buttonsTableLayoutPanel.Controls.Add(this.commentBtn, 2, 0);
            this.buttonsTableLayoutPanel.Controls.Add(this.serverTypeComboBox, 3, 0);
            this.buttonsTableLayoutPanel.Controls.Add(this.serverComboBox, 4, 0);
            this.buttonsTableLayoutPanel.Controls.Add(this.testConnBtn, 5, 0);
            this.buttonsTableLayoutPanel.Controls.Add(this.openInNotepadBtn, 6, 0);
            this.buttonsTableLayoutPanel.Controls.Add(this.runSelectionBtn, 7, 0);
            this.buttonsTableLayoutPanel.Controls.Add(this.validateBtn, 0, 1);
            this.buttonsTableLayoutPanel.Controls.Add(this.pasteRngFilterBtn, 1, 1);
            this.buttonsTableLayoutPanel.Controls.Add(this.wrapIntoBlockBtn, 2, 1);
            this.buttonsTableLayoutPanel.Controls.Add(this.savedQueriesComboBox, 3, 1);
            this.buttonsTableLayoutPanel.Controls.Add(this.saveQueryBtn, 6, 1);
            this.buttonsTableLayoutPanel.Controls.Add(this.runBtn, 7, 1);
            this.buttonsTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.buttonsTableLayoutPanel.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
            this.buttonsTableLayoutPanel.Location = new System.Drawing.Point(3, 676);
            this.buttonsTableLayoutPanel.MaximumSize = new System.Drawing.Size(0, 62);
            this.buttonsTableLayoutPanel.Name = "buttonsTableLayoutPanel";
            this.buttonsTableLayoutPanel.RowCount = 2;
            this.buttonsTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.buttonsTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.buttonsTableLayoutPanel.Size = new System.Drawing.Size(979, 59);
            this.buttonsTableLayoutPanel.TabIndex = 24;
            // 
            // validateSelectionBtn
            // 
            this.validateSelectionBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.validateSelectionBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.validateSelectionBtn.Location = new System.Drawing.Point(3, 3);
            this.validateSelectionBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.validateSelectionBtn.Name = "validateSelectionBtn";
            this.validateSelectionBtn.Size = new System.Drawing.Size(101, 25);
            this.validateSelectionBtn.TabIndex = 3;
            this.validateSelectionBtn.Text = "Validate selection";
            this.validateSelectionBtn.UseVisualStyleBackColor = true;
            this.validateSelectionBtn.Click += new System.EventHandler(this.validateSelectionBtn_Click);
            // 
            // pasteRngBtn
            // 
            this.pasteRngBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.pasteRngBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pasteRngBtn.Location = new System.Drawing.Point(110, 3);
            this.pasteRngBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.pasteRngBtn.Name = "pasteRngBtn";
            this.pasteRngBtn.Size = new System.Drawing.Size(101, 25);
            this.pasteRngBtn.TabIndex = 5;
            this.pasteRngBtn.Text = "Paste range";
            this.pasteRngBtn.UseVisualStyleBackColor = true;
            this.pasteRngBtn.Click += new System.EventHandler(this.pasteRngBtn_Click);
            // 
            // commentBtn
            // 
            this.commentBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.commentBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commentBtn.Location = new System.Drawing.Point(217, 3);
            this.commentBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.commentBtn.Name = "commentBtn";
            this.commentBtn.Size = new System.Drawing.Size(42, 25);
            this.commentBtn.TabIndex = 4;
            this.commentBtn.Text = "- - ...";
            this.commentBtn.UseVisualStyleBackColor = true;
            this.commentBtn.Click += new System.EventHandler(this.commentBtn_Click);
            // 
            // serverTypeComboBox
            // 
            this.serverTypeComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem;
            this.serverTypeComboBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.serverTypeComboBox.DropDownHeight = 210;
            this.serverTypeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.serverTypeComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.serverTypeComboBox.FormattingEnabled = true;
            this.serverTypeComboBox.IntegralHeight = false;
            this.serverTypeComboBox.ItemHeight = 15;
            this.serverTypeComboBox.Location = new System.Drawing.Point(265, 3);
            this.serverTypeComboBox.MaxDropDownItems = 15;
            this.serverTypeComboBox.Name = "serverTypeComboBox";
            this.serverTypeComboBox.Size = new System.Drawing.Size(189, 23);
            this.serverTypeComboBox.TabIndex = 8;
            this.serverTypeComboBox.SelectedIndexChanged += new System.EventHandler(this.serverTypeComboBox_SelectedIndexChanged);
            // 
            // serverComboBox
            // 
            this.serverComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem;
            this.serverComboBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.serverComboBox.DropDownHeight = 210;
            this.serverComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.serverComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.serverComboBox.FormattingEnabled = true;
            this.serverComboBox.IntegralHeight = false;
            this.serverComboBox.ItemHeight = 15;
            this.serverComboBox.Location = new System.Drawing.Point(460, 3);
            this.serverComboBox.MaxDropDownItems = 15;
            this.serverComboBox.Name = "serverComboBox";
            this.serverComboBox.Size = new System.Drawing.Size(189, 23);
            this.serverComboBox.TabIndex = 16;
            this.serverComboBox.SelectedIndexChanged += new System.EventHandler(this.serverComboBox_SelectedIndexChanged);
            // 
            // testConnBtn
            // 
            this.testConnBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.testConnBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.testConnBtn.Location = new System.Drawing.Point(655, 3);
            this.testConnBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.testConnBtn.Name = "testConnBtn";
            this.testConnBtn.Size = new System.Drawing.Size(101, 25);
            this.testConnBtn.TabIndex = 7;
            this.testConnBtn.Text = "Test connection";
            this.testConnBtn.UseVisualStyleBackColor = true;
            this.testConnBtn.Click += new System.EventHandler(this.testConnBtn_Click);
            // 
            // openInNotepadBtn
            // 
            this.openInNotepadBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.openInNotepadBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.openInNotepadBtn.Location = new System.Drawing.Point(762, 3);
            this.openInNotepadBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.openInNotepadBtn.Name = "openInNotepadBtn";
            this.openInNotepadBtn.Size = new System.Drawing.Size(101, 25);
            this.openInNotepadBtn.TabIndex = 23;
            this.openInNotepadBtn.Text = "Notepad";
            this.openInNotepadBtn.UseVisualStyleBackColor = true;
            this.openInNotepadBtn.Click += new System.EventHandler(this.openInNotepadBtn_Click);
            // 
            // runSelectionBtn
            // 
            this.runSelectionBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.runSelectionBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.runSelectionBtn.Location = new System.Drawing.Point(869, 3);
            this.runSelectionBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.runSelectionBtn.Name = "runSelectionBtn";
            this.runSelectionBtn.Size = new System.Drawing.Size(107, 25);
            this.runSelectionBtn.TabIndex = 2;
            this.runSelectionBtn.Text = "Run selection";
            this.runSelectionBtn.UseVisualStyleBackColor = true;
            this.runSelectionBtn.Click += new System.EventHandler(this.runSelectionBtn_Click);
            // 
            // validateBtn
            // 
            this.validateBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.validateBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.validateBtn.Location = new System.Drawing.Point(3, 32);
            this.validateBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.validateBtn.Name = "validateBtn";
            this.validateBtn.Size = new System.Drawing.Size(101, 25);
            this.validateBtn.TabIndex = 11;
            this.validateBtn.Text = "Validate";
            this.validateBtn.UseVisualStyleBackColor = true;
            this.validateBtn.Click += new System.EventHandler(this.validateBtn_Click);
            // 
            // pasteRngFilterBtn
            // 
            this.pasteRngFilterBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.pasteRngFilterBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pasteRngFilterBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pasteRngFilterBtn.Location = new System.Drawing.Point(110, 32);
            this.pasteRngFilterBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.pasteRngFilterBtn.Name = "pasteRngFilterBtn";
            this.pasteRngFilterBtn.Size = new System.Drawing.Size(101, 25);
            this.pasteRngFilterBtn.TabIndex = 13;
            this.pasteRngFilterBtn.Text = "Range as filter";
            this.pasteRngFilterBtn.UseVisualStyleBackColor = true;
            this.pasteRngFilterBtn.Click += new System.EventHandler(this.pasteRngFilterBtn_Click);
            // 
            // wrapIntoBlockBtn
            // 
            this.wrapIntoBlockBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.wrapIntoBlockBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wrapIntoBlockBtn.Location = new System.Drawing.Point(217, 32);
            this.wrapIntoBlockBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.wrapIntoBlockBtn.Name = "wrapIntoBlockBtn";
            this.wrapIntoBlockBtn.Size = new System.Drawing.Size(42, 25);
            this.wrapIntoBlockBtn.TabIndex = 22;
            this.wrapIntoBlockBtn.Text = "( ... )";
            this.wrapIntoBlockBtn.UseVisualStyleBackColor = true;
            this.wrapIntoBlockBtn.Click += new System.EventHandler(this.wrapIntoBlockBtn_Click);
            // 
            // savedQueriesComboBox
            // 
            this.savedQueriesComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystem;
            this.buttonsTableLayoutPanel.SetColumnSpan(this.savedQueriesComboBox, 3);
            this.savedQueriesComboBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.savedQueriesComboBox.DropDownHeight = 210;
            this.savedQueriesComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.savedQueriesComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.savedQueriesComboBox.FormattingEnabled = true;
            this.savedQueriesComboBox.IntegralHeight = false;
            this.savedQueriesComboBox.ItemHeight = 15;
            this.savedQueriesComboBox.Location = new System.Drawing.Point(265, 32);
            this.savedQueriesComboBox.MaxDropDownItems = 15;
            this.savedQueriesComboBox.Name = "savedQueriesComboBox";
            this.savedQueriesComboBox.Size = new System.Drawing.Size(491, 23);
            this.savedQueriesComboBox.TabIndex = 15;
            this.savedQueriesComboBox.SelectedIndexChanged += new System.EventHandler(this.savedQueriesComboBox_SelectedIndexChanged);
            // 
            // saveQueryBtn
            // 
            this.saveQueryBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.saveQueryBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.saveQueryBtn.Location = new System.Drawing.Point(762, 32);
            this.saveQueryBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.saveQueryBtn.Name = "saveQueryBtn";
            this.saveQueryBtn.Size = new System.Drawing.Size(101, 25);
            this.saveQueryBtn.TabIndex = 14;
            this.saveQueryBtn.Text = "Save query";
            this.saveQueryBtn.UseVisualStyleBackColor = true;
            this.saveQueryBtn.Click += new System.EventHandler(this.saveQueryBtn_Click);
            // 
            // runBtn
            // 
            this.runBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.runBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.runBtn.Location = new System.Drawing.Point(869, 32);
            this.runBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.runBtn.Name = "runBtn";
            this.runBtn.Size = new System.Drawing.Size(107, 25);
            this.runBtn.TabIndex = 10;
            this.runBtn.Text = "Run";
            this.runBtn.UseVisualStyleBackColor = true;
            this.runBtn.Click += new System.EventHandler(this.runBtn_Click);
            // 
            // SqlEditorForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(985, 738);
            this.Controls.Add(this.mainTableLayoutPanel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SqlEditorForm";
            this.Opacity = 0.95D;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "SQL Editor";
            this.Load += new System.EventHandler(this.SqlEditorForm_Load);
            this.mainTableLayoutPanel.ResumeLayout(false);
            this.upperTableLayoutPanel.ResumeLayout(false);
            this.parametersTableLayoutPanel.ResumeLayout(false);
            this.parametersTableLayoutPanel.PerformLayout();
            this.worksheetTableLayoutPanel.ResumeLayout(false);
            this.worksheetTableLayoutPanel.PerformLayout();
            this.objectsTableLayoutPanel.ResumeLayout(false);
            this.objectsTableLayoutPanel.PerformLayout();
            this.objectsButtonsTableLayoutPanel.ResumeLayout(false);
            this.buttonsTableLayoutPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }


        #endregion
        private ScintillaNET.Scintilla sqlEditorScintilla;
        private System.Windows.Forms.Button transferToQueryBtn;
        private System.Windows.Forms.Button fetchBtn;
        private System.Windows.Forms.ListBox objectsListBox;
        private System.Windows.Forms.Label listObjectsTypeLabel;
        private TextBox searchTextBox;
        private CheckBox pasteResultsToSelectionCheckBox;
        private TextBox sheetNameTextBox;
        private Button fillSheetNameBtn;
        private CheckBox headersCheckBox;
        private Label clearEditorLabel;
        private TableLayoutPanel mainTableLayoutPanel;
        private TableLayoutPanel parametersTableLayoutPanel;
        private CheckBox pasteToDataTableCheckBox;
        private TableLayoutPanel upperTableLayoutPanel;
        private TableLayoutPanel objectsButtonsTableLayoutPanel;
        private TableLayoutPanel objectsTableLayoutPanel;
        private TableLayoutPanel worksheetTableLayoutPanel;
        private TableLayoutPanel buttonsTableLayoutPanel;
        private Button validateSelectionBtn;
        private Button pasteRngBtn;
        private Button commentBtn;
        private ComboBox serverTypeComboBox;
        private ComboBox serverComboBox;
        private Button testConnBtn;
        private Button openInNotepadBtn;
        private Button runSelectionBtn;
        private Button validateBtn;
        private Button pasteRngFilterBtn;
        private Button wrapIntoBlockBtn;
        private ComboBox savedQueriesComboBox;
        private Button saveQueryBtn;
        private Button runBtn;
    }
}