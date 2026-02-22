namespace ExcelEssentials.Forms
{
    partial class SqlExcelTableForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SqlExcelTableForm));
            this.sqlEditorScintilla = new ScintillaNET.Scintilla();
            this.mainTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.topTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.loadFromFilebutton = new System.Windows.Forms.Button();
            this.runButton = new System.Windows.Forms.Button();
            this.runSelectionButton = new System.Windows.Forms.Button();
            this.loadButton = new System.Windows.Forms.Button();
            this.dataHasHeadersCheckBox = new System.Windows.Forms.CheckBox();
            this.dimentionsLabel = new System.Windows.Forms.Label();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.bottomTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.commentBtn = new System.Windows.Forms.Button();
            this.wrapIntoBlockBtn = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();
            this.pasteButton = new System.Windows.Forms.Button();
            this.validateButton = new System.Windows.Forms.Button();
            this.mainTableLayoutPanel.SuspendLayout();
            this.topTableLayoutPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.bottomTableLayoutPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // sqlEditorScintilla
            // 
            this.sqlEditorScintilla.AllowDrop = true;
            this.sqlEditorScintilla.BorderStyle = ScintillaNET.BorderStyle.FixedSingle;
            this.sqlEditorScintilla.CaretForeColor = System.Drawing.Color.White;
            this.sqlEditorScintilla.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sqlEditorScintilla.LexerName = "sql";
            this.sqlEditorScintilla.Location = new System.Drawing.Point(3, 304);
            this.sqlEditorScintilla.Name = "sqlEditorScintilla";
            this.sqlEditorScintilla.Size = new System.Drawing.Size(857, 478);
            this.sqlEditorScintilla.TabIndex = 10;
            this.sqlEditorScintilla.Text = "SELECT * FROM DataTable";
            this.sqlEditorScintilla.UseTabs = true;
            this.sqlEditorScintilla.WrapIndentMode = ScintillaNET.WrapIndentMode.Indent;
            this.sqlEditorScintilla.WrapMode = ScintillaNET.WrapMode.Word;
            this.sqlEditorScintilla.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.sqlEditorScintilla_KeyPress);
            this.sqlEditorScintilla.KeyUp += new System.Windows.Forms.KeyEventHandler(this.sqlEditorScintilla_KeyUp);
            // 
            // mainTableLayoutPanel
            // 
            this.mainTableLayoutPanel.ColumnCount = 1;
            this.mainTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainTableLayoutPanel.Controls.Add(this.topTableLayoutPanel, 0, 0);
            this.mainTableLayoutPanel.Controls.Add(this.sqlEditorScintilla, 0, 2);
            this.mainTableLayoutPanel.Controls.Add(this.dataGridView, 0, 1);
            this.mainTableLayoutPanel.Controls.Add(this.bottomTableLayoutPanel, 0, 3);
            this.mainTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainTableLayoutPanel.Location = new System.Drawing.Point(0, 0);
            this.mainTableLayoutPanel.Name = "mainTableLayoutPanel";
            this.mainTableLayoutPanel.RowCount = 4;
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 65F));
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.mainTableLayoutPanel.Size = new System.Drawing.Size(863, 826);
            this.mainTableLayoutPanel.TabIndex = 11;
            // 
            // topTableLayoutPanel
            // 
            this.topTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.topTableLayoutPanel.ColumnCount = 6;
            this.topTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.topTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.topTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.topTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.topTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.topTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.66667F));
            this.topTableLayoutPanel.Controls.Add(this.loadFromFilebutton, 1, 0);
            this.topTableLayoutPanel.Controls.Add(this.runButton, 5, 0);
            this.topTableLayoutPanel.Controls.Add(this.runSelectionButton, 4, 0);
            this.topTableLayoutPanel.Controls.Add(this.loadButton, 0, 0);
            this.topTableLayoutPanel.Controls.Add(this.dataHasHeadersCheckBox, 2, 0);
            this.topTableLayoutPanel.Controls.Add(this.dimentionsLabel, 3, 0);
            this.topTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.topTableLayoutPanel.Location = new System.Drawing.Point(3, 3);
            this.topTableLayoutPanel.Name = "topTableLayoutPanel";
            this.topTableLayoutPanel.RowCount = 1;
            this.topTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.topTableLayoutPanel.Size = new System.Drawing.Size(857, 34);
            this.topTableLayoutPanel.TabIndex = 13;
            // 
            // loadFromFilebutton
            // 
            this.loadFromFilebutton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.loadFromFilebutton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.loadFromFilebutton.Location = new System.Drawing.Point(145, 3);
            this.loadFromFilebutton.Name = "loadFromFilebutton";
            this.loadFromFilebutton.Size = new System.Drawing.Size(136, 28);
            this.loadFromFilebutton.TabIndex = 5;
            this.loadFromFilebutton.Text = "Load File";
            this.loadFromFilebutton.UseVisualStyleBackColor = true;
            this.loadFromFilebutton.Click += new System.EventHandler(this.loadFromFilebutton_Click);
            // 
            // runButton
            // 
            this.runButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.runButton.BackColor = System.Drawing.Color.LightCoral;
            this.runButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.runButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.runButton.ForeColor = System.Drawing.Color.DarkRed;
            this.runButton.Location = new System.Drawing.Point(713, 3);
            this.runButton.Name = "runButton";
            this.runButton.Size = new System.Drawing.Size(141, 28);
            this.runButton.TabIndex = 4;
            this.runButton.Text = "Run";
            this.runButton.UseVisualStyleBackColor = false;
            this.runButton.Click += new System.EventHandler(this.runBtn_Click);
            // 
            // runSelectionButton
            // 
            this.runSelectionButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.runSelectionButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.runSelectionButton.Location = new System.Drawing.Point(571, 3);
            this.runSelectionButton.Name = "runSelectionButton";
            this.runSelectionButton.Size = new System.Drawing.Size(136, 28);
            this.runSelectionButton.TabIndex = 3;
            this.runSelectionButton.Text = "Run selection";
            this.runSelectionButton.UseVisualStyleBackColor = true;
            this.runSelectionButton.Click += new System.EventHandler(this.runSelectionBtn_Click);
            // 
            // loadButton
            // 
            this.loadButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.loadButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.loadButton.Location = new System.Drawing.Point(3, 3);
            this.loadButton.Name = "loadButton";
            this.loadButton.Size = new System.Drawing.Size(136, 28);
            this.loadButton.TabIndex = 0;
            this.loadButton.Text = "Load";
            this.loadButton.UseVisualStyleBackColor = true;
            this.loadButton.Click += new System.EventHandler(this.loadButton_Click);
            // 
            // dataHasHeadersCheckBox
            // 
            this.dataHasHeadersCheckBox.AutoSize = true;
            this.dataHasHeadersCheckBox.Checked = true;
            this.dataHasHeadersCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.dataHasHeadersCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataHasHeadersCheckBox.Location = new System.Drawing.Point(287, 3);
            this.dataHasHeadersCheckBox.Name = "dataHasHeadersCheckBox";
            this.dataHasHeadersCheckBox.Size = new System.Drawing.Size(136, 28);
            this.dataHasHeadersCheckBox.TabIndex = 1;
            this.dataHasHeadersCheckBox.Text = "Data has headers";
            this.dataHasHeadersCheckBox.UseVisualStyleBackColor = true;
            // 
            // dimentionsLabel
            // 
            this.dimentionsLabel.AutoSize = true;
            this.dimentionsLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dimentionsLabel.Location = new System.Drawing.Point(429, 0);
            this.dimentionsLabel.Name = "dimentionsLabel";
            this.dimentionsLabel.Size = new System.Drawing.Size(136, 34);
            this.dimentionsLabel.TabIndex = 2;
            this.dimentionsLabel.Text = "Dimentions";
            this.dimentionsLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToOrderColumns = true;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView.Location = new System.Drawing.Point(3, 43);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(857, 255);
            this.dataGridView.TabIndex = 11;
            // 
            // bottomTableLayoutPanel
            // 
            this.bottomTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.bottomTableLayoutPanel.ColumnCount = 6;
            this.bottomTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.bottomTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.bottomTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.bottomTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.bottomTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.bottomTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.bottomTableLayoutPanel.Controls.Add(this.commentBtn, 1, 0);
            this.bottomTableLayoutPanel.Controls.Add(this.wrapIntoBlockBtn, 2, 0);
            this.bottomTableLayoutPanel.Controls.Add(this.closeButton, 5, 0);
            this.bottomTableLayoutPanel.Controls.Add(this.pasteButton, 3, 0);
            this.bottomTableLayoutPanel.Controls.Add(this.validateButton, 0, 0);
            this.bottomTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bottomTableLayoutPanel.Location = new System.Drawing.Point(3, 788);
            this.bottomTableLayoutPanel.Name = "bottomTableLayoutPanel";
            this.bottomTableLayoutPanel.RowCount = 1;
            this.bottomTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.bottomTableLayoutPanel.Size = new System.Drawing.Size(857, 35);
            this.bottomTableLayoutPanel.TabIndex = 12;
            // 
            // commentBtn
            // 
            this.commentBtn.AutoSize = true;
            this.commentBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.commentBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commentBtn.Location = new System.Drawing.Point(174, 3);
            this.commentBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.commentBtn.Name = "commentBtn";
            this.commentBtn.Size = new System.Drawing.Size(79, 29);
            this.commentBtn.TabIndex = 23;
            this.commentBtn.Text = "- - ...";
            this.commentBtn.UseVisualStyleBackColor = true;
            this.commentBtn.Click += new System.EventHandler(this.commentBtn_Click);
            // 
            // wrapIntoBlockBtn
            // 
            this.wrapIntoBlockBtn.AutoSize = true;
            this.wrapIntoBlockBtn.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.wrapIntoBlockBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wrapIntoBlockBtn.Location = new System.Drawing.Point(259, 3);
            this.wrapIntoBlockBtn.MinimumSize = new System.Drawing.Size(0, 25);
            this.wrapIntoBlockBtn.Name = "wrapIntoBlockBtn";
            this.wrapIntoBlockBtn.Size = new System.Drawing.Size(79, 29);
            this.wrapIntoBlockBtn.TabIndex = 24;
            this.wrapIntoBlockBtn.Text = "( ... )";
            this.wrapIntoBlockBtn.UseVisualStyleBackColor = true;
            this.wrapIntoBlockBtn.Click += new System.EventHandler(this.wrapIntoBlockBtn_Click);
            // 
            // closeButton
            // 
            this.closeButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.closeButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.closeButton.Location = new System.Drawing.Point(686, 3);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(168, 29);
            this.closeButton.TabIndex = 5;
            this.closeButton.Text = "Close";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.cancelBtn_Click);
            // 
            // pasteButton
            // 
            this.pasteButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.bottomTableLayoutPanel.SetColumnSpan(this.pasteButton, 2);
            this.pasteButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pasteButton.Location = new System.Drawing.Point(344, 3);
            this.pasteButton.Name = "pasteButton";
            this.pasteButton.Size = new System.Drawing.Size(336, 29);
            this.pasteButton.TabIndex = 4;
            this.pasteButton.Text = "Paste";
            this.pasteButton.UseVisualStyleBackColor = true;
            this.pasteButton.Click += new System.EventHandler(this.pasteButton_Click);
            // 
            // validateButton
            // 
            this.validateButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.validateButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.validateButton.Location = new System.Drawing.Point(3, 3);
            this.validateButton.Name = "validateButton";
            this.validateButton.Size = new System.Drawing.Size(165, 29);
            this.validateButton.TabIndex = 1;
            this.validateButton.Text = "Validate";
            this.validateButton.UseVisualStyleBackColor = true;
            this.validateButton.Click += new System.EventHandler(this.validateSelectionBtn_Click);
            // 
            // SqlExcelTableForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(863, 826);
            this.Controls.Add(this.mainTableLayoutPanel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SqlExcelTableForm";
            this.Opacity = 0.95D;
            this.Text = "Sql Editor for Excel Tables";
            this.mainTableLayoutPanel.ResumeLayout(false);
            this.topTableLayoutPanel.ResumeLayout(false);
            this.topTableLayoutPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.bottomTableLayoutPanel.ResumeLayout(false);
            this.bottomTableLayoutPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private ScintillaNET.Scintilla sqlEditorScintilla;
        private System.Windows.Forms.TableLayoutPanel mainTableLayoutPanel;
        private System.Windows.Forms.TableLayoutPanel topTableLayoutPanel;
        private System.Windows.Forms.Button loadButton;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.TableLayoutPanel bottomTableLayoutPanel;
        private System.Windows.Forms.Button runButton;
        private System.Windows.Forms.Button runSelectionButton;
        private System.Windows.Forms.CheckBox dataHasHeadersCheckBox;
        private System.Windows.Forms.Label dimentionsLabel;
        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.Button pasteButton;
        private System.Windows.Forms.Button validateButton;
        private System.Windows.Forms.Button commentBtn;
        private System.Windows.Forms.Button wrapIntoBlockBtn;
        private System.Windows.Forms.Button loadFromFilebutton;
    }
}