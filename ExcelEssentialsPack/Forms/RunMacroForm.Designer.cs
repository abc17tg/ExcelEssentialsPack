namespace ExcelEssentials.Forms
{
    partial class RunMacroForm
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
            this.vbaEditorScintilla = new ScintillaNET.Scintilla();
            this.refreshBtn = new System.Windows.Forms.Button();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.macrosListView = new System.Windows.Forms.ListView();
            this.closeBtn = new System.Windows.Forms.Button();
            this.workbookPickComboBox = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // vbaEditorScintilla
            // 
            this.vbaEditorScintilla.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.vbaEditorScintilla.CaretForeColor = System.Drawing.Color.White;
            this.vbaEditorScintilla.Lexer = ScintillaNET.Lexer.Vb;
            this.vbaEditorScintilla.Location = new System.Drawing.Point(396, 29);
            this.vbaEditorScintilla.Name = "vbaEditorScintilla";
            this.vbaEditorScintilla.Size = new System.Drawing.Size(434, 560);
            this.vbaEditorScintilla.TabIndex = 10;
            this.vbaEditorScintilla.TabStop = false;
            this.vbaEditorScintilla.UseTabs = true;
            this.vbaEditorScintilla.WrapIndentMode = ScintillaNET.WrapIndentMode.Indent;
            // 
            // refreshBtn
            // 
            this.refreshBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.refreshBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.refreshBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.refreshBtn.ForeColor = System.Drawing.Color.SpringGreen;
            this.refreshBtn.Location = new System.Drawing.Point(325, 29);
            this.refreshBtn.Name = "refreshBtn";
            this.refreshBtn.Size = new System.Drawing.Size(65, 24);
            this.refreshBtn.TabIndex = 13;
            this.refreshBtn.Text = "Refresh";
            this.refreshBtn.UseVisualStyleBackColor = true;
            this.refreshBtn.Click += new System.EventHandler(this.refreshBtn_Click);
            // 
            // searchTextBox
            // 
            this.searchTextBox.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.searchTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.searchTextBox.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.searchTextBox.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.searchTextBox.Location = new System.Drawing.Point(4, 30);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(315, 23);
            this.searchTextBox.TabIndex = 1;
            this.searchTextBox.Text = "Search";
            this.searchTextBox.WordWrap = false;
            this.searchTextBox.TextChanged += new System.EventHandler(this.searchTextBox_TextChanged);
            this.searchTextBox.Enter += new System.EventHandler(this.searchTextBox_Enter);
            this.searchTextBox.Leave += new System.EventHandler(this.searchTextBox_Leave);
            // 
            // macrosListView
            // 
            this.macrosListView.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.macrosListView.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(50)))), ((int)(((byte)(54)))));
            this.macrosListView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.macrosListView.Cursor = System.Windows.Forms.Cursors.Default;
            this.macrosListView.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.macrosListView.ForeColor = System.Drawing.Color.MediumPurple;
            this.macrosListView.FullRowSelect = true;
            this.macrosListView.GridLines = true;
            this.macrosListView.HideSelection = false;
            this.macrosListView.LabelWrap = false;
            this.macrosListView.Location = new System.Drawing.Point(4, 57);
            this.macrosListView.MultiSelect = false;
            this.macrosListView.Name = "macrosListView";
            this.macrosListView.ShowItemToolTips = true;
            this.macrosListView.Size = new System.Drawing.Size(386, 532);
            this.macrosListView.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.macrosListView.TabIndex = 2;
            this.macrosListView.TileSize = new System.Drawing.Size(280, 35);
            this.macrosListView.UseCompatibleStateImageBehavior = false;
            this.macrosListView.View = System.Windows.Forms.View.Tile;
            this.macrosListView.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.macrosListView_ItemSelectionChanged);
            this.macrosListView.DoubleClick += new System.EventHandler(this.macrosListView_DoubleClick);
            this.macrosListView.KeyUp += new System.Windows.Forms.KeyEventHandler(this.macrosListView_KeyUp);
            this.macrosListView.MouseEnter += new System.EventHandler(this.macrosListView_MouseEnter);
            this.macrosListView.MouseLeave += new System.EventHandler(this.macrosListView_MouseLeave);
            // 
            // closeBtn
            // 
            this.closeBtn.BackColor = System.Drawing.Color.Salmon;
            this.closeBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closeBtn.Location = new System.Drawing.Point(813, 6);
            this.closeBtn.Name = "closeBtn";
            this.closeBtn.Size = new System.Drawing.Size(17, 17);
            this.closeBtn.TabIndex = 14;
            this.closeBtn.TabStop = false;
            this.closeBtn.UseVisualStyleBackColor = false;
            this.closeBtn.Click += new System.EventHandler(this.closeBtn_Click);
            // 
            // workbookPickComboBox
            // 
            this.workbookPickComboBox.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.workbookPickComboBox.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.workbookPickComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.workbookPickComboBox.ForeColor = System.Drawing.Color.PaleGreen;
            this.workbookPickComboBox.FormattingEnabled = true;
            this.workbookPickComboBox.Location = new System.Drawing.Point(4, 6);
            this.workbookPickComboBox.Name = "workbookPickComboBox";
            this.workbookPickComboBox.Size = new System.Drawing.Size(315, 21);
            this.workbookPickComboBox.TabIndex = 15;
            this.workbookPickComboBox.TabStop = false;
            // 
            // RunMacroForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.MediumSpringGreen;
            this.ClientSize = new System.Drawing.Size(836, 601);
            this.Controls.Add(this.workbookPickComboBox);
            this.Controls.Add(this.closeBtn);
            this.Controls.Add(this.refreshBtn);
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.macrosListView);
            this.Controls.Add(this.vbaEditorScintilla);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RunMacroForm";
            this.Opacity = 0.95D;
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "RunMacroForm";
            this.Load += new System.EventHandler(this.RunMacroForm_Load);
            this.MouseClick += new System.Windows.Forms.MouseEventHandler(this.RunMacroForm_MouseClick);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ScintillaNET.Scintilla vbaEditorScintilla;
        private System.Windows.Forms.Button refreshBtn;
        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.ListView macrosListView;
        private System.Windows.Forms.Button closeBtn;
        private System.Windows.Forms.ComboBox workbookPickComboBox;
    }
}