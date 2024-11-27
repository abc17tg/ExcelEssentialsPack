namespace ExcelEssentials.Forms
{
    partial class RunMacroWithSearchPhraseForm
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
            this.templatesListView = new System.Windows.Forms.ListView();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.refreshBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // pvTemplatesListView
            // 
            this.templatesListView.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.templatesListView.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(50)))), ((int)(((byte)(54)))));
            this.templatesListView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.templatesListView.Cursor = System.Windows.Forms.Cursors.Hand;
            this.templatesListView.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.templatesListView.ForeColor = System.Drawing.Color.MediumPurple;
            this.templatesListView.GridLines = true;
            this.templatesListView.HideSelection = false;
            this.templatesListView.LabelWrap = false;
            this.templatesListView.Location = new System.Drawing.Point(4, 33);
            this.templatesListView.MultiSelect = false;
            this.templatesListView.Name = "pvTemplatesListView";
            this.templatesListView.ShowItemToolTips = true;
            this.templatesListView.Size = new System.Drawing.Size(290, 520);
            this.templatesListView.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.templatesListView.TabIndex = 2;
            this.templatesListView.TileSize = new System.Drawing.Size(280, 30);
            this.templatesListView.UseCompatibleStateImageBehavior = false;
            this.templatesListView.View = System.Windows.Forms.View.Tile;
            this.templatesListView.DoubleClick += new System.EventHandler(this.templatesListView_DoubleClick);
            this.templatesListView.KeyUp += new System.Windows.Forms.KeyEventHandler(this.templatesListView_KeyUp);
            this.templatesListView.MouseEnter += new System.EventHandler(this.templatesListView_MouseEnter);
            this.templatesListView.MouseLeave += new System.EventHandler(this.templatesListView_MouseLeave);
            // 
            // searchTextBox
            // 
            this.searchTextBox.BackColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.searchTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.searchTextBox.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.searchTextBox.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.searchTextBox.Location = new System.Drawing.Point(4, 6);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(219, 23);
            this.searchTextBox.TabIndex = 1;
            this.searchTextBox.Text = "Search";
            this.searchTextBox.WordWrap = false;
            this.searchTextBox.TextChanged += new System.EventHandler(this.searchTextBox_TextChanged);
            this.searchTextBox.Enter += new System.EventHandler(this.searchTextBox_Enter);
            this.searchTextBox.Leave += new System.EventHandler(this.searchTextBox_Leave);
            // 
            // refreshBtn
            // 
            this.refreshBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.refreshBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.refreshBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.refreshBtn.ForeColor = System.Drawing.Color.SpringGreen;
            this.refreshBtn.Location = new System.Drawing.Point(229, 5);
            this.refreshBtn.Name = "refreshBtn";
            this.refreshBtn.Size = new System.Drawing.Size(65, 23);
            this.refreshBtn.TabIndex = 3;
            this.refreshBtn.Text = "Refresh";
            this.refreshBtn.UseVisualStyleBackColor = true;
            this.refreshBtn.Click += new System.EventHandler(this.refreshBtn_Click);
            // 
            // PvTemplatesListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(300, 560);
            this.Controls.Add(this.refreshBtn);
            this.Controls.Add(this.searchTextBox);
            this.Controls.Add(this.templatesListView);
            this.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PvTemplatesListForm";
            this.Opacity = 0.95D;
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "PvTemplatesListForm";
            this.Load += new System.EventHandler(this.RunMacroWithSearchPhraseForm_Load);
            this.MouseClick += new System.Windows.Forms.MouseEventHandler(this.RunMacroWithSearchPhraseForm_MouseClick);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView templatesListView;
        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.Button refreshBtn;
    }
}