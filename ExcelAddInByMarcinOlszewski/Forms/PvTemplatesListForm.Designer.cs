namespace ExcelAddInByMarcinOlszewski.Forms
{
    partial class PvTemplatesListForm
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
            this.pvTemplatesListView = new System.Windows.Forms.ListView();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.refreshBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // pvTemplatesListView
            // 
            this.pvTemplatesListView.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.pvTemplatesListView.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(44)))), ((int)(((byte)(50)))), ((int)(((byte)(54)))));
            this.pvTemplatesListView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pvTemplatesListView.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pvTemplatesListView.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pvTemplatesListView.ForeColor = System.Drawing.Color.MediumPurple;
            this.pvTemplatesListView.GridLines = true;
            this.pvTemplatesListView.HideSelection = false;
            this.pvTemplatesListView.LabelWrap = false;
            this.pvTemplatesListView.Location = new System.Drawing.Point(4, 33);
            this.pvTemplatesListView.MultiSelect = false;
            this.pvTemplatesListView.Name = "pvTemplatesListView";
            this.pvTemplatesListView.ShowItemToolTips = true;
            this.pvTemplatesListView.Size = new System.Drawing.Size(290, 520);
            this.pvTemplatesListView.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.pvTemplatesListView.TabIndex = 2;
            this.pvTemplatesListView.TileSize = new System.Drawing.Size(280, 30);
            this.pvTemplatesListView.UseCompatibleStateImageBehavior = false;
            this.pvTemplatesListView.View = System.Windows.Forms.View.Tile;
            this.pvTemplatesListView.DoubleClick += new System.EventHandler(this.pvTemplatesListView_DoubleClick);
            this.pvTemplatesListView.KeyUp += new System.Windows.Forms.KeyEventHandler(this.pvTemplatesListView_KeyUp);
            this.pvTemplatesListView.MouseEnter += new System.EventHandler(this.pvTemplatesListView_MouseEnter);
            this.pvTemplatesListView.MouseLeave += new System.EventHandler(this.pvTemplatesListView_MouseLeave);
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
            this.Controls.Add(this.pvTemplatesListView);
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
            this.Load += new System.EventHandler(this.PvTemplatesListForm_Load);
            this.MouseClick += new System.Windows.Forms.MouseEventHandler(this.PvTemplatesListForm_MouseClick);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView pvTemplatesListView;
        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.Button refreshBtn;
    }
}