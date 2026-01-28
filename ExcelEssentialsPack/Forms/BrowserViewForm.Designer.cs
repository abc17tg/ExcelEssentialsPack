namespace ExcelEssentials.Forms
{
    partial class BrowserViewForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BrowserViewForm));
            this.webView2 = new Microsoft.Web.WebView2.WinForms.WebView2();
            this.urlTextBox = new System.Windows.Forms.TextBox();
            this.searchButton = new System.Windows.Forms.Button();
            this.saveButton = new System.Windows.Forms.Button();
            this.bookmarksComboBox = new System.Windows.Forms.ComboBox();
            this.addressTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            ((System.ComponentModel.ISupportInitialize)(this.webView2)).BeginInit();
            this.addressTableLayoutPanel.SuspendLayout();
            this.tableLayoutPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // webView2
            // 
            this.webView2.AllowExternalDrop = true;
            this.webView2.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.webView2.CreationProperties = null;
            this.webView2.DefaultBackgroundColor = System.Drawing.Color.Black;
            this.webView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webView2.Location = new System.Drawing.Point(3, 39);
            this.webView2.Name = "webView2";
            this.webView2.Size = new System.Drawing.Size(1258, 769);
            this.webView2.TabIndex = 1;
            this.webView2.ZoomFactor = 0.8D;
            this.webView2.KeyUp += new System.Windows.Forms.KeyEventHandler(this.webView21_KeyUp);
            // 
            // urlTextBox
            // 
            this.urlTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.urlTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.urlTextBox.Location = new System.Drawing.Point(3, 3);
            this.urlTextBox.Name = "urlTextBox";
            this.urlTextBox.Size = new System.Drawing.Size(874, 22);
            this.urlTextBox.TabIndex = 2;
            this.urlTextBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.urlTextBox_KeyDown);
            // 
            // searchButton
            // 
            this.searchButton.AutoSize = true;
            this.searchButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.searchButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.searchButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.searchButton.Location = new System.Drawing.Point(1172, 3);
            this.searchButton.Name = "searchButton";
            this.searchButton.Size = new System.Drawing.Size(83, 24);
            this.searchButton.TabIndex = 3;
            this.searchButton.Text = "Search";
            this.searchButton.UseVisualStyleBackColor = true;
            this.searchButton.Click += new System.EventHandler(this.searchButton_Click);
            // 
            // saveButton
            // 
            this.saveButton.AutoSize = true;
            this.saveButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.saveButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.saveButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.saveButton.Location = new System.Drawing.Point(1084, 3);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(82, 24);
            this.saveButton.TabIndex = 98;
            this.saveButton.Text = "Save";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // bookmarksComboBox
            // 
            this.bookmarksComboBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bookmarksComboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bookmarksComboBox.FormattingEnabled = true;
            this.bookmarksComboBox.Location = new System.Drawing.Point(883, 3);
            this.bookmarksComboBox.Name = "bookmarksComboBox";
            this.bookmarksComboBox.Size = new System.Drawing.Size(195, 23);
            this.bookmarksComboBox.TabIndex = 4;
            this.bookmarksComboBox.SelectedIndexChanged += new System.EventHandler(this.bookmarksComboBox_SelectedIndexChanged);
            // 
            // addressTableLayoutPanel
            // 
            this.addressTableLayoutPanel.AutoSize = true;
            this.addressTableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.addressTableLayoutPanel.ColumnCount = 4;
            this.addressTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.addressTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16F));
            this.addressTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 7F));
            this.addressTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 7F));
            this.addressTableLayoutPanel.Controls.Add(this.bookmarksComboBox, 1, 0);
            this.addressTableLayoutPanel.Controls.Add(this.searchButton, 3, 0);
            this.addressTableLayoutPanel.Controls.Add(this.urlTextBox, 0, 0);
            this.addressTableLayoutPanel.Controls.Add(this.saveButton, 2, 0);
            this.addressTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.addressTableLayoutPanel.Location = new System.Drawing.Point(3, 3);
            this.addressTableLayoutPanel.MinimumSize = new System.Drawing.Size(0, 30);
            this.addressTableLayoutPanel.Name = "addressTableLayoutPanel";
            this.addressTableLayoutPanel.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.addressTableLayoutPanel.RowCount = 1;
            this.addressTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.addressTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.addressTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.addressTableLayoutPanel.Size = new System.Drawing.Size(1258, 30);
            this.addressTableLayoutPanel.TabIndex = 99;
            // 
            // tableLayoutPanel
            // 
            this.tableLayoutPanel.AutoSize = true;
            this.tableLayoutPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel.ColumnCount = 1;
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel.Controls.Add(this.addressTableLayoutPanel, 0, 0);
            this.tableLayoutPanel.Controls.Add(this.webView2, 0, 1);
            this.tableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel.Name = "tableLayoutPanel";
            this.tableLayoutPanel.RowCount = 2;
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel.Size = new System.Drawing.Size(1264, 811);
            this.tableLayoutPanel.TabIndex = 100;
            // 
            // BrowserViewForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1264, 811);
            this.Controls.Add(this.tableLayoutPanel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "BrowserViewForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Edge";
            this.Load += new System.EventHandler(this.BrowserViewForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.webView2)).EndInit();
            this.addressTableLayoutPanel.ResumeLayout(false);
            this.addressTableLayoutPanel.PerformLayout();
            this.tableLayoutPanel.ResumeLayout(false);
            this.tableLayoutPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Microsoft.Web.WebView2.WinForms.WebView2 webView2;
        private System.Windows.Forms.TextBox urlTextBox;
        private System.Windows.Forms.Button searchButton;
        private System.Windows.Forms.Button saveButton;
        private System.Windows.Forms.ComboBox bookmarksComboBox;
        private System.Windows.Forms.TableLayoutPanel addressTableLayoutPanel;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel;
    }
}