namespace ExcelEssentials.Forms
{
    partial class SearchColumnsForm
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SearchColumnsForm));
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.searchContentsCheckBox = new System.Windows.Forms.CheckBox();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.ResultColumn = new System.Windows.Forms.DataGridViewButtonColumn();
            this.CountColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Select = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Clear = new System.Windows.Forms.DataGridViewButtonColumn();
            this.cancelBtn = new System.Windows.Forms.Button();
            this.okBtn = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.countsCheckBox = new System.Windows.Forms.CheckBox();
            this.useDataTableCheckBox = new System.Windows.Forms.CheckBox();
            this.searchBtn = new System.Windows.Forms.Button();
            this.fetchBtn = new System.Windows.Forms.Button();
            this.useDataTableToolTip = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // searchTextBox
            // 
            this.searchTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.searchTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.searchTextBox.Location = new System.Drawing.Point(3, 8);
            this.searchTextBox.Margin = new System.Windows.Forms.Padding(3, 5, 3, 3);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(311, 21);
            this.searchTextBox.TabIndex = 0;
            this.searchTextBox.WordWrap = false;
            this.searchTextBox.TextChanged += new System.EventHandler(this.searchTextBox_TextChanged);
            this.searchTextBox.KeyUp += new System.Windows.Forms.KeyEventHandler(this.searchTextBox_KeyUp);
            // 
            // searchContentsCheckBox
            // 
            this.searchContentsCheckBox.AutoSize = true;
            this.searchContentsCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.searchContentsCheckBox.Location = new System.Drawing.Point(355, 3);
            this.searchContentsCheckBox.Name = "searchContentsCheckBox";
            this.searchContentsCheckBox.Size = new System.Drawing.Size(107, 29);
            this.searchContentsCheckBox.TabIndex = 1;
            this.searchContentsCheckBox.Text = "Search contents";
            this.searchContentsCheckBox.UseVisualStyleBackColor = true;
            this.searchContentsCheckBox.CheckedChanged += new System.EventHandler(this.searchContentsCheckBox_CheckedChanged);
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.AllowUserToResizeRows = false;
            this.dataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ResultColumn,
            this.CountColumn,
            this.Select,
            this.Clear});
            this.tableLayoutPanel1.SetColumnSpan(this.dataGridView, 5);
            this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView.Location = new System.Drawing.Point(3, 38);
            this.dataGridView.MultiSelect = false;
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView.Size = new System.Drawing.Size(648, 310);
            this.dataGridView.TabIndex = 4;
            this.dataGridView.TabStop = false;
            this.dataGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellClick);
            this.dataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_OnCellValueChanged);
            this.dataGridView.CurrentCellDirtyStateChanged += new System.EventHandler(this.dataGridView_CurrentCellDirtyStateChanged);
            // 
            // ResultColumn
            // 
            this.ResultColumn.FillWeight = 60F;
            this.ResultColumn.HeaderText = "ID";
            this.ResultColumn.MinimumWidth = 100;
            this.ResultColumn.Name = "ResultColumn";
            this.ResultColumn.ReadOnly = true;
            // 
            // CountColumn
            // 
            dataGridViewCellStyle1.Format = "N0";
            dataGridViewCellStyle1.NullValue = null;
            this.CountColumn.DefaultCellStyle = dataGridViewCellStyle1;
            this.CountColumn.FillWeight = 10F;
            this.CountColumn.HeaderText = "Count";
            this.CountColumn.Name = "CountColumn";
            this.CountColumn.ReadOnly = true;
            // 
            // Select
            // 
            this.Select.FillWeight = 10F;
            this.Select.HeaderText = "Select with color";
            this.Select.Name = "Select";
            // 
            // Clear
            // 
            this.Clear.FillWeight = 15F;
            this.Clear.HeaderText = "Clear contents";
            this.Clear.Name = "Clear";
            // 
            // cancelBtn
            // 
            this.cancelBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cancelBtn.Location = new System.Drawing.Point(468, 354);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(67, 24);
            this.cancelBtn.TabIndex = 5;
            this.cancelBtn.Text = "Cancel";
            this.cancelBtn.UseVisualStyleBackColor = true;
            this.cancelBtn.Click += new System.EventHandler(this.cancelBtn_Click);
            // 
            // okBtn
            // 
            this.okBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.okBtn.Location = new System.Drawing.Point(541, 354);
            this.okBtn.Name = "okBtn";
            this.okBtn.Size = new System.Drawing.Size(110, 24);
            this.okBtn.TabIndex = 6;
            this.okBtn.Text = "Ok";
            this.okBtn.UseVisualStyleBackColor = true;
            this.okBtn.Click += new System.EventHandler(this.okBtn_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 5;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 51.3694F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.39679F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 11.83703F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 18.39678F));
            this.tableLayoutPanel1.Controls.Add(this.countsCheckBox, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.useDataTableCheckBox, 4, 0);
            this.tableLayoutPanel1.Controls.Add(this.searchTextBox, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.okBtn, 4, 2);
            this.tableLayoutPanel1.Controls.Add(this.searchBtn, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.cancelBtn, 3, 2);
            this.tableLayoutPanel1.Controls.Add(this.searchContentsCheckBox, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.dataGridView, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.fetchBtn, 0, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(654, 381);
            this.tableLayoutPanel1.TabIndex = 7;
            // 
            // countsCheckBox
            // 
            this.countsCheckBox.AutoSize = true;
            this.countsCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.countsCheckBox.Location = new System.Drawing.Point(468, 3);
            this.countsCheckBox.Name = "countsCheckBox";
            this.countsCheckBox.Size = new System.Drawing.Size(67, 29);
            this.countsCheckBox.TabIndex = 9;
            this.countsCheckBox.Text = "Counts";
            this.countsCheckBox.UseVisualStyleBackColor = true;
            this.countsCheckBox.CheckedChanged += new System.EventHandler(this.countsCheckBox_CheckedChanged);
            // 
            // useDataTableCheckBox
            // 
            this.useDataTableCheckBox.AutoSize = true;
            this.useDataTableCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.useDataTableCheckBox.Location = new System.Drawing.Point(541, 3);
            this.useDataTableCheckBox.Name = "useDataTableCheckBox";
            this.useDataTableCheckBox.Size = new System.Drawing.Size(110, 29);
            this.useDataTableCheckBox.TabIndex = 8;
            this.useDataTableCheckBox.Text = "Use DataTable";
            this.useDataTableCheckBox.UseVisualStyleBackColor = true;
            // 
            // searchBtn
            // 
            this.searchBtn.BackgroundImage = global::ExcelEssentials.Properties.Resources.search_outline_filled;
            this.searchBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.searchBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.searchBtn.Enabled = false;
            this.searchBtn.Location = new System.Drawing.Point(322, 5);
            this.searchBtn.Margin = new System.Windows.Forms.Padding(5);
            this.searchBtn.Name = "searchBtn";
            this.searchBtn.Size = new System.Drawing.Size(25, 25);
            this.searchBtn.TabIndex = 2;
            this.searchBtn.UseVisualStyleBackColor = true;
            this.searchBtn.Click += new System.EventHandler(this.searchBtn_Click);
            // 
            // fetchBtn
            // 
            this.fetchBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fetchBtn.Location = new System.Drawing.Point(3, 354);
            this.fetchBtn.Name = "fetchBtn";
            this.fetchBtn.Size = new System.Drawing.Size(311, 24);
            this.fetchBtn.TabIndex = 7;
            this.fetchBtn.Text = "Fetch";
            this.fetchBtn.UseVisualStyleBackColor = true;
            this.fetchBtn.Click += new System.EventHandler(this.fetchBtn_Click);
            // 
            // useDataTableToolTip
            // 
            this.useDataTableToolTip.BackColor = System.Drawing.SystemColors.InactiveBorder;
            // 
            // SearchColumnsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(654, 381);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SearchColumnsForm";
            this.Opacity = 0.95D;
            this.Text = "Search Columns Form";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.CheckBox searchContentsCheckBox;
        private System.Windows.Forms.Button searchBtn;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.DataGridViewCheckBoxColumn SelectColumn;
        private System.Windows.Forms.Button cancelBtn;
        private System.Windows.Forms.Button okBtn;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button fetchBtn;
        private System.Windows.Forms.CheckBox useDataTableCheckBox;
        private System.Windows.Forms.ToolTip useDataTableToolTip;
        private System.Windows.Forms.CheckBox countsCheckBox;
        private System.Windows.Forms.DataGridViewButtonColumn ResultColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn CountColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn SelectCol;
        private System.Windows.Forms.DataGridViewButtonColumn ClearCol;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Select;
        private System.Windows.Forms.DataGridViewButtonColumn Clear;
    }
}