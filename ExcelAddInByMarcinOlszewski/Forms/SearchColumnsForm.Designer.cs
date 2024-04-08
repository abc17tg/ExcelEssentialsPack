namespace ExcelAddInByMarcinOlszewski.Forms
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.searchContentsCheckBox = new System.Windows.Forms.CheckBox();
            this.searchBtn = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.ResultColumn = new System.Windows.Forms.DataGridViewButtonColumn();
            this.CountColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Select = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Clear = new System.Windows.Forms.DataGridViewButtonColumn();
            this.cancelBtn = new System.Windows.Forms.Button();
            this.okBtn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // searchTextBox
            // 
            this.searchTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.searchTextBox.Location = new System.Drawing.Point(5, 9);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(232, 21);
            this.searchTextBox.TabIndex = 0;
            this.searchTextBox.Text = "Search";
            this.searchTextBox.WordWrap = false;
            this.searchTextBox.TextChanged += new System.EventHandler(this.searchTextBox_TextChanged);
            // 
            // searchContentsCheckBox
            // 
            this.searchContentsCheckBox.AutoSize = true;
            this.searchContentsCheckBox.Location = new System.Drawing.Point(276, 5);
            this.searchContentsCheckBox.Name = "searchContentsCheckBox";
            this.searchContentsCheckBox.Size = new System.Drawing.Size(67, 30);
            this.searchContentsCheckBox.TabIndex = 1;
            this.searchContentsCheckBox.Text = "Search\r\ncontents";
            this.searchContentsCheckBox.UseVisualStyleBackColor = true;
            this.searchContentsCheckBox.CheckedChanged += new System.EventHandler(this.searchContentsCheckBox_CheckedChanged);
            // 
            // searchBtn
            // 
            this.searchBtn.Enabled = false;
            this.searchBtn.Location = new System.Drawing.Point(243, 8);
            this.searchBtn.Name = "searchBtn";
            this.searchBtn.Size = new System.Drawing.Size(27, 23);
            this.searchBtn.TabIndex = 2;
            this.searchBtn.Text = "🔍";
            this.searchBtn.UseVisualStyleBackColor = true;
            this.searchBtn.Click += new System.EventHandler(this.searchBtn_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.AllowUserToResizeRows = false;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ResultColumn,
            this.CountColumn,
            this.Select,
            this.Clear});
            this.dataGridView.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridView.Location = new System.Drawing.Point(4, 37);
            this.dataGridView.MultiSelect = false;
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView.Size = new System.Drawing.Size(483, 438);
            this.dataGridView.TabIndex = 4;
            this.dataGridView.TabStop = false;
            this.dataGridView.SelectionChanged += new System.EventHandler(this.ColumnSelect);
            // 
            // ResultColumn
            // 
            this.ResultColumn.HeaderText = "ID";
            this.ResultColumn.Name = "ResultColumn";
            this.ResultColumn.ReadOnly = true;
            this.ResultColumn.Width = 200;
            // 
            // CountColumn
            // 
            dataGridViewCellStyle6.Format = "N0";
            dataGridViewCellStyle6.NullValue = null;
            this.CountColumn.DefaultCellStyle = dataGridViewCellStyle6;
            this.CountColumn.HeaderText = "Count";
            this.CountColumn.Name = "CountColumn";
            this.CountColumn.ReadOnly = true;
            this.CountColumn.Width = 80;
            // 
            // Select
            // 
            this.Select.HeaderText = "Select with colour";
            this.Select.Name = "Select";
            this.Select.Width = 80;
            // 
            // Clear
            // 
            this.Clear.HeaderText = "Clear contents";
            this.Clear.Name = "Clear";
            this.Clear.Width = 80;
            // 
            // cancelBtn
            // 
            this.cancelBtn.Location = new System.Drawing.Point(331, 481);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(75, 23);
            this.cancelBtn.TabIndex = 5;
            this.cancelBtn.Text = "Cancel";
            this.cancelBtn.UseVisualStyleBackColor = true;
            // 
            // okBtn
            // 
            this.okBtn.Location = new System.Drawing.Point(412, 481);
            this.okBtn.Name = "okBtn";
            this.okBtn.Size = new System.Drawing.Size(75, 23);
            this.okBtn.TabIndex = 6;
            this.okBtn.Text = "Ok";
            this.okBtn.UseVisualStyleBackColor = true;
            // 
            // SearchColumnsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(491, 512);
            this.Controls.Add(this.okBtn);
            this.Controls.Add(this.cancelBtn);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.searchBtn);
            this.Controls.Add(this.searchContentsCheckBox);
            this.Controls.Add(this.searchTextBox);
            this.Name = "SearchColumnsForm";
            this.Text = "SearchColumnsForm";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.CheckBox searchContentsCheckBox;
        private System.Windows.Forms.Button searchBtn;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.DataGridViewButtonColumn ResultColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn CountColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Select;
        private System.Windows.Forms.DataGridViewButtonColumn Clear;
        private System.Windows.Forms.Button cancelBtn;
        private System.Windows.Forms.Button okBtn;
    }
}