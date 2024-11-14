namespace ExcelAddInByMarcinOlszewski.Forms
{
    partial class FormatToDateForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.rngDataGridView = new System.Windows.Forms.DataGridView();
            this.fetchButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.checkButton = new System.Windows.Forms.Button();
            this.predefinedFormatsComboBox = new System.Windows.Forms.ComboBox();
            this.formatTextBox = new System.Windows.Forms.TextBox();
            this.countLabel = new System.Windows.Forms.Label();
            this.choosePredefinedLabel = new System.Windows.Forms.Label();
            this.parsedDateLabel = new System.Windows.Forms.Label();
            this.inputFormatButton = new System.Windows.Forms.Button();
            this.tableLayoutPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rngDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel
            // 
            this.tableLayoutPanel.ColumnCount = 3;
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.2985F));
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 29.85075F));
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 29.85075F));
            this.tableLayoutPanel.Controls.Add(this.inputFormatButton, 1, 0);
            this.tableLayoutPanel.Controls.Add(this.rngDataGridView, 0, 1);
            this.tableLayoutPanel.Controls.Add(this.fetchButton, 0, 3);
            this.tableLayoutPanel.Controls.Add(this.cancelButton, 1, 3);
            this.tableLayoutPanel.Controls.Add(this.okButton, 2, 3);
            this.tableLayoutPanel.Controls.Add(this.checkButton, 1, 2);
            this.tableLayoutPanel.Controls.Add(this.predefinedFormatsComboBox, 2, 1);
            this.tableLayoutPanel.Controls.Add(this.formatTextBox, 1, 1);
            this.tableLayoutPanel.Controls.Add(this.countLabel, 0, 0);
            this.tableLayoutPanel.Controls.Add(this.choosePredefinedLabel, 2, 0);
            this.tableLayoutPanel.Controls.Add(this.parsedDateLabel, 2, 2);
            this.tableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel.Name = "tableLayoutPanel";
            this.tableLayoutPanel.RowCount = 4;
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel.Size = new System.Drawing.Size(495, 242);
            this.tableLayoutPanel.TabIndex = 0;
            // 
            // rngDataGridView
            // 
            this.rngDataGridView.AllowUserToAddRows = false;
            this.rngDataGridView.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.rngDataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.rngDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.rngDataGridView.DefaultCellStyle = dataGridViewCellStyle2;
            this.rngDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rngDataGridView.Location = new System.Drawing.Point(3, 33);
            this.rngDataGridView.Name = "rngDataGridView";
            this.rngDataGridView.ReadOnly = true;
            this.tableLayoutPanel.SetRowSpan(this.rngDataGridView, 2);
            this.rngDataGridView.Size = new System.Drawing.Size(193, 176);
            this.rngDataGridView.TabIndex = 0;
            // 
            // fetchButton
            // 
            this.fetchButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.fetchButton.Location = new System.Drawing.Point(3, 215);
            this.fetchButton.Name = "fetchButton";
            this.fetchButton.Size = new System.Drawing.Size(193, 24);
            this.fetchButton.TabIndex = 1;
            this.fetchButton.Text = "Fetch";
            this.fetchButton.UseVisualStyleBackColor = true;
            this.fetchButton.Click += new System.EventHandler(this.fetchButton_Click);
            // 
            // cancelButton
            // 
            this.cancelButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cancelButton.Location = new System.Drawing.Point(202, 215);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(141, 24);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // okButton
            // 
            this.okButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.okButton.Location = new System.Drawing.Point(349, 215);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(143, 24);
            this.okButton.TabIndex = 3;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // checkButton
            // 
            this.checkButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checkButton.Location = new System.Drawing.Point(202, 185);
            this.checkButton.Name = "checkButton";
            this.checkButton.Size = new System.Drawing.Size(141, 24);
            this.checkButton.TabIndex = 4;
            this.checkButton.Text = "Check";
            this.checkButton.UseVisualStyleBackColor = true;
            this.checkButton.Click += new System.EventHandler(this.checkButton_Click);
            // 
            // predefinedFormatsComboBox
            // 
            this.predefinedFormatsComboBox.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.predefinedFormatsComboBox.FormattingEnabled = true;
            this.predefinedFormatsComboBox.Location = new System.Drawing.Point(349, 158);
            this.predefinedFormatsComboBox.Name = "predefinedFormatsComboBox";
            this.predefinedFormatsComboBox.Size = new System.Drawing.Size(143, 21);
            this.predefinedFormatsComboBox.TabIndex = 5;
            this.predefinedFormatsComboBox.SelectedIndexChanged += new System.EventHandler(this.predefinedFormatsComboBox_SelectedIndexChanged);
            // 
            // formatTextBox
            // 
            this.formatTextBox.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.formatTextBox.Location = new System.Drawing.Point(202, 159);
            this.formatTextBox.Name = "formatTextBox";
            this.formatTextBox.Size = new System.Drawing.Size(141, 20);
            this.formatTextBox.TabIndex = 6;
            this.formatTextBox.TextChanged += new System.EventHandler(this.formatTextBox_TextChanged);
            // 
            // countLabel
            // 
            this.countLabel.AutoSize = true;
            this.countLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.countLabel.Location = new System.Drawing.Point(3, 0);
            this.countLabel.Name = "countLabel";
            this.countLabel.Size = new System.Drawing.Size(193, 30);
            this.countLabel.TabIndex = 7;
            this.countLabel.Text = "Cell count";
            this.countLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // choosePredefinedLabel
            // 
            this.choosePredefinedLabel.AutoSize = true;
            this.choosePredefinedLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.choosePredefinedLabel.Location = new System.Drawing.Point(349, 0);
            this.choosePredefinedLabel.Name = "choosePredefinedLabel";
            this.choosePredefinedLabel.Size = new System.Drawing.Size(143, 30);
            this.choosePredefinedLabel.TabIndex = 9;
            this.choosePredefinedLabel.Text = "Choose predefined formats";
            this.choosePredefinedLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // parsedDateLabel
            // 
            this.parsedDateLabel.AutoSize = true;
            this.parsedDateLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.parsedDateLabel.Location = new System.Drawing.Point(349, 182);
            this.parsedDateLabel.Name = "parsedDateLabel";
            this.parsedDateLabel.Size = new System.Drawing.Size(143, 30);
            this.parsedDateLabel.TabIndex = 10;
            this.parsedDateLabel.Text = "Parsed date";
            this.parsedDateLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // inputFormatButton
            // 
            this.inputFormatButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.inputFormatButton.Location = new System.Drawing.Point(202, 3);
            this.inputFormatButton.Name = "inputFormatButton";
            this.inputFormatButton.Size = new System.Drawing.Size(141, 24);
            this.inputFormatButton.TabIndex = 11;
            this.inputFormatButton.Text = "Input format of date cells";
            this.inputFormatButton.UseVisualStyleBackColor = true;
            this.inputFormatButton.Click += new System.EventHandler(this.inputFormatButton_Click);
            // 
            // FormatToDateForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(495, 242);
            this.Controls.Add(this.tableLayoutPanel);
            this.Name = "FormatToDateForm";
            this.Text = "String to date";
            this.tableLayoutPanel.ResumeLayout(false);
            this.tableLayoutPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rngDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel;
        private System.Windows.Forms.DataGridView rngDataGridView;
        private System.Windows.Forms.Button fetchButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button checkButton;
        private System.Windows.Forms.ComboBox predefinedFormatsComboBox;
        private System.Windows.Forms.TextBox formatTextBox;
        private System.Windows.Forms.Label countLabel;
        private System.Windows.Forms.Label choosePredefinedLabel;
        private System.Windows.Forms.Label parsedDateLabel;
        private System.Windows.Forms.Button inputFormatButton;
    }
}