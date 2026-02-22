namespace ExcelEssentials.Forms
{
    partial class ColorCellsWithTextForm
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
            this.cancelBtn = new System.Windows.Forms.Button();
            this.okBtn = new System.Windows.Forms.Button();
            this.redLabel = new System.Windows.Forms.Label();
            this.greenLabel = new System.Windows.Forms.Label();
            this.blueLabel = new System.Windows.Forms.Label();
            this.redTextBox = new System.Windows.Forms.TextBox();
            this.greenTextBox = new System.Windows.Forms.TextBox();
            this.blueTextBox = new System.Windows.Forms.TextBox();
            this.colorPictureBox = new System.Windows.Forms.PictureBox();
            this.searchWordTextBox = new System.Windows.Forms.TextBox();
            this.colorDialog = new System.Windows.Forms.ColorDialog();
            this.invertFontColorCheckBox = new System.Windows.Forms.CheckBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.exactCheckBox = new System.Windows.Forms.CheckBox();
            this.extStringCheckBox = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.colorPictureBox)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cancelBtn
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.cancelBtn, 2);
            this.cancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cancelBtn.Location = new System.Drawing.Point(3, 119);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(162, 29);
            this.cancelBtn.TabIndex = 2;
            this.cancelBtn.Text = "Cancel";
            this.cancelBtn.UseVisualStyleBackColor = true;
            this.cancelBtn.Click += new System.EventHandler(this.cancelBtn_Click);
            // 
            // okBtn
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.okBtn, 2);
            this.okBtn.Dock = System.Windows.Forms.DockStyle.Fill;
            this.okBtn.Enabled = false;
            this.okBtn.Location = new System.Drawing.Point(171, 119);
            this.okBtn.Name = "okBtn";
            this.okBtn.Size = new System.Drawing.Size(150, 29);
            this.okBtn.TabIndex = 1;
            this.okBtn.Text = "Accept";
            this.okBtn.UseVisualStyleBackColor = true;
            this.okBtn.Click += new System.EventHandler(this.okBtn_Click);
            // 
            // redLabel
            // 
            this.redLabel.AutoSize = true;
            this.redLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.redLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.redLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.redLabel.Location = new System.Drawing.Point(93, 0);
            this.redLabel.Name = "redLabel";
            this.redLabel.Size = new System.Drawing.Size(72, 30);
            this.redLabel.TabIndex = 4;
            this.redLabel.Text = "Red";
            this.redLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // greenLabel
            // 
            this.greenLabel.AutoSize = true;
            this.greenLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.greenLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.greenLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.greenLabel.Location = new System.Drawing.Point(171, 0);
            this.greenLabel.Name = "greenLabel";
            this.greenLabel.Size = new System.Drawing.Size(72, 30);
            this.greenLabel.TabIndex = 5;
            this.greenLabel.Text = "Green";
            this.greenLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // blueLabel
            // 
            this.blueLabel.AutoSize = true;
            this.blueLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.blueLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.blueLabel.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.blueLabel.Location = new System.Drawing.Point(249, 0);
            this.blueLabel.Name = "blueLabel";
            this.blueLabel.Size = new System.Drawing.Size(72, 30);
            this.blueLabel.TabIndex = 6;
            this.blueLabel.Text = "Blue";
            this.blueLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // redTextBox
            // 
            this.redTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.redTextBox.Location = new System.Drawing.Point(93, 33);
            this.redTextBox.Name = "redTextBox";
            this.redTextBox.Size = new System.Drawing.Size(72, 20);
            this.redTextBox.TabIndex = 5;
            this.redTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.redTextBox.TextChanged += new System.EventHandler(this.redTextBox_TextChanged);
            // 
            // greenTextBox
            // 
            this.greenTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.greenTextBox.Location = new System.Drawing.Point(171, 33);
            this.greenTextBox.Name = "greenTextBox";
            this.greenTextBox.Size = new System.Drawing.Size(72, 20);
            this.greenTextBox.TabIndex = 6;
            this.greenTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.greenTextBox.TextChanged += new System.EventHandler(this.greenTextBox_TextChanged);
            // 
            // blueTextBox
            // 
            this.blueTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.blueTextBox.Location = new System.Drawing.Point(249, 33);
            this.blueTextBox.Name = "blueTextBox";
            this.blueTextBox.Size = new System.Drawing.Size(72, 20);
            this.blueTextBox.TabIndex = 7;
            this.blueTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.blueTextBox.TextChanged += new System.EventHandler(this.blueTextBox_TextChanged);
            // 
            // colorPictureBox
            // 
            this.colorPictureBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.colorPictureBox.Location = new System.Drawing.Point(3, 3);
            this.colorPictureBox.Name = "colorPictureBox";
            this.tableLayoutPanel1.SetRowSpan(this.colorPictureBox, 3);
            this.colorPictureBox.Size = new System.Drawing.Size(84, 84);
            this.colorPictureBox.TabIndex = 10;
            this.colorPictureBox.TabStop = false;
            this.colorPictureBox.Click += new System.EventHandler(this.colorPictureBox_Click);
            // 
            // searchWordTextBox
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.searchWordTextBox, 3);
            this.searchWordTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.searchWordTextBox.Location = new System.Drawing.Point(3, 93);
            this.searchWordTextBox.Name = "searchWordTextBox";
            this.searchWordTextBox.Size = new System.Drawing.Size(240, 20);
            this.searchWordTextBox.TabIndex = 0;
            this.searchWordTextBox.Text = "Search word";
            this.searchWordTextBox.TextChanged += new System.EventHandler(this.searchWordTextBox_TextChanged);
            this.searchWordTextBox.Enter += new System.EventHandler(this.searchWordTextBox_Enter);
            this.searchWordTextBox.KeyUp += new System.Windows.Forms.KeyEventHandler(this.searchWordTextBox_KeyUp);
            this.searchWordTextBox.Leave += new System.EventHandler(this.searchWordTextBox_Leave);
            // 
            // colorDialog
            // 
            this.colorDialog.AnyColor = true;
            this.colorDialog.Color = System.Drawing.Color.PaleGreen;
            this.colorDialog.SolidColorOnly = true;
            // 
            // invertFontColorCheckBox
            // 
            this.invertFontColorCheckBox.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.invertFontColorCheckBox, 2);
            this.invertFontColorCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.invertFontColorCheckBox.Location = new System.Drawing.Point(93, 63);
            this.invertFontColorCheckBox.Name = "invertFontColorCheckBox";
            this.invertFontColorCheckBox.Size = new System.Drawing.Size(150, 24);
            this.invertFontColorCheckBox.TabIndex = 3;
            this.invertFontColorCheckBox.Text = "Invert font color";
            this.invertFontColorCheckBox.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 90F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.Controls.Add(this.colorPictureBox, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.searchWordTextBox, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.invertFontColorCheckBox, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.okBtn, 2, 4);
            this.tableLayoutPanel1.Controls.Add(this.cancelBtn, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.redTextBox, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.greenTextBox, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.blueLabel, 3, 0);
            this.tableLayoutPanel1.Controls.Add(this.redLabel, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.greenLabel, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.blueTextBox, 3, 1);
            this.tableLayoutPanel1.Controls.Add(this.exactCheckBox, 3, 2);
            this.tableLayoutPanel1.Controls.Add(this.extStringCheckBox, 3, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(324, 151);
            this.tableLayoutPanel1.TabIndex = 11;
            // 
            // exactCheckBox
            // 
            this.exactCheckBox.AutoSize = true;
            this.exactCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.exactCheckBox.Location = new System.Drawing.Point(249, 63);
            this.exactCheckBox.Name = "exactCheckBox";
            this.exactCheckBox.Size = new System.Drawing.Size(72, 24);
            this.exactCheckBox.TabIndex = 4;
            this.exactCheckBox.Text = "Exact";
            this.exactCheckBox.UseVisualStyleBackColor = true;
            this.exactCheckBox.CheckedChanged += new System.EventHandler(this.exactCheckBox_CheckedChanged);
            // 
            // extStringCheckBox
            // 
            this.extStringCheckBox.AutoSize = true;
            this.extStringCheckBox.Checked = true;
            this.extStringCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.extStringCheckBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.extStringCheckBox.Font = new System.Drawing.Font("Consolas", 9F);
            this.extStringCheckBox.Location = new System.Drawing.Point(249, 93);
            this.extStringCheckBox.Name = "extStringCheckBox";
            this.extStringCheckBox.Size = new System.Drawing.Size(72, 20);
            this.extStringCheckBox.TabIndex = 11;
            this.extStringCheckBox.Text = "\\t\\n";
            this.extStringCheckBox.UseVisualStyleBackColor = true;
            this.extStringCheckBox.CheckedChanged += new System.EventHandler(this.extStringCheckBox_CheckedChanged);
            // 
            // ColorCellsWithTextForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(324, 151);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ColorCellsWithTextForm";
            this.ShowIcon = false;
            this.Text = "Pick color of rows and text";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.colorPictureBox)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cancelBtn;
        private System.Windows.Forms.Button okBtn;
        private System.Windows.Forms.Label redLabel;
        private System.Windows.Forms.Label greenLabel;
        private System.Windows.Forms.Label blueLabel;
        private System.Windows.Forms.TextBox redTextBox;
        private System.Windows.Forms.TextBox greenTextBox;
        private System.Windows.Forms.TextBox blueTextBox;
        private System.Windows.Forms.PictureBox colorPictureBox;
        private System.Windows.Forms.TextBox searchWordTextBox;
        private System.Windows.Forms.ColorDialog colorDialog;
        private System.Windows.Forms.CheckBox invertFontColorCheckBox;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.CheckBox exactCheckBox;
        private System.Windows.Forms.CheckBox extStringCheckBox;
    }
}