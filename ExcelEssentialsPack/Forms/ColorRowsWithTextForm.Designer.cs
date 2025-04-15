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
            ((System.ComponentModel.ISupportInitialize)(this.colorPictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // cancelBtn
            // 
            this.cancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelBtn.Location = new System.Drawing.Point(11, 112);
            this.cancelBtn.Name = "cancelBtn";
            this.cancelBtn.Size = new System.Drawing.Size(96, 27);
            this.cancelBtn.TabIndex = 7;
            this.cancelBtn.Text = "Cancel";
            this.cancelBtn.UseVisualStyleBackColor = true;
            this.cancelBtn.Click += new System.EventHandler(this.cancelBtn_Click);
            // 
            // okBtn
            // 
            this.okBtn.Enabled = false;
            this.okBtn.Location = new System.Drawing.Point(113, 113);
            this.okBtn.Name = "okBtn";
            this.okBtn.Size = new System.Drawing.Size(96, 27);
            this.okBtn.TabIndex = 2;
            this.okBtn.Text = "Accept";
            this.okBtn.UseVisualStyleBackColor = true;
            this.okBtn.Click += new System.EventHandler(this.okBtn_Click);
            // 
            // redLabel
            // 
            this.redLabel.AutoSize = true;
            this.redLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.redLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.redLabel.Location = new System.Drawing.Point(95, 17);
            this.redLabel.Name = "redLabel";
            this.redLabel.Size = new System.Drawing.Size(30, 13);
            this.redLabel.TabIndex = 4;
            this.redLabel.Text = "Red";
            this.redLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // greenLabel
            // 
            this.greenLabel.AutoSize = true;
            this.greenLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.greenLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.greenLabel.Location = new System.Drawing.Point(132, 17);
            this.greenLabel.Name = "greenLabel";
            this.greenLabel.Size = new System.Drawing.Size(41, 13);
            this.greenLabel.TabIndex = 5;
            this.greenLabel.Text = "Green";
            this.greenLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // blueLabel
            // 
            this.blueLabel.AutoSize = true;
            this.blueLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.blueLabel.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.blueLabel.Location = new System.Drawing.Point(176, 17);
            this.blueLabel.Name = "blueLabel";
            this.blueLabel.Size = new System.Drawing.Size(32, 13);
            this.blueLabel.TabIndex = 6;
            this.blueLabel.Text = "Blue";
            this.blueLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // redTextBox
            // 
            this.redTextBox.Location = new System.Drawing.Point(94, 34);
            this.redTextBox.Name = "redTextBox";
            this.redTextBox.Size = new System.Drawing.Size(32, 20);
            this.redTextBox.TabIndex = 3;
            this.redTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.redTextBox.TextChanged += new System.EventHandler(this.redTextBox_TextChanged);
            // 
            // greenTextBox
            // 
            this.greenTextBox.Location = new System.Drawing.Point(135, 34);
            this.greenTextBox.Name = "greenTextBox";
            this.greenTextBox.Size = new System.Drawing.Size(32, 20);
            this.greenTextBox.TabIndex = 4;
            this.greenTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.greenTextBox.TextChanged += new System.EventHandler(this.greenTextBox_TextChanged);
            // 
            // blueTextBox
            // 
            this.blueTextBox.Location = new System.Drawing.Point(176, 34);
            this.blueTextBox.Name = "blueTextBox";
            this.blueTextBox.Size = new System.Drawing.Size(32, 20);
            this.blueTextBox.TabIndex = 5;
            this.blueTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.blueTextBox.TextChanged += new System.EventHandler(this.blueTextBox_TextChanged);
            // 
            // colorPictureBox
            // 
            this.colorPictureBox.Location = new System.Drawing.Point(12, 9);
            this.colorPictureBox.Name = "colorPictureBox";
            this.colorPictureBox.Size = new System.Drawing.Size(76, 71);
            this.colorPictureBox.TabIndex = 10;
            this.colorPictureBox.TabStop = false;
            this.colorPictureBox.Click += new System.EventHandler(this.colorPictureBox_Click);
            // 
            // searchWordTextBox
            // 
            this.searchWordTextBox.Location = new System.Drawing.Point(11, 86);
            this.searchWordTextBox.Name = "searchWordTextBox";
            this.searchWordTextBox.Size = new System.Drawing.Size(198, 20);
            this.searchWordTextBox.TabIndex = 1;
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
            this.invertFontColorCheckBox.Location = new System.Drawing.Point(95, 63);
            this.invertFontColorCheckBox.Name = "invertFontColorCheckBox";
            this.invertFontColorCheckBox.Size = new System.Drawing.Size(100, 17);
            this.invertFontColorCheckBox.TabIndex = 6;
            this.invertFontColorCheckBox.Text = "Invert font color";
            this.invertFontColorCheckBox.UseVisualStyleBackColor = true;
            // 
            // ColorCellsWithTextForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(220, 147);
            this.Controls.Add(this.invertFontColorCheckBox);
            this.Controls.Add(this.searchWordTextBox);
            this.Controls.Add(this.colorPictureBox);
            this.Controls.Add(this.blueTextBox);
            this.Controls.Add(this.greenTextBox);
            this.Controls.Add(this.redTextBox);
            this.Controls.Add(this.blueLabel);
            this.Controls.Add(this.greenLabel);
            this.Controls.Add(this.redLabel);
            this.Controls.Add(this.cancelBtn);
            this.Controls.Add(this.okBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ColorCellsWithTextForm";
            this.ShowIcon = false;
            this.Text = "Pick color of rows and text";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.colorPictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

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
    }
}