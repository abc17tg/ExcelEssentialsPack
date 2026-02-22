using System;
using System.Windows.Forms;
using System.Drawing;

namespace ExcelEssentials.Forms
{
    partial class RenameForm
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
            this.mainTableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.findLabel = new System.Windows.Forms.Label();
            this.findTextBox = new System.Windows.Forms.TextBox();
            this.replaceLabel = new System.Windows.Forms.Label();
            this.replaceTextBox = new System.Windows.Forms.TextBox();
            this.previewButton = new System.Windows.Forms.Button();
            this.regexCheckBox = new System.Windows.Forms.CheckBox();
            this.cancelButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.previewPanel = new System.Windows.Forms.Panel();
            this.previewDataGridView = new System.Windows.Forms.DataGridView();
            this.beforeColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.afterColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mainTableLayoutPanel.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.previewPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.previewDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // mainTableLayoutPanel
            // 
            this.mainTableLayoutPanel.ColumnCount = 1;
            this.mainTableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainTableLayoutPanel.Controls.Add(this.tableLayoutPanel1, 0, 0);
            this.mainTableLayoutPanel.Controls.Add(this.previewPanel, 0, 1);
            this.mainTableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainTableLayoutPanel.Location = new System.Drawing.Point(0, 0);
            this.mainTableLayoutPanel.Name = "mainTableLayoutPanel";
            this.mainTableLayoutPanel.RowCount = 2;
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 125F));
            this.mainTableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 200F));
            this.mainTableLayoutPanel.Size = new System.Drawing.Size(334, 131);
            this.mainTableLayoutPanel.TabIndex = 0;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.tableLayoutPanel1.Controls.Add(this.findLabel, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.findTextBox, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.replaceLabel, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.replaceTextBox, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.previewButton, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.regexCheckBox, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.cancelButton, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.okButton, 1, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(328, 119);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // findLabel
            // 
            this.findLabel.AutoSize = true;
            this.findLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.findLabel.Location = new System.Drawing.Point(3, 3);
            this.findLabel.Margin = new System.Windows.Forms.Padding(3);
            this.findLabel.Name = "findLabel";
            this.findLabel.Size = new System.Drawing.Size(92, 23);
            this.findLabel.TabIndex = 0;
            this.findLabel.Text = "Find:";
            this.findLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // findTextBox
            // 
            this.findTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.findTextBox.Location = new System.Drawing.Point(101, 4);
            this.findTextBox.Name = "findTextBox";
            this.findTextBox.Size = new System.Drawing.Size(224, 20);
            this.findTextBox.TabIndex = 1;
            // 
            // replaceLabel
            // 
            this.replaceLabel.AutoSize = true;
            this.replaceLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.replaceLabel.Location = new System.Drawing.Point(3, 32);
            this.replaceLabel.Margin = new System.Windows.Forms.Padding(3);
            this.replaceLabel.Name = "replaceLabel";
            this.replaceLabel.Size = new System.Drawing.Size(92, 23);
            this.replaceLabel.TabIndex = 2;
            this.replaceLabel.Text = "Replace with:";
            this.replaceLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // replaceTextBox
            // 
            this.replaceTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.replaceTextBox.Location = new System.Drawing.Point(101, 33);
            this.replaceTextBox.Name = "replaceTextBox";
            this.replaceTextBox.Size = new System.Drawing.Size(224, 20);
            this.replaceTextBox.TabIndex = 3;
            // 
            // previewButton
            // 
            this.previewButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.previewButton.Location = new System.Drawing.Point(3, 61);
            this.previewButton.Name = "previewButton";
            this.previewButton.Size = new System.Drawing.Size(92, 23);
            this.previewButton.TabIndex = 5;
            this.previewButton.Text = "Preview ▼";
            this.previewButton.UseVisualStyleBackColor = true;
            this.previewButton.Click += new System.EventHandler(this.previewButton_Click);
            // 
            // regexCheckBox
            // 
            this.regexCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.regexCheckBox.AutoSize = true;
            this.regexCheckBox.Location = new System.Drawing.Point(246, 61);
            this.regexCheckBox.Name = "regexCheckBox";
            this.regexCheckBox.Size = new System.Drawing.Size(79, 23);
            this.regexCheckBox.TabIndex = 4;
            this.regexCheckBox.Text = "Use Regex";
            this.regexCheckBox.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cancelButton.Location = new System.Drawing.Point(3, 90);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(92, 26);
            this.cancelButton.TabIndex = 7;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // okButton
            // 
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.Dock = System.Windows.Forms.DockStyle.Fill;
            this.okButton.Location = new System.Drawing.Point(101, 90);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(224, 26);
            this.okButton.TabIndex = 6;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // previewPanel
            // 
            this.previewPanel.Controls.Add(this.previewDataGridView);
            this.previewPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.previewPanel.Location = new System.Drawing.Point(3, 128);
            this.previewPanel.Name = "previewPanel";
            this.previewPanel.Size = new System.Drawing.Size(328, 194);
            this.previewPanel.TabIndex = 1;
            this.previewPanel.Visible = false;
            // 
            // previewDataGridView
            // 
            this.previewDataGridView.AllowUserToAddRows = false;
            this.previewDataGridView.AllowUserToDeleteRows = false;
            this.previewDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.previewDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.previewDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.beforeColumn,
            this.afterColumn});
            this.previewDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.previewDataGridView.Location = new System.Drawing.Point(0, 0);
            this.previewDataGridView.Name = "previewDataGridView";
            this.previewDataGridView.ReadOnly = true;
            this.previewDataGridView.RowHeadersVisible = false;
            this.previewDataGridView.Size = new System.Drawing.Size(328, 194);
            this.previewDataGridView.TabIndex = 0;
            // 
            // beforeColumn
            // 
            this.beforeColumn.HeaderText = "Before";
            this.beforeColumn.Name = "beforeColumn";
            this.beforeColumn.ReadOnly = true;
            // 
            // afterColumn
            // 
            this.afterColumn.HeaderText = "After";
            this.afterColumn.Name = "afterColumn";
            this.afterColumn.ReadOnly = true;
            // 
            // RenameForm
            // 
            this.AcceptButton = this.okButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(334, 131);
            this.Controls.Add(this.mainTableLayoutPanel);
            this.MaximumSize = new System.Drawing.Size(1080, 600);
            this.MinimumSize = new System.Drawing.Size(350, 170);
            this.Name = "RenameForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Rename Worksheets";
            this.TopMost = true;
            this.mainTableLayoutPanel.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.previewPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.previewDataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        #region Private Fields

        private TableLayoutPanel mainTableLayoutPanel;
        private TableLayoutPanel tableLayoutPanel1;
        private Label findLabel;
        private TextBox findTextBox;
        private Label replaceLabel;
        private TextBox replaceTextBox;
        private CheckBox regexCheckBox;
        private Button previewButton;
        private Button okButton;
        private Button cancelButton;
        private Panel previewPanel;
        private DataGridView previewDataGridView;
        private DataGridViewTextBoxColumn beforeColumn;
        private DataGridViewTextBoxColumn afterColumn;
        #endregion
    }
}