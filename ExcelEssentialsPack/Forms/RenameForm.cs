using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace ExcelEssentials.Forms
{
    public partial class RenameForm : Form
    {
        public bool isExpanded { get; private set; }
        public string FindText { get; private set; }
        public string ReplaceText { get; private set; }
        public bool UseRegex { get; private set; }
        public Dictionary<string, string> RenamedValues { get; private set; }

        private List<string> originalStrings;
        private readonly int collapsedHeight;
        private int previewPanelHeight; // Made non-readonly to set after init

        public RenameForm(List<string> items)
        {
            InitializeComponent();
            originalStrings = items ?? new List<string>();
            isExpanded = true; // Start as "expanded" internally to immediately collapse

            // Store heights after InitializeComponent() runs
            collapsedHeight = this.Height;
            previewPanelHeight = (int)mainTableLayoutPanel.RowStyles[1].Height;

            // Collapse the form on startup by simulating a click
            previewButton_Click(null, null);
            RenamedValues = new Dictionary<string, string>();
        }

        private void previewButton_Click(object sender, EventArgs e)
        {
            if (isExpanded)
            {
                // Collapse
                mainTableLayoutPanel.RowStyles[1].Height = 0;
                this.Height = collapsedHeight;
                previewPanel.Visible = false;
                previewButton.Text = "Preview ▼";
                isExpanded = false;
            }
            else
            {
                // Expand and populate preview
                if (!IsInputValid()) return; // Validate before expanding
                PopulatePreview();
                mainTableLayoutPanel.RowStyles[1].Height = previewPanelHeight;
                this.Height += previewPanelHeight;
                previewPanel.Visible = true;
                previewButton.Text = "Preview ▲";
                isExpanded = true;
            }
            this.PerformLayout(); // Force form re-layout after resize/visibility change
            this.Refresh();
        }

        private void PopulatePreview()
        {
            previewDataGridView.Rows.Clear();

            string find = findTextBox.Text;
            string replace = replaceTextBox.Text;
            bool useRegex = regexCheckBox.Checked;

            if (originalStrings == null || !originalStrings.Any())
            {
                previewDataGridView.Rows.Add("No items to rename", "");
                return;
            }

            foreach (string originalString in originalStrings)
            {
                string newString = ComputeNewString(originalString, find, replace, useRegex);
                previewDataGridView.Rows.Add(originalString, newString);
            }

            previewDataGridView.PerformLayout();
            previewDataGridView.Refresh();

            // Calculate preferred height: header + rows (assume default row height; adjust if custom)
            int headerHeight = previewDataGridView.ColumnHeadersHeight;
            int rowsHeight = previewDataGridView.Rows.Count * (previewDataGridView.RowTemplate.Height + 1); // +1 for grid line
            int preferredHeight = headerHeight + rowsHeight + 2; // Small buffer for borders/lines
            const int maxPreviewHeight = 300; // Cap to prevent huge forms; use scrollbar beyond this
            previewPanelHeight = Math.Min(preferredHeight, maxPreviewHeight);
            mainTableLayoutPanel.RowStyles[1].Height = previewPanelHeight;
        }

        private string ComputeNewString(string oldString, string find, string replace, bool useRegex)
        {
            // If 'find' is empty, replace the entire string.
            if (string.IsNullOrEmpty(find))
                return replace;

            if (useRegex)
            {
                try
                {
                    // Escapes '$' in the replacement string to prevent accidental capture group substitution
                    return Regex.Replace(oldString, find, replace.Replace("$", "$$"));
                }
                catch (ArgumentException)
                {
                    // This should not be reached due to IsValidRegex check, but is a good safeguard.
                    return oldString;
                }
            }
            else
            {
                return oldString.Replace(find, replace);
            }
        }

        private bool IsInputValid()
        {
            string find = findTextBox.Text;
            bool useRegex = regexCheckBox.Checked;

            // This validation was removed to allow the new "replace all" feature.

            if (useRegex && !string.IsNullOrEmpty(find) && !IsValidRegex(find))
            {
                MessageBox.Show("Invalid regex pattern. Please check your 'Find' input.", "Invalid Regex", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private bool IsValidRegex(string pattern)
        {
            if (string.IsNullOrEmpty(pattern)) return false; // An empty pattern isn't useful to validate here
            try
            {
                new Regex(pattern);
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            if (!IsInputValid())
            {
                return;
            }

            FindText = findTextBox.Text;
            ReplaceText = replaceTextBox.Text;
            UseRegex = regexCheckBox.Checked;

            if (originalStrings == null || !originalStrings.Any())
            {
                MessageBox.Show("No items provided to rename.", "No Items", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                RenamedValues = originalStrings
                    .GroupBy(value => value) // Handles duplicate source strings
                    .ToDictionary(
                        group => group.Key,
                        group => ComputeNewString(group.Key, FindText, ReplaceText, UseRegex)
                    );

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while processing the replacements: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}