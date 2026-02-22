using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ExcelEssentials.Forms
{
    public partial class FileDropForm : Form
    {
        public List<string> FilesPaths = new List<string>();
        private List<string> m_filter;
        private List<string> selectedFiles = new List<string>();

        public const Int32 WM_SYSCOMMAND = 0x112;
        public const Int32 MF_BYPOSITION = 0x400;
        public const Int32 ToggleTopMostMenuItem = 1000;
        public const Int32 CenterFormMenuItem = 1001;

        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll")]
        private static extern bool InsertMenu(IntPtr hMenu, Int32 wPosition, Int32 wFlags, Int32 wIDNewItem, string lpNewItem);

        public FileDropForm(List<string> filter = null)
        {
            InitializeComponent();
            m_filter = filter?.Select(p => p.Trim('.', ' ')).ToList();
            Activate();
            pathTextBox.Focus();
            Load += (o, s) => pathTextBox.Focus();
            pathTextBox.ReadOnly = true;
        }

        ~FileDropForm()
        {
            Load -= (o, s) => pathTextBox.Focus();
        }


        private void FileDropForm_Load(object sender, EventArgs e)
        {
            Utils.MoveFormToCursor(this);
            IntPtr MenuHandle = GetSystemMenu(this.Handle, false);
            InsertMenu(MenuHandle, 5, MF_BYPOSITION, ToggleTopMostMenuItem, "Pin/Unpin this window");
            InsertMenu(MenuHandle, 6, MF_BYPOSITION, CenterFormMenuItem, "Center window");
        }

        protected override void WndProc(ref Message msg)
        {
            if (msg.Msg == WM_SYSCOMMAND)
            {
                switch (msg.WParam.ToInt32())
                {
                    case ToggleTopMostMenuItem:
                        ToggleTopMost();
                        return;
                    case CenterFormMenuItem:
                        Utils.MoveFormToCenter(this);
                        return;
                    default:
                        break;
                }
            }
            base.WndProc(ref msg);
        }

        private void ToggleTopMost()
        {
            this.TopMost = !this.TopMost;
        }

        private void dropFileField_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void FileDropped(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            selectedFiles = files.ToList();
            UpdateUI();
        }

        private void UpdateUI()
        {
            if (selectedFiles.Count == 0)
            {
                pathTextBox.Text = string.Empty;
                dropFileLabel.Visible = true;
                droppedFileIcon.Visible = false;
                okBtn.Enabled = false;
                pathTextBox.BackColor = SystemColors.Window;
            }
            else
            {
                pathTextBox.Text = string.Join(";", selectedFiles);
                dropFileLabel.Visible = false;
                droppedFileIcon.Visible = true;

                int count = selectedFiles.Count;
                int cols = (int)Math.Ceiling(Math.Sqrt(count));
                int rows = (int)Math.Ceiling((double)count / cols);

                Bitmap gridImage = new Bitmap(droppedFileIcon.Width, droppedFileIcon.Height);
                using (Graphics g = Graphics.FromImage(gridImage))
                {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None;

                    int cellWidth = droppedFileIcon.Width / cols;
                    int cellHeight = droppedFileIcon.Height / rows;

                    int index = 0;
                    for (int r = 0; r < rows; r++)
                    {
                        for (int c = 0; c < cols; c++)
                        {
                            if (index < count)
                            {
                                string filePath = selectedFiles[index];
                                if (File.Exists(filePath))
                                {
                                    using (Icon icon = Icon.ExtractAssociatedIcon(filePath))
                                    using (Image image = icon.ToBitmap())
                                    {
                                        int S = Math.Min(cellWidth, cellHeight);
                                        int iconX = c * cellWidth + (cellWidth - S) / 2;
                                        int iconY = r * cellHeight + (cellHeight - S) / 2;
                                        g.DrawImage(image, iconX, iconY, S, S);
                                    }
                                }
                                index++;
                            }
                        }
                    }
                }
                droppedFileIcon.Image = gridImage;

                bool allValid = selectedFiles.All(f => File.Exists(f) &&
                    (m_filter == null || m_filter.Contains(Path.GetExtension(f).TrimStart('.').ToLower())));
                okBtn.Enabled = allValid;
                pathTextBox.BackColor = allValid ? Color.PaleGreen : Color.LightPink;
            }
        }

        private void okBtn_Click(object sender, EventArgs e)
        {
            FilesPaths = selectedFiles;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void dropFileLabel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = m_filter != null ? "Files|" + string.Join(";", m_filter.Select(p => "*." + p)) : "Files|*.*";
            var result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                selectedFiles = openFileDialog.FileNames.ToList();
                UpdateUI();
            }
        }

    }
}
