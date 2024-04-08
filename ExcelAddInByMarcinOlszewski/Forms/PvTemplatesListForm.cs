using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelVB = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelAddInByMarcinOlszewski.Forms
{
    public partial class PvTemplatesListForm : Form
    {
        private Excel.Workbook m_macroWorkbook;
        private List<ListViewItem> m_itemList = new List<ListViewItem>();

        private bool m_mouseDown;
        private Point m_lastLocation;
        private Dictionary<char, Color> m_colorMapping = new Dictionary<char, Color>();

        public PvTemplatesListForm(Excel.Workbook macroWb)
        {
            InitializeComponent();
            string[] mappingFileLines = File.ReadAllLines(Path.Combine(FileManager.PropertiesFilesPath, "PvTemplatesTextColorMapping.txt"));

            foreach (string line in mappingFileLines)
            {
                string[] parts = line.Split(',');

                char letter = parts[0][0];
                int r = int.Parse(parts[1]);
                int g = int.Parse(parts[2]);
                int b = int.Parse(parts[3]);

                m_colorMapping.Add(letter, Color.FromArgb(r, g, b));
            }

            m_macroWorkbook = macroWb;
            if (m_macroWorkbook != null)
                RefreshList();
            else
                this.Close();
            this.MouseDown += new MouseEventHandler(PvTemplatesListForm_MouseDown);
            this.MouseMove += new MouseEventHandler(PvTemplatesListForm_MouseMove);
            this.MouseUp += new MouseEventHandler(PvTemplatesListForm_MouseUp);
            foreach (var control in this.Controls.Cast<Control>().Where(p => p != searchTextBox))
                control.MouseClick += PvTemplatesListForm_MouseClick;
        }

        private void PvTemplatesListForm_MouseDown(object sender, MouseEventArgs e)
        {
            m_mouseDown = true;
            m_lastLocation = e.Location;
        }

        private void PvTemplatesListForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (m_mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - m_lastLocation.X) + e.X, (this.Location.Y - m_lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void PvTemplatesListForm_MouseUp(object sender, MouseEventArgs e)
        {
            m_mouseDown = false;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                this.Close();
                return true; // Indicate that you handled the key event
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }


        /* void PvTemplatesListForm_KeyDown(object sender, KeyEventArgs e)
         {
             if (e.KeyCode == Keys.Escape)
                 this.Close();
         }*/

        private void PvTemplatesListForm_Load(object sender, EventArgs e)
        {
            Utils.MoveFormToCursor(this);
        }

        private void refreshBtn_Click(object sender, EventArgs e)
        {
            RefreshList();
        }

        private void RefreshList()
        {
            List<string> macros = new List<string>();

            pvTemplatesListView.Items.Clear();
            m_itemList.Clear();
            foreach (ExcelVB.VBComponent component in m_macroWorkbook.VBProject.VBComponents)
            {
                macros = new List<string>();
                for (int i = 1; i < component.CodeModule.CountOfLines; i++)
                {
                    string str = component.CodeModule.ProcOfLine[i, out ExcelVB.vbext_ProcKind procKind];
                    if (!string.IsNullOrWhiteSpace(str) && str.StartsWith("CreatePvFrom") && str.Length > "CreatePvFrom".Length)
                        macros.Add(str);
                }

                if (macros.Count < 1)
                    continue;

                foreach (var macro in macros.Distinct().OrderBy(p => p))
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Name = macro + "Btn";
                    lvi.Text = macro.Replace("CreatePvFrom", "");
                    lvi.ToolTipText = $"{component.Name}.{macro}";
                    char firstLetter = lvi.Text[0].ToString().ToUpper()[0];

                    if (m_colorMapping.ContainsKey(firstLetter))
                        lvi.ForeColor = m_colorMapping[firstLetter];
                    else
                        lvi.ForeColor = Color.White;
                    m_itemList.Add(lvi);
                }
            }
            
            if (m_itemList.Count > 0)
            {
                pvTemplatesListView.Items.AddRange(m_itemList.ToArray());
            }
        }

        private void searchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (searchTextBox.Text.Length == 0 || searchTextBox.Text == "Search")
            {
                pvTemplatesListView.Items.Clear();
                pvTemplatesListView.Items.AddRange(m_itemList.ToArray());
                return;
            }

            if (searchTextBox.Text.Length > 0)
            {
                pvTemplatesListView.Items.Clear();
                ListViewItem[] items = m_itemList.Where(p => p.Text.ToLower().Contains(searchTextBox.Text.ToLower())).ToArray();
                if (items != null && items.Length > 0)
                    pvTemplatesListView.Items.AddRange(items);
                return;
            }
        }

        private void searchTextBox_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(searchTextBox.Text))
                searchTextBox.Text = "Search";
        }

        private void searchTextBox_Enter(object sender, EventArgs e)
        {
            if (searchTextBox.Text == "Search")
                searchTextBox.Text = string.Empty;
        }

        private void pvTemplatesListView_DoubleClick(object sender, EventArgs e)
        {
            if (pvTemplatesListView.SelectedItems.Count > 0)
            {
                UtilsExcel.RunMacro(pvTemplatesListView.SelectedItems[pvTemplatesListView.SelectedItems.Count - 1].ToolTipText);
                this.Close();
            }
        }

        private void PvTemplatesListForm_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                this.Close();
            }
        }


        // Save the selected index and top item index when the ListView loses focus
        private void pvTemplatesListView_MouseLeave(object sender, EventArgs e)
        {
            if (!searchTextBox.Focused)
                searchTextBox.Focus();
        }

        // Restore the selected index and top item index when the ListView receives focus
        private void pvTemplatesListView_MouseEnter(object sender, EventArgs e)
        {
            if (!pvTemplatesListView.Focused)
                pvTemplatesListView.Focus();
        }

        private void pvTemplatesListView_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                if (pvTemplatesListView.SelectedItems.Count > 0)
                {
                    UtilsExcel.RunMacro(pvTemplatesListView.SelectedItems[pvTemplatesListView.SelectedItems.Count - 1].ToolTipText);
                    this.Close();
                }
            }
        }
    }
}
