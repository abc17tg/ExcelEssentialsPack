using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelVB = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddInByMarcinOlszewski.Forms
{
    public partial class SapS4ExtractionForm : Form
    {
        private Excel.Workbook m_macroWorkbook;
        private List<ListViewItem> m_itemList = new List<ListViewItem>();

        private bool m_mouseDown;
        private Point m_lastLocation;

        public SapS4ExtractionForm(Excel.Workbook macroWb)
        {
            InitializeComponent();

            m_macroWorkbook = macroWb;
            if (m_macroWorkbook != null)
                RefreshList();
            else
                this.Close();
            this.MouseDown += new MouseEventHandler(SapS4ExtractionForm_MouseDown);
            this.MouseMove += new MouseEventHandler(SapS4ExtractionForm_MouseMove);
            this.MouseUp += new MouseEventHandler(SapS4ExtractionForm_MouseUp);
            foreach (var control in this.Controls.Cast<Control>().Where(p => p != searchTextBox))
                control.MouseClick += SapS4ExtractionForm_MouseClick;
        }

        private void SapS4ExtractionForm_MouseDown(object sender, MouseEventArgs e)
        {
            m_mouseDown = true;
            m_lastLocation = e.Location;
        }

        private void SapS4ExtractionForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (m_mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - m_lastLocation.X) + e.X, (this.Location.Y - m_lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void SapS4ExtractionForm_MouseUp(object sender, MouseEventArgs e)
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

        private void SapS4ExtractionForm_Load(object sender, EventArgs e)
        {
            Utils.MoveFormToCursor(this);
        }

        private void SapS4ExtractionForm_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                this.Close();
            }
        }

        private void refreshBtn_Click(object sender, EventArgs e)
        {
            RefreshList();
        }

        private void RefreshList()
        {
            List<string> macros = new List<string>();

            sapS4TemplatesListView.Items.Clear();
            m_itemList.Clear();
            foreach (ExcelVB.VBComponent component in m_macroWorkbook.VBProject.VBComponents)
            {
                macros = new List<string>();
                for (int i = 1; i < component.CodeModule.CountOfLines; i++)
                {
                    string str = component.CodeModule.ProcOfLine[i, out ExcelVB.vbext_ProcKind procKind];
                    if (!string.IsNullOrWhiteSpace(str) && str.StartsWith("LoginToS4AndGet") && str.Length > "LoginToS4AndGet".Length)
                        macros.Add(str);
                }

                if (macros.Count < 1)
                    continue;

                foreach (var macro in macros.Distinct().OrderBy(p => p))
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Name = macro + "Btn";
                    lvi.Text = macro.Replace("LoginToS4AndGet", "");
                    lvi.ToolTipText = $"{component.Name}.{macro}";
                    lvi.ForeColor = Color.White;
                    m_itemList.Add(lvi);
                }
            }

            if (m_itemList.Count > 0)
            {
                sapS4TemplatesListView.Items.AddRange(m_itemList.ToArray());
            }
        }

        private void searchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (searchTextBox.Text.Length == 0 || searchTextBox.Text == "Search")
            {
                sapS4TemplatesListView.Items.Clear();
                sapS4TemplatesListView.Items.AddRange(m_itemList.ToArray());
                return;
            }

            if (searchTextBox.Text.Length > 0)
            {
                sapS4TemplatesListView.Items.Clear();
                ListViewItem[] items = m_itemList.Where(p => p.Text.ToLower().Contains(searchTextBox.Text.ToLower())).ToArray();
                if (items != null && items.Length > 0)
                    sapS4TemplatesListView.Items.AddRange(items);
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

        private void sapS4TemplatesListView_DoubleClick(object sender, EventArgs e)
        {
            if (sapS4TemplatesListView.SelectedItems.Count > 0)
            {
                UtilsExcel.RunMacro(sapS4TemplatesListView.SelectedItems[sapS4TemplatesListView.SelectedItems.Count - 1].ToolTipText);
                this.Close();
            }
        }



        // Save the selected index and top item index when the ListView loses focus
        private void sapS4TemplatesListView_MouseLeave(object sender, EventArgs e)
        {
            if (!searchTextBox.Focused)
                searchTextBox.Focus();
        }

        // Restore the selected index and top item index when the ListView receives focus
        private void sapS4TemplatesListView_MouseEnter(object sender, EventArgs e)
        {
            if (!sapS4TemplatesListView.Focused)
                sapS4TemplatesListView.Focus();
        }

        private void sapS4TemplatesListView_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                if (sapS4TemplatesListView.SelectedItems.Count > 0)
                {
                    UtilsExcel.RunMacro(sapS4TemplatesListView.SelectedItems[sapS4TemplatesListView.SelectedItems.Count - 1].ToolTipText);
                    this.Close();
                }
            }
        }
    }
}

