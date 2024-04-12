using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelVB = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAddInByMarcinOlszewski.Scripts;

namespace ExcelAddInByMarcinOlszewski.Forms
{
    public partial class RunMacroForm : Form
    {
        private Excel.Workbook m_macroWorkbook;
        private List<Excel.Workbook> m_macroWorkbookList;
        private List<ListViewItem> m_itemList = new List<ListViewItem>();
        private List<Macro> m_macrosList = new List<Macro>();
        private bool m_mouseDown;
        private Point m_lastLocation;

        public RunMacroForm(Excel.Workbook macroWb)
        {
            InitializeComponent();

            m_macroWorkbook = macroWb;
            if (m_macroWorkbook != null)
                RefreshList();
            else
                this.Close();

            m_macroWorkbookList = macroWb.Application.Workbooks.Cast<Excel.Workbook>().Where(p => p.HasVBProject).ToList();
            workbookPickComboBox.Items.AddRange(m_macroWorkbookList.Select(p => p.Name).ToArray());
            workbookPickComboBox.SelectedIndex = workbookPickComboBox.Items.IndexOf(macroWb.Name);
            workbookPickComboBox.SelectedIndexChanged += workbookPickComboBox_SelectedIndexChanged;

            this.MouseDown += new MouseEventHandler(RunMacroForm_MouseDown);
            this.MouseMove += new MouseEventHandler(RunMacroForm_MouseMove);
            this.MouseUp += new MouseEventHandler(RunMacroForm_MouseUp);
            foreach (var control in this.Controls.Cast<Control>().Where(p => p != searchTextBox || p != vbaEditorScintilla))
                control.MouseClick += RunMacroForm_MouseClick;

            UtilsScintilla.SetupVbaEditor(vbaEditorScintilla);
        }

        private void RunMacroForm_MouseDown(object sender, MouseEventArgs e)
        {
            m_mouseDown = true;
            m_lastLocation = e.Location;
        }

        private void RunMacroForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (m_mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - m_lastLocation.X) + e.X, (this.Location.Y - m_lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void RunMacroForm_MouseUp(object sender, MouseEventArgs e)
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

        private void RunMacroForm_Load(object sender, EventArgs e)
        {
            Utils.MoveFormToCursor(this);
        }

        private void refreshBtn_Click(object sender, EventArgs e)
        {
            RefreshList();
        }

        private void RefreshWbList()
        {

        }

        private void RefreshList()
        {
            macrosListView.Clear();
            m_itemList.Clear();
            m_macrosList.Clear();
            foreach (ExcelVB.VBComponent component in m_macroWorkbook.VBProject.VBComponents)
            {
                if (component.Type == ExcelVB.vbext_ComponentType.vbext_ct_StdModule)
                    for (int i = 1; i < component.CodeModule.CountOfLines; i++)
                    {
                        string macroName = component.CodeModule.ProcOfLine[i, out ExcelVB.vbext_ProcKind procKind];
                        if (!string.IsNullOrWhiteSpace(macroName) && procKind == ExcelVB.vbext_ProcKind.vbext_pk_Proc && !component.CodeModule.Lines[component.CodeModule.ProcStartLine[macroName, procKind], 1].Contains("Function"))
                        {
                            m_macrosList.Add(new Macro
                            {
                                Name = macroName,
                                ModuleName = component.Name,
                                Code = component.CodeModule.Lines[component.CodeModule.ProcStartLine[macroName, procKind], component.CodeModule.ProcCountLines[macroName, procKind]]
                            });
                            i += component.CodeModule.ProcCountLines[macroName, procKind] - 1;
                        }
                    }
            }

            m_macrosList = m_macrosList.Distinct().Where(p => p.FirstCodeLine.TrimEnd().EndsWith("()")).ToList();

            foreach (var macro in m_macrosList)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Name = macro.Name + "Btn";
                lvi.Text = macro.Name;
                lvi.ToolTipText = macro.FullName;
                lvi.ForeColor = Color.LightBlue;
                m_itemList.Add(lvi);
            }

            if (m_itemList.Count > 0)
            {
                macrosListView.Items.AddRange(m_itemList.ToArray());
            }

        }

        private void searchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (searchTextBox.Text.Length == 0 || searchTextBox.Text == "Search")
            {
                if (macrosListView.Items.Count == m_itemList.Count)
                    return;
                macrosListView.Clear();
                macrosListView.Items.AddRange(m_itemList.ToArray());
                return;
            }

            if (searchTextBox.Text.Length > 0)
            {
                macrosListView.Clear();
                ListViewItem[] items = m_itemList.Where(p => p.Text.ToLower().Contains(searchTextBox.Text.ToLower())).ToArray();
                if (items != null && items.Length > 0)
                    macrosListView.Items.AddRange(items);
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

        private void macrosListView_DoubleClick(object sender, EventArgs e)
        {
            if (macrosListView.SelectedItems.Count > 0)
            {
                UtilsExcel.RunMacro(macrosListView.SelectedItems[macrosListView.SelectedItems.Count - 1].ToolTipText, m_macroWorkbook.Name);
                this.Close();
            }
        }

        private void RunMacroForm_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                this.Close();
            }
        }

        // Save the selected index and top item index when the ListView loses focus
        private void macrosListView_MouseLeave(object sender, EventArgs e)
        {
            if (!searchTextBox.Focused)
                searchTextBox.Focus();
        }

        // Restore the selected index and top item index when the ListView receives focus
        private void macrosListView_MouseEnter(object sender, EventArgs e)
        {
            if (!macrosListView.Focused)
                macrosListView.Focus();
        }

        private void macrosListView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            vbaEditorScintilla.ReadOnly = false;
            vbaEditorScintilla.Text = m_macrosList.FirstOrDefault(p => p.FullName == e.Item.ToolTipText)?.Code ?? "";
            vbaEditorScintilla.ReadOnly = true;
        }

        private void closeBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void workbookPickComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_macroWorkbook = m_macroWorkbookList.Where(p => p.Name == workbookPickComboBox.SelectedItem?.ToString()).FirstOrDefault() ?? m_macroWorkbook;
            RefreshList();
        }

        private void macrosListView_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                if (macrosListView.SelectedItems.Count > 0)
                {
                    UtilsExcel.RunMacro(macrosListView.SelectedItems[macrosListView.SelectedItems.Count - 1].ToolTipText);
                    this.Close();
                }
            }
        }
    }
}
