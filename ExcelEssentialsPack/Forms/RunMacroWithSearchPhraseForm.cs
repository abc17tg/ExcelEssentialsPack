using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelEssentials.Scripts;
using ExcelVB = Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using System.IO;


namespace ExcelEssentials.Forms
{
    public partial class RunMacroWithSearchPhraseForm : Form
    {
        public bool WasRun = false;

        private Excel.Workbook m_macroWorkbook;
        private List<ListViewItem> m_itemList = new List<ListViewItem>();
        Dictionary<string, bool?> m_searchPhrasesDic = new Dictionary<string, bool?>();
        private string m_xmlElementsName;
        private Dictionary<char, Color> m_colorMapping = new Dictionary<char, Color>();

        private bool m_mouseDown;
        private Point m_lastLocation;

        public RunMacroWithSearchPhraseForm(Excel.Workbook macroWb, string xmlElementsName, bool colorMapping = false)
        {
            InitializeComponent();

            if (colorMapping)
            {
                string[] mappingFileLines = File.ReadAllLines(Path.Combine(FileManager.PropertiesFilesPath, "RunMacroWithSearchPhraseFormColorMapping.txt"));

                foreach (string line in mappingFileLines)
                {
                    string[] parts = line.Split(',');

                    char letter = parts[0][0];
                    int r = int.Parse(parts[1]);
                    int g = int.Parse(parts[2]);
                    int b = int.Parse(parts[3]);

                    m_colorMapping.Add(letter, Color.FromArgb(r, g, b));
                }
            }

            m_macroWorkbook = macroWb;
            m_xmlElementsName = xmlElementsName;

            if (m_macroWorkbook != null)
                RefreshList();
            else
                this.Close();

            m_searchPhrasesDic = GetSearchPhrases();
            this.MouseDown += new MouseEventHandler(RunMacroWithSearchPhraseForm_MouseDown);
            this.MouseMove += new MouseEventHandler(RunMacroWithSearchPhraseForm_MouseMove);
            this.MouseUp += new MouseEventHandler(RunMacroWithSearchPhraseForm_MouseUp);
            foreach (var control in this.Controls.Cast<Control>().Where(p => p != searchTextBox))
                control.MouseClick += RunMacroWithSearchPhraseForm_MouseClick;
        }

        public Dictionary<string, bool?> GetSearchPhrases()
        {
            try
            {
                var xe = XElement.Load(Path.Combine(FileManager.PropertiesFilesPath, "FormsMacrosSearchKeys.xml"));
                Dictionary<string, bool?> phrasesDic = xe.Elements(m_xmlElementsName)
                    .Select(p => new
                    {
                        s = (string)p.Element("SearchPhrase"),
                        kn = (bool?)p.Element("KeepSearchPhrase")
                    })
                    .ToDictionary(p => p.s, p => p.kn);

                if (phrasesDic != null && phrasesDic.Keys.Count > 0)
                    return phrasesDic;
                else
                    return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void RunMacroWithSearchPhraseForm_MouseDown(object sender, MouseEventArgs e)
        {
            m_mouseDown = true;
            m_lastLocation = e.Location;
        }

        private void RunMacroWithSearchPhraseForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (m_mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - m_lastLocation.X) + e.X, (this.Location.Y - m_lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void RunMacroWithSearchPhraseForm_MouseUp(object sender, MouseEventArgs e)
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

        private void RunMacroWithSearchPhraseForm_Load(object sender, EventArgs e)
        {
            Utils.MoveFormToCursor(this);
        }

        private void RunMacroWithSearchPhraseForm_MouseClick(object sender, MouseEventArgs e)
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
            Dictionary<string, KeyValuePair<string, bool?>> macrosD = new Dictionary<string, KeyValuePair<string, bool?>>();

            templatesListView.Items.Clear();
            m_itemList.Clear();
            m_searchPhrasesDic = GetSearchPhrases();
            foreach (ExcelVB.VBComponent component in m_macroWorkbook.VBProject.VBComponents)
            {
                macrosD = new Dictionary<string, KeyValuePair<string, bool?>>();
                for (int i = 1; i < component.CodeModule.CountOfLines; i++)
                {
                    string macroName = component.CodeModule.ProcOfLine[i, out ExcelVB.vbext_ProcKind procKind];

                    if (string.IsNullOrWhiteSpace(macroName))
                        continue;

                    if (macrosD.Keys.Contains(macroName))
                    {
                        i += component.CodeModule.ProcCountLines[macroName, procKind] - 1;
                        continue;
                    }

                    bool isNoArgument = component.CodeModule.Lines[component.CodeModule.ProcStartLine[macroName, procKind], component.CodeModule.ProcStartLine[macroName, procKind]].Contains($"Sub {macroName}()");

                    if (!isNoArgument)
                    {
                        i += component.CodeModule.ProcCountLines[macroName, procKind] - 1;
                        continue;
                    }

                    var keyValuePair = m_searchPhrasesDic.FirstOrDefault(p => (macroName.StartsWith(p.Key) && (macroName != p.Key || (p.Value ?? false))));
                    if (keyValuePair.Key != null && keyValuePair.Value != null)
                        macrosD.Add(macroName, keyValuePair);
                }

                if (macrosD.Keys.Count < 1)
                    continue;

                foreach (var macro in macrosD.OrderBy(p => p.Key))
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Name = macro.Key + "Btn";
                    lvi.Text = (!macro.Value.Value ?? true) ? macro.Key.Replace(macro.Value.Key, "") : macro.Key;
                    lvi.ToolTipText = $"{component.Name}.{macro.Key}";

                    char firstLetter = lvi.Text[0].ToString().ToUpper()[0];
                    if (m_colorMapping != null && m_colorMapping.ContainsKey(firstLetter))
                        lvi.ForeColor = m_colorMapping[firstLetter];
                    else
                        lvi.ForeColor = Color.White;

                    m_itemList.Add(lvi);
                }
            }

            if (m_itemList.Count > 0)
            {
                templatesListView.Items.AddRange(m_itemList.ToArray());
            }
        }

        private void searchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (searchTextBox.Text.Length == 0 || searchTextBox.Text == "Search")
            {
                templatesListView.Items.Clear();
                templatesListView.Items.AddRange(m_itemList.ToArray());
                return;
            }

            if (searchTextBox.Text.Length > 0)
            {
                templatesListView.Items.Clear();
                ListViewItem[] items = m_itemList.Where(p => p.Text.ToLower().Contains(searchTextBox.Text.ToLower())).ToArray();
                if (items != null && items.Length > 0)
                    templatesListView.Items.AddRange(items);
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

        private void templatesListView_DoubleClick(object sender, EventArgs e)
        {
            if (templatesListView.SelectedItems.Count > 0)
            {
                UtilsExcel.RunMacro(templatesListView.SelectedItems[templatesListView.SelectedItems.Count - 1].ToolTipText);
                WasRun = true;
                this.Close();
            }
        }



        // Save the selected index and top item index when the ListView loses focus
        private void templatesListView_MouseLeave(object sender, EventArgs e)
        {
            if (!searchTextBox.Focused)
                searchTextBox.Focus();
        }

        // Restore the selected index and top item index when the ListView receives focus
        private void templatesListView_MouseEnter(object sender, EventArgs e)
        {
            if (!templatesListView.Focused)
                templatesListView.Focus();
        }

        private void templatesListView_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                if (templatesListView.SelectedItems.Count > 0)
                {
                    UtilsExcel.RunMacro(templatesListView.SelectedItems[templatesListView.SelectedItems.Count - 1].ToolTipText);
                    WasRun = true;
                    this.Close();
                }
            }
        }
    }
}

