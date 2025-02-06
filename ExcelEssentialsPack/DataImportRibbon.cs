using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using ExcelEssentials.Forms;
using ExcelEssentials.Scripts;
using ScintillaNET;

namespace ExcelEssentials
{
    public partial class MiscRibbon
    {
        private void sqlEditorDataFolderBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Process.Start("explorer.exe", FileManager.SqlQueriesPath);
        }

        private void runS4ExtractButton_Click(object sender, RibbonControlEventArgs e)
        {
            //SapS4ExtractionForm form = new SapS4ExtractionForm(m_macroWorkbook);
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            RunMacroWithSearchPhraseForm form = new RunMacroWithSearchPhraseForm(m_macroWorkbook, "SapS4ExtractionSearch");
            form.Show();
        }
        
        private void runSdeButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(sdeQueryComboBox.Text))
                return;

            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection.GetUsableRange();

            if (!rng.Valid() || rng.Columns.Count > 1 || rng.Cells.Count > 50000)
                return;

            List<string> items = rng.Cells?.Cast<Excel.Range>()?.Select(p => (string)p.Value)?.Distinct()?.ToList();

            if (items == null || items.Count == 0)
                return;

            Sde sde = new Sde(sdeQueryComboBox.Text, int.Parse(sdeInstancesEditBox.Text), items);
            sde.Run();
        }

        private void sdeInstancesEditBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string numericText = new string(sdeInstancesEditBox.Text.Where(char.IsDigit).ToArray());

            int number = int.Parse(numericText);
            if (number > 50)
                number = 50;
            else if (number < 1)
                number = 1;

            numericText = number.ToString();

            if (numericText != sdeInstancesEditBox.Text)
                sdeInstancesEditBox.Text = numericText;
        }

        private void browserButton_Click(object sender, RibbonControlEventArgs e)
        {
            var bookmarks = BrowserViewForm.GetBookmarks();
            string bookmark;
            BrowserViewForm form;
            if (bookmarks != null && bookmarks.Count > 0 && bookmarks.TryGetValue(browserWebsitesComboBox.Text, out bookmark))
                form = new BrowserViewForm(importFromBrowserCheckBox.Checked, bookmark);
            else
                form = new BrowserViewForm(importFromBrowserCheckBox.Checked);
            form.Show();
        }

        private void loadToDataTableButton_Click(object sender, RibbonControlEventArgs e)
        {
            SqlExcelTableForm form = new SqlExcelTableForm(Globals.ThisAddIn.Application);
            form.Show();
        }
    }
}
