using ExcelEssentials.Forms;
using ExcelEssentials.Scripts;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEssentials
{
    public partial class MiscRibbon
    {
        private void generatePivotTemlateCodeButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesTemplates.GenerateCreatePivotTableCode");
        }

        private void formatClickedPivotButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesFormat.FormatSelectedPivotTableDesign");
        }

        private void formatAllPivotButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesFormat.FormatAllPivotTableDesign");
        }

        private void refreshPivotsButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesFormat.RefreshPivotTables");
        }

        private void runPvTemplateButton_Click(object sender, RibbonControlEventArgs e)
        {
            //PvTemplatesListForm form = new PvTemplatesListForm(m_macroWorkbook);
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            RunMacroWithSearchPhraseForm form = new RunMacroWithSearchPhraseForm(m_macroWorkbook, "PivotTableSearch", true);
            form.Show();
        }

        private void combinedTableFromPvValuesButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesFormat.FormatSelectedPivotTableDesign");
        }

        private void updatePivotTableSourceButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
        }

        private void changePivotTableSourceButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
        }

        private void grandTotalsToggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;
            if (rng.IsPivotCell() && rng.PivotCell.PivotTable != null)
            {
                var pv = rng.PivotCell.PivotTable;
                bool toggle = !pv.ColumnGrand && !pv.RowGrand;
                pv.ColumnGrand = toggle;
                pv.RowGrand = toggle;
                grandTotalsToggleButton.Checked = toggle;
            }
        }

        private void subtotalsToggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;
            if (rng.IsPivotCell() && rng.PivotCell.PivotTable != null)
            {
                var pf = rng.PivotCell.PivotField;
                pf.Subtotals[1] = !pf.Subtotals[1];
                subtotalsToggleButton.Checked = pf.Subtotals[1];
            }
        }
    }
}
