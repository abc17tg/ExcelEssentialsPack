using ExcelAddInByMarcinOlszewski.Forms;
using ExcelAddInByMarcinOlszewski.Scripts;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddInByMarcinOlszewski
{
    public partial class MiscRibbon
    {
        private void generatePivotTemlateCodeButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesTemplates.GenerateCreatePivotTableCode");
        }

        private void formatClickedPivotButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesFormat.FormatSelectedPivotTableDesign");
        }

        private void formatAllPivotButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesFormat.FormatAllPivotTableDesign");
        }

        private void refreshPivotsButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesFormat.RefreshPivotTables");
        }

        private void runPvTemplateButton_Click(object sender, RibbonControlEventArgs e)
        {
            //PvTemplatesListForm form = new PvTemplatesListForm(m_macroWorkbook);
            RunMacroWithSearchPhraseForm form = new RunMacroWithSearchPhraseForm(m_macroWorkbook, "PivotTableSearch", true);
            form.Show();
        }

        private void combinedTableFromPvValuesButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
            //UtilsExcel.RunMacro("PivotTablesFormat.FormatSelectedPivotTableDesign");
        }

        private void updatePivotTableSourceButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Name, m_macroWorkbook));
        }
    }
}
