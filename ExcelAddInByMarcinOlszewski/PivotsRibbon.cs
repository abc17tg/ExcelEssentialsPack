using ExcelAddInByMarcinOlszewski.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddInByMarcinOlszewski
{
    public partial class MiscRibbon
    {
        private void generatePivotTemlateCodeButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro("PivotTablesTemplates.GenerateCreatePivotTableCode");
        }

        private void formatClickedPivotButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro("PivotTablesFormat.FormatSelectedPivotTableDesign");
        }

        private void formatAllPivotButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro("PivotTablesFormat.FormatAllPivotTableDesign");
        }

        private void refreshPivots_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro("PivotTablesFormat.RefreshPivotTables");
        }

        private void runPvTemplateBtn_Click(object sender, RibbonControlEventArgs e)
        {
            PvTemplatesListForm pvTemplatesListForm = new PvTemplatesListForm(m_macroWorkbook);
            pvTemplatesListForm.ShowDialog();
        }
    }
}
