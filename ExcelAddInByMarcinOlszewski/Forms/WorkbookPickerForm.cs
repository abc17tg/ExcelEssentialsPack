using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInByMarcinOlszewski.Forms
{
    public partial class WorkbookPickerForm : Form
    {
        public Excel.Workbook Workbook;
        private List<Excel.Workbook> m_macroWorkbookList;

        public WorkbookPickerForm(Excel.Workbook wb)
        {
            InitializeComponent();
            if(wb == null)
                this.Close();

            Workbook = wb;

            m_macroWorkbookList = Workbook.Application.Workbooks.Cast<Excel.Workbook>().Where(p => p.HasVBProject || p.IsAddin).ToList();
            var addInsList = Workbook.Application.AddIns2.Cast<Excel.AddIn>().Where(p=>p.IsOpen).Select(a => Workbook.Application.Workbooks[a.Name]).ToList();
            m_macroWorkbookList.AddRange(addInsList);
            workbookPickerComboBox.Items.AddRange(m_macroWorkbookList.Select(p => p.Name).ToArray());
            workbookPickerComboBox.SelectedIndex = workbookPickerComboBox.Items.IndexOf(Workbook.Name);
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            this.Close();
        }

        private void workbookPickerComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Workbook = m_macroWorkbookList.Where(p => p.Name == workbookPickerComboBox.SelectedItem?.ToString()).FirstOrDefault();
        }
    }
}
