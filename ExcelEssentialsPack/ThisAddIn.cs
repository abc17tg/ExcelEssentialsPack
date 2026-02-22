using System.IO;
using System.Windows.Threading;
using ExcelEssentials.Scripts;
using static ScintillaNET.Style;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEssentials
{
    public partial class ThisAddIn
    {
        private Dispatcher m_dispatcher = Dispatcher.CurrentDispatcher;
        public Dispatcher Dispatcher { get { return m_dispatcher; } }

        public int DefaultZoom = 100;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            DefaultZoom = LoadDefaultZoom();

            var app = Globals.ThisAddIn.Application;
            ((Excel.AppEvents_Event)app).NewWorkbook += App_NewWorkbook;
            app.WorkbookNewSheet += App_WorkbookNewSheet;

        }

        private void App_NewWorkbook(Excel.Workbook wb)
        {
            Application.ActiveWindow.Zoom = DefaultZoom;
        }

        public void SetDefaultZoom()
        {
            Application.ActiveWindow.Zoom = DefaultZoom;
        }

        private void App_WorkbookNewSheet(Excel.Workbook wb, object sheet)
        {
            Excel.Worksheet ws = sheet as Excel.Worksheet;

            foreach (Excel.Window win in wb.Windows)
            {
                if (win.ActiveSheet is Excel.Worksheet active && active == ws)
                {
                    win.Zoom = DefaultZoom;
                    return;
                }
            }
            Application.ActiveWindow.Zoom = DefaultZoom;
        }

        private int LoadDefaultZoom()
        {
            const int fallback = 100;
            try
            {
                var settingsFile = Path.Combine(FileManager.PropertiesFilesPath, "DefaultZoom.txt");
                if (!File.Exists(settingsFile))
                    return fallback;

                var text = File.ReadAllText(settingsFile).Trim();
                return int.TryParse(text, out var zoom) ? zoom : fallback;
            }
            catch
            {
                return fallback;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            ((Excel.AppEvents_Event)Application).NewWorkbook -= App_NewWorkbook;
            Application.WorkbookNewSheet -= App_WorkbookNewSheet;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
