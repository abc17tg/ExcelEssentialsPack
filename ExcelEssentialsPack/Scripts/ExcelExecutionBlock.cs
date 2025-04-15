using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelEssentials.Scripts
{
    public class ExcelExecutionBlock : IDisposable
    {
        private Excel.Application m_app;
        private Excel.XlCalculation m_xlCalculation;
        private bool m_screenUpdating;
        private bool m_events;
        private bool m_interactive;
        private bool m_turnOnUpdatesOnDispose;
        public ExcelExecutionBlock(Excel.Application app, bool turnOnUpdatesOnDispose = false)
        {
            m_turnOnUpdatesOnDispose = turnOnUpdatesOnDispose;
            m_app = app;
            m_screenUpdating = app.ScreenUpdating;
            m_events = app.EnableEvents;
            m_xlCalculation = app.Calculation;
            m_interactive = app.Interactive;

            m_app.ScreenUpdating = false;
            m_app.EnableEvents = false;
            m_app.Calculation = Excel.XlCalculation.xlCalculationManual;
            try
            {
                if (m_app.Interactive)
                    m_app.Interactive = false;
            }
            catch (COMException) { }
        }

        public void Dispose()
        {
            m_app.ScreenUpdating = m_turnOnUpdatesOnDispose ? true : m_screenUpdating;
            m_app.EnableEvents = m_turnOnUpdatesOnDispose ? true : m_events;
            m_app.Calculation = m_turnOnUpdatesOnDispose ? Excel.XlCalculation.xlCalculationAutomatic : m_xlCalculation;
            try
            {
                if (!m_app.Interactive)
                    m_app.Interactive = true;
            }
            catch (COMException) { }
        }
    }
}
