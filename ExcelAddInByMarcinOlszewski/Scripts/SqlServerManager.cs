using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Windows.Forms;
using ColorMine.ColorSpaces;
using ExcelAddInByMarcinOlszewski.Forms;
using ExcelAddInByMarcinOlszewski.Scripts;
using Oracle.ManagedDataAccess.Client;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInByMarcinOlszewski
{
    public class SqlServerManager
    {
        public List<SqlCommand> RunningQueriesCmdSqlServer = new List<SqlCommand>();
        public List<OracleCommand> RunningQueriesCmdOracle = new List<OracleCommand>();

        // Define events for command completion
        public event Action<SqlCommand> SqlServerCommandFinished;
        public event Action<OracleCommand> OracleCommandFinished;

        public static readonly string LookupString = "#TkL@.qKs1Hm8hJ-[nxB";

        public SqlServerManager()
        {
            // Subscribe to events
            SqlServerCommandFinished += OnSqlServerCommandFinished;
            OracleCommandFinished += OnOracleCommandFinished;
        }

        // Method to handle SQL Server command completion
        private void OnSqlServerCommandFinished(SqlCommand command)
        {
            try
            {
                RunningQueriesCmdSqlServer.RemoveAll(cmd => cmd == null);
                RunningQueriesCmdSqlServer.Remove(command);
            }
            catch (Exception) { }
        }

        // Method to handle Oracle command completion
        private void OnOracleCommandFinished(OracleCommand command)
        {
            try
            {
                RunningQueriesCmdOracle.RemoveAll(cmd => cmd == null);
                RunningQueriesCmdOracle.Remove(command);
            }
            catch (Exception) { }
        }

        public static bool AddSqlConnection()
        {
            Form.ActiveForm.TopMost = false;
            ServerConnectionForm serverConnectionForm = new ServerConnectionForm();
            //serverConnectionForm.TopMost = true;
            var result = serverConnectionForm.ShowDialog();
            if (result == DialogResult.OK)
            {
                Form.ActiveForm.TopMost = true;
                return true;
            }
            else
            {
                Form.ActiveForm.TopMost = true;
                return false;
            }
        }

        public static bool TestConnectionOracle(string connectionString)
        {
            bool result;
            try
            {
                using (OracleConnection connection = new OracleConnection(connectionString))
                {
                    connection.Open();
                    connection.Close();
                    result = true;
                }
            }
            catch (OracleException ex)
            {
                result = false;
                MessageBox.Show(ex.Message.ToString());
            }
            return result;
        }

        public static bool TestConnectionSqlServer(string connectionString)
        {
            bool result;
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    connection.Close();
                    result = true;
                }
            }
            catch (SqlException ex)
            {
                result = false;
                MessageBox.Show(ex.Message.ToString());
            }
            return result;
        }

        public static bool GetDataFromServerToNewSheet(Control control, SqlServerManager manager, string query, SqlConn sqlConn, bool headers = true, string wsName = "")
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = wb.Sheets.Add();
            if (!string.IsNullOrEmpty(wsName))
                ws.Rename(wsName);

            SqlResult sqlResult = GetDataFromServer(manager, query, sqlConn);
            if (sqlResult.HasErrors || sqlResult.DataTable.Rows.Count < 1)
            {
                MessageBox.Show($"No data extracted\n{(sqlResult.Errors == null ? string.Empty : sqlResult.Errors)}", "Query finished", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }

            if (!ws.Exists())
                return false;

            if (sqlResult.DataTable.Rows.Count >= ws.Rows.Count - 1)
            {
                if (control == null || control.IsDisposed)
                    return false;

                control.Invoke(new Action(() =>
                {
                    MessageBoxForm messageBoxForm = new MessageBoxForm($"Query finished and too big to be pasted. Display as DataTable or discard? To discard close the message\n\n\n{query}", "Query finished and too big", true);
                    messageBoxForm.ShowDialog();
                    if (messageBoxForm.DialogResult == DialogResult.OK)
                    {
                        if (control == null || control.IsDisposed)
                            return;

                        DataTableForm dataTableForm = new DataTableForm(sqlResult.DataTable, query, ws.Application);
                        dataTableForm.Show();
                        dataTableForm.Activate();
                    }
                }));
                return false;
            }
            else
            {
                UtilsExcel.PasteDataTableToRange(sqlResult.DataTable, ws.Cells[1, 1], headers);
                return true;
            }
        }

        public static bool GetDataFromServerToSelection(Control control, SqlServerManager manager, string query, SqlConn sqlConn, Excel.Range rng, bool headers = true)
        {
            SqlResult sqlResult = GetDataFromServer(manager, query, sqlConn, 180);
            if (sqlResult.HasErrors || sqlResult.DataTable.Rows.Count < 1)
            {
                MessageBox.Show($"No data extracted\n{(sqlResult.Errors == null ? string.Empty : sqlResult.Errors)}", "Query finished", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                return false;
            }

            if (!rng.Valid())
                return false;

            if (sqlResult.DataTable.Rows.Count >= rng.Worksheet.Rows.Count - rng.Row + 1)
            {
                if (control == null || control.IsDisposed)
                    return false;

                control.Invoke(new Action(() =>
                {
                    MessageBoxForm messageBoxForm = new MessageBoxForm($"Query finished and too big to be pasted. Display as DataTable or discard? To discard close the message\n\n\n{query}", "Query finished and too big", true);
                    messageBoxForm.ShowDialog();
                    if (messageBoxForm.DialogResult == DialogResult.OK)
                    {
                        if (control == null || control.IsDisposed)
                            return;

                        DataTableForm form = new DataTableForm(sqlResult.DataTable, query, rng.Application);
                        form.Show();
                        form.Activate();
                    }
                }));
                return false;
            }
            else
            {
                UtilsExcel.PasteDataTableToRange(sqlResult.DataTable, rng, headers);
                return true;
            }
        }

        public static SqlResult GetDataFromServer(SqlServerManager manager, string query, SqlConn sqlConn, int timeout = -1)
        {
            DataTable dt = new DataTable();
            SqlResult sqlResult = null;
            switch (sqlConn.Type)
            {
                case ServerType.SqlServer:
                    sqlResult = GetDataFromSqlServer(manager, query, sqlConn, timeout);
                    break;
                case ServerType.Oracle:
                    sqlResult = GetDataFromOracleSqlServer(manager, query, sqlConn, timeout);
                    break;
                case ServerType.Excel:
                    sqlResult = GetDataFromExcelSqlTables(query);
                    break;
            }
            return sqlResult;
        }

        public static SqlResult GetDataFromExcelSqlTables(string query)
        {
            try
            {
                object rs = UtilsExcel.RunMacro("SqlQueries.ExecuteSQLQuery", new object[] { query });

                OleDbDataAdapter adapter = new OleDbDataAdapter();
                DataTable dt = new DataTable();
                adapter.Fill(dt, rs);

                return new SqlResult(dt, null);
            }
            catch (Exception ex)
            {
                return new SqlResult(null, ex.Message);
            }

        }

        public static SqlResult GetDataFromOracleSqlServer(SqlServerManager manager, string query, SqlConn sqlConn, int timeout = -1)
        {
            try
            {
                using (OracleConnection con = new OracleConnection(sqlConn.ConnectionString()))
                {
                    con.Open();
                    OracleCommand cmd = new OracleCommand(query, con);
                    cmd.CommandTimeout = timeout > 0 ? timeout : cmd.CommandTimeout;
                    manager.RunningQueriesCmdOracle.Add(cmd);
                    using (OracleDataReader rdr = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        rdr.SuppressGetDecimalInvalidCastException = true;
                        dt.Load(rdr);
                        manager.OracleCommandFinished?.Invoke(cmd);
                        return new SqlResult(dt, null);
                    }
                }
            }
            catch (OracleException ex)
            {
                manager.OracleCommandFinished?.Invoke(null);
                return new SqlResult(null, ex.Message);
            }
        }

        public static SqlResult GetDataFromSqlServer(SqlServerManager manager, string query, SqlConn sqlConn, int timeout = -1)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(sqlConn.ConnectionString()))
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query, con);
                    cmd.CommandTimeout = timeout > 0 ? timeout : cmd.CommandTimeout;
                    manager.RunningQueriesCmdSqlServer.Add(cmd);
                    //cmd.StatementCompleted += (s, a) => SqlServerCommandFinished?.Invoke(s as SqlCommand);
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Load(rdr);
                        manager.SqlServerCommandFinished?.Invoke(cmd);
                        return new SqlResult(dt, null);
                    }
                }
            }
            catch (SqlException ex)
            {
                return new SqlResult(null, ex.Message);
            }
        }

        public enum ServerType
        {
            SqlServer = 0,
            Oracle = 1,
            Excel = 2
        }
    }
}
