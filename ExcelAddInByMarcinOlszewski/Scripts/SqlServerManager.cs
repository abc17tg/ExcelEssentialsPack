using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Windows.Forms;
using ExcelAddInByMarcinOlszewski.Scripts;
using Oracle.ManagedDataAccess.Client;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInByMarcinOlszewski
{
    public class SqlServerManager
    {
        public static readonly string LookupString = "#TkL@.qKs1Hm8hJ-[nxB";

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

        public static void GetDataFromServerToNewSheet(string query, SqlConn sqlConn, bool headers = true, string wsName = "")
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = wb.Sheets.Add();
            if (!string.IsNullOrEmpty(wsName))
                ws.Rename(wsName);

            SqlResult sqlResult = GetDataFromServer(query, sqlConn);
            if (sqlResult.HasErrors || sqlResult.DataTable.Rows.Count < 1)
            {
                MessageBox.Show($"No data extracted\n{(sqlResult.Errors == null ? string.Empty : sqlResult.Errors)}", "Query finished", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                return;
            }

            if (!ws.Exists())
                return;

            if (sqlResult.DataTable.Rows.Count >= ws.Rows.Count)
                Utils.SaveDataTableToTxt(sqlResult.DataTable, string.Empty, ws.Name, (!string.IsNullOrEmpty(wb.Path) ? wb.Path : string.Empty));
            else
                UtilsExcel.PasteDataTableToRange(sqlResult.DataTable, ws.Cells[1, 1], headers);
        }

        public static void GetDataFromServerToSelection(string query, SqlConn sqlConn, Excel.Range rng, bool headers = true)
        {
            SqlResult sqlResult = GetDataFromServer(query, sqlConn, 180);
            if (sqlResult.HasErrors || sqlResult.DataTable.Rows.Count < 1)
            {
                MessageBox.Show($"No data extracted\n{(sqlResult.Errors == null ? string.Empty : sqlResult.Errors)}", "Query finished", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                return;
            }

            if (!rng.Valid())
                return;

            if (sqlResult.DataTable.Rows.Count >= rng.Worksheet.Rows.Count - rng.Row + 1)
                Utils.SaveDataTableToTxt(sqlResult.DataTable, string.Empty, rng.Worksheet.Name, (!string.IsNullOrEmpty(rng.Worksheet.Parent.Path) ? rng.Worksheet.Parent.Path : string.Empty));
            else
                UtilsExcel.PasteDataTableToRange(sqlResult.DataTable, rng, headers);
        }

        SaveFileDialog saveDlg = new SaveFileDialog();


        public static SqlResult GetDataFromServer(string query, SqlConn sqlConn, int timeout = -1)
        {
            DataTable dt = new DataTable();
            SqlResult sqlResult = null;
            switch (sqlConn.Type)
            {
                case ServerType.SqlServer:
                    sqlResult = GetDataFromSqlServer(query, sqlConn, timeout);
                    break;
                case ServerType.Oracle:
                    sqlResult = GetDataFromOracleSqlServer(query, sqlConn, timeout);
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

                return new SqlResult(dt,null);
            }
            catch (Exception ex)
            {
                return new SqlResult(null, ex.Message);
            }

        }

        public static SqlResult GetDataFromOracleSqlServer(string query, SqlConn sqlConn, int timeout = -1)
        {
            try
            {
                using (OracleConnection con = new OracleConnection(sqlConn.ConnectionString()))
                {
                    con.Open();
                    OracleCommand cmd = new OracleCommand(query, con);
                    cmd.CommandTimeout = timeout > 0 ? timeout : cmd.CommandTimeout;
                    OracleDataReader rdr = cmd.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(rdr);
                    return new SqlResult(dt, null);
                }
            }
            catch (OracleException ex)
            {
                return new SqlResult(null, ex.Message);
            }
        }

        public static SqlResult GetDataFromSqlServer(string query, SqlConn sqlConn, int timeout = -1)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(sqlConn.ConnectionString()))
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand(query, con);
                    cmd.CommandTimeout = timeout > 0 ? timeout : cmd.CommandTimeout;
                    SqlDataReader rdr = cmd.ExecuteReader();

                    DataTable dt = new DataTable();
                    dt.Load(rdr);
                    return new SqlResult(dt, null);
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
