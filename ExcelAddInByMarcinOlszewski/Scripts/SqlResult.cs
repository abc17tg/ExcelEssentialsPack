using System.Data;

namespace ExcelAddInByMarcinOlszewski.Scripts
{
    public class SqlResult
    {
        public DataTable DataTable;
        public string Errors;
        public bool HasErrors => !string.IsNullOrEmpty(Errors) || DataTable == null;
        public SqlResult(DataTable dataTable, string errors) 
        { 
            DataTable = dataTable;
            Errors = errors;
        }

    }
}
