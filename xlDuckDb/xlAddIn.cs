using ExcelDna.Integration;
using ExcelDna.Logging;
using System.Diagnostics;

namespace xlDuckDb
{
    public class xlAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
        }

        public void AutoClose()
        {
        }

        [ExcelFunction(Description = "Executes a DuckDB query.", Category = "xlDuckDb",
            HelpTopic = "https://duckdb.org/docs/index")]
        public static object DuckDbQuery(
            [ExcelArgument(Name = "SQL", Description = "The SQL query to execute.")]
            string query,
            [ExcelArgument(Name = "Database File",
                Description = "Optionally specify the DuckDB database to use. Defaults to an in-memory database")]
            string dataSource = null)
        {
            if (ExcelDnaUtil.IsInFunctionWizard()) return null;

            var result = DuckDbHelper.ExecuteQuery(query, dataSource);
            return result.Length == 1 ? new object[,] {{ExcelError.ExcelErrorNA}} : result;
        }
    }
}