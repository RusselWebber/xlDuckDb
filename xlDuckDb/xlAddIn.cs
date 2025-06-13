using ExcelDna.Integration;

namespace xlDuckDb;

public class xlAddIn : IExcelAddIn
{
    private const string ErrorPrefix = "#ERR";

    public void AutoOpen()
    {
        ExcelIntegration.RegisterUnhandledExceptionHandler(
            ex =>
            {
                if (ex is Exception e) return $"{ErrorPrefix} - {e.Message}";
                return $"{ErrorPrefix} - {ex}";
            });
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
        string dataSource = "")
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var result = DuckDbHelper.ExecuteQuery(query, dataSource);
        return result.Length == 1 ? new object[,] {{ExcelError.ExcelErrorNA}} : result;
    }
}