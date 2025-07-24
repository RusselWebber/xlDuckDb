using ExcelDna.Integration;
using System.Diagnostics.CodeAnalysis;

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

    [Experimental("DuckDBNET001")]
    [ExcelFunction(Description = "Executes a DuckDB query.", Category = "xlDuckDb",
        HelpTopic = "https://duckdb.org/docs/index", IsMacroType = true)]
    public static object DuckDbQuery(
        [ExcelArgument(Name = "SQL", Description = "The SQL query to execute.")]
        string query,
        [ExcelArgument(Name = "Database File",
            Description = "Optionally specify the DuckDB database to use. Defaults to an in-memory database")]
        string dataSource,
        [ExcelArgument(Name = "Excel Range",
            Description = "An Excel range can be passed and then referenced in the SQL as SELECT * FROM xlRange",
            AllowReference = true)]
        object r)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;
        var rangeAddress = "";

        if (r is ExcelReference excelReference)
        {
            rangeAddress = XlCall.Excel(XlCall.xlfReftext, excelReference, true).ToString() ?? throw new InvalidOperationException("Failed to determine the address of the supplied range.");
        }

        var result = DuckDbHelper.ExecuteQuery(query, dataSource, rangeAddress);
        return result.Length == 1 ? new object[,] {{ExcelError.ExcelErrorNA}} : result;
    }

}