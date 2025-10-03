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
        [ExcelArgument(Name = "Excel Ranges",
            Description = "Excel ranges can be passed and then referenced in the SQL as SELECT * FROM xlRange, if multiple ranges are passed use xlRange[1], xlRange[2], etc",
            AllowReference = true)]
        params object[] ranges)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var rangeAddresses = new string[ranges.Length];
        for (var i = 0; i < ranges.Length; i++)
        {
            if (ranges[i] is ExcelReference excelReference)
            {
                rangeAddresses[i] = XlCall.Excel(XlCall.xlfReftext, excelReference, true).ToString() ?? throw new InvalidOperationException("Failed to determine the address of the supplied range.");
            }
        }

        var result = DuckDbHelper.ExecuteQuery(query, dataSource, rangeAddresses);
        return result.Length == 1 ? new object[,] {{ExcelError.ExcelErrorNA}} : result;
    }

}