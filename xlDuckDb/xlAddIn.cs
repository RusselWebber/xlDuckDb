using ExcelDna.Integration;
using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;

namespace xlDuckDb;

public class xlAddIn : IExcelAddIn
{
    private const string ErrorPrefix = "#ERR";
    private static readonly ConcurrentDictionary<string, object[,]> _rangeCache =
        new ConcurrentDictionary<string, object[,]>();

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
        _rangeCache.Clear();
    }

    /// <summary>
    /// Clears all cached range data.
    /// </summary>
    public static void ClearCache()
    {
        _rangeCache.Clear();
    }

    /// <summary>
    /// Gets cached range data by cache key (range address).
    /// </summary>
    /// <param name="cacheKey">The cache key (range address or array key).</param>
    /// <returns>The cached 2D array, or null if not found.</returns>
    internal static object[,]? GetCachedRangeData(string cacheKey)
    {
        return _rangeCache.TryGetValue(cacheKey, out var data) ? data : null;
    }

    /// <summary>
    /// Removes a cached range item by cache key (range address or array key).
    /// </summary>
    /// <param name="cacheKey">The cache key to remove from cache.</param>
    /// <returns>True if the item was removed, false if it was not found.</returns>
    public static bool RemoveCachedRangeData(string cacheKey)
    {
        return _rangeCache.TryRemove(cacheKey, out _);
    }

    /// <summary>
    /// Caches range data with the given cache key (range address or array key).
    /// </summary>
    /// <param name="cacheKey">The cache key to use.</param>
    /// <param name="data">The 2D array to cache.</param>
    internal static void CacheRangeData(string cacheKey, object[,] data)
    {
        _rangeCache.AddOrUpdate(cacheKey, data, (_, _) => data);
    }

    [Experimental("DuckDBNET001")]
    [ExcelFunction(Description = "Executes a DuckDB query.", Category = "xlDuckDb",
        HelpTopic = "https://duckdb.org/docs/index", IsMacroType = true)]
    /// <summary>
    /// Executes a DuckDB query, supporting both Excel ranges and array arguments. Handles caching and cache cleanup.
    /// </summary>
    /// <param name="query">The SQL query to execute.</param>
    /// <param name="dataSource">The DuckDB database file or ":memory:" for in-memory.</param>
    /// <param name="ranges">Excel ranges (ExcelReference) or arrays (object[,]) to use as tables in the query.</param>
    /// <returns>Query result as a 2D object array.</returns>
    public static object DuckDbQuery(
        [ExcelArgument(Name = "SQL", Description = "The SQL query to execute.")]
        string query,
        [ExcelArgument(Name = "Database File",
            Description = "Optionally specify the DuckDB database to use. Defaults to an in-memory database")]
        string dataSource,
        [ExcelArgument(Name = "Excel Ranges",
            Description = "Excel ranges or arrays can be passed and then referenced in the SQL as SELECT * FROM xlRange, if multiple ranges are passed use xlRange[1], xlRange[2], etc",
            AllowReference = true)]
        params object[] ranges)
    {
        if (ExcelDnaUtil.IsInFunctionWizard()) return ExcelError.ExcelErrorNull;

        var cacheKeys = new string[ranges.Length];
        for (var i = 0; i < ranges.Length; i++)
        {
            if (ranges[i] is ExcelReference excelReference)
            {
                var address = XlCall.Excel(XlCall.xlfReftext, excelReference, true)?.ToString()
                    ?? throw new InvalidOperationException($"Failed to determine the address of the supplied range at index {i}.");
                cacheKeys[i] = address;
                if (GetCachedRangeData(address) == null)
                {
                    var rangeData = ExcelHelper.GetRangeValues(address);
                    CacheRangeData(address, rangeData);
                }
            }
            else if (ranges[i] is object[,] arrayArg)
            {
                var guidKey = $"array:{Guid.NewGuid()}";
                cacheKeys[i] = guidKey;
                CacheRangeData(guidKey, arrayArg);
            }
            else
            {
                var typeName = ranges[i]?.GetType().FullName ?? "null";
                // Log unexpected type (could be extended to a logger)
                System.Diagnostics.Debug.WriteLine($"[xlDuckDb] Unsupported range argument type at index {i}: {typeName}");
                throw new ArgumentException($"Unsupported range argument type at index {i}: {typeName}. Only ExcelReference and object[,] are supported.");
            }
        }

        var result = DuckDbHelper.ExecuteQuery(query, dataSource, cacheKeys);

        // Remove cached range data before returning to prevent stale data
        foreach (var cacheKey in cacheKeys)
        {
            if (!string.IsNullOrEmpty(cacheKey))
            {
                RemoveCachedRangeData(cacheKey);
            }
        }

        return result.Length == 1 ? new object[,] {{ExcelError.ExcelErrorNA}} : result;
    }

}