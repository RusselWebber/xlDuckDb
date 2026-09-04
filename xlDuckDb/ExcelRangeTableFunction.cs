using DuckDB.NET.Data;
using DuckDB.NET.Data.DataChunk.Writer;
using DuckDB.NET.Native;
using System.Diagnostics.CodeAnalysis;

namespace xlDuckDb;

internal record RowDataAndTypes(object[] Data, Type[] Types);

internal static class ExcelRangeTableFunctions
{
    internal static TableFunction ResultCallback(IReadOnlyList<IDuckDBValueReader> parameters)
    {
        if (parameters == null || parameters.Count == 0)
            throw new ArgumentException("Parameters cannot be null or empty", nameof(parameters));
                
        var base64CacheKey = parameters[0].GetValue<string>() ?? 
                    throw new ArgumentException("Cache key parameter cannot be null");

        var cacheKey = System.Text.Encoding.Unicode.GetString(Convert.FromBase64String(base64CacheKey));

        // Get data from cache
        var data = xlAddIn.GetCachedRangeData(cacheKey) ?? 
                   throw new ArgumentException("Range data not found in cache");
        
        var rowLength = data.GetLength(0);
        var colLength = data.GetLength(1);

        if (rowLength < 2)
            throw new ArgumentException("At least two rows required - headers and data types.");
        if (colLength < 1)
            throw new ArgumentException("At least one column required.");

        // Use first row for header names
        var columnNames = new string[colLength];
        var nameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < colLength; i++)
        {
            var originalName = data[0, i]?.ToString() ?? string.Empty;
            var name = string.IsNullOrWhiteSpace(originalName) ? $"col_{i + 1}" : originalName;
            if (nameCounts.TryGetValue(name, out var count))
            {
                nameCounts[name] = count + 1;
                name = $"{name}_{count}";
            }
            else
            {
                nameCounts[name] = 1;
            }
            columnNames[i] = name;
        }

        // For data types, default to string
        // and let the first double or bool value 
        // override the type for that column
        var dataTypes = new Type[colLength];
        var columns = new List<ColumnInfo>(colLength);
        for (var i = 0; i < colLength; i++)
        {
            dataTypes[i] = typeof(string); // Default to string

            for (var j = 1; j < rowLength; j++)
            {
                if (data[j, i] is double)
                {
                    dataTypes[i] = typeof(double);
                    break;
                }
                else if (data[j, i] is bool)
                {
                    dataTypes[i] = typeof(bool);
                    break;
                }
            }

            columns.Add(new ColumnInfo(columnNames[i], dataTypes[i]));
        }

        var dataList = new List<RowDataAndTypes>();

        for (var i = 1; i < rowLength; i++)
        {
            var row = new object[colLength];
            for (var j = 0; j < colLength; j++)
            {
                row[j] = data[i, j];                    
            }
            dataList.Add(new RowDataAndTypes(row, dataTypes));  
        }

        return new TableFunction(columns, dataList);
    }

    [Experimental("DuckDBNET001")]
    internal static void MapperCallback(object? item, IDuckDBDataWriter[] writers, ulong rowIndex)
    {
        if (item == null) return;
            
        var (row, types) = (RowDataAndTypes)item;
        var colLength = row.Length;
            
        for (var i = 0; i < colLength; i++)
        {
            try
            {
                if (row[i] == DBNull.Value)
                {
                    // Write null value
                    writers[i].WriteNull(rowIndex);
                    continue;
                }

                switch (types[i])
                {
                    case { } t when t == typeof(double):
                        writers[i].WriteValue(row[i] is double d ? d : double.NaN, rowIndex);
                        break;
                    case { } t when t == typeof(bool):
                        writers[i].WriteValue(row[i] is true, rowIndex);
                        break;
                    default:
                        writers[i].WriteValue(row[i].ToString() ?? string.Empty, rowIndex);
                        break;
                }
            }
            catch (Exception)
            {
                // Write null value on error
                writers[i].WriteNull(rowIndex);
            }
        }
    }
}