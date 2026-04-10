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
                
        var base64Range = parameters[0].GetValue<string>() ?? 
                    throw new ArgumentException("Range parameter cannot be null");

        var range = System.Text.Encoding.Unicode.GetString(Convert.FromBase64String(base64Range));

        var data = ExcelHelper.GetRangeValues(range) ?? throw new ArgumentException("Retrieved range cannot be null");
        var rowLength = data.GetLength(0);
        var colLength = data.GetLength(1);

        if (rowLength < 2)
            throw new ArgumentException("At least two rows required - headers and data types.");
        if (colLength < 1)
            throw new ArgumentException("At least one column required.");

        // Use first row for header names
        // Use second row for data types
        var dataTypes = new Type[colLength];
        var columnNames = new string[colLength];
        var columns = new List<ColumnInfo>(colLength);
        var nameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < colLength; i++)
        {
            dataTypes[i] = data[1, i] switch
            {
                double => typeof(double),
                bool => typeof(bool),
                _ => typeof(string)
            };
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
            columns.Add(new ColumnInfo(name, dataTypes[i]));
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
                // Write default values for the specific type
                if (types[i] == typeof(double))
                    writers[i].WriteValue(double.NaN, rowIndex);
                else if (types[i] == typeof(bool))
                    writers[i].WriteValue(false, rowIndex);
                else
                    writers[i].WriteValue(string.Empty, rowIndex);
            }
        }
    }
}