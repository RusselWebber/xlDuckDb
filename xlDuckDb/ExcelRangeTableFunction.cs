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
                
        var range = parameters[0].GetValue<string>() ?? 
                    throw new ArgumentException("Range parameter cannot be null");

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
        for (var i = 0; i < colLength; i++)
        {
            dataTypes[i] = data[1, i] switch
            {
                double => typeof(double),
                bool => typeof(bool),
                _ => typeof(string)
            };
            columnNames[i] = data[0, i].ToString() ?? string.Empty;
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