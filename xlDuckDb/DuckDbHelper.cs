using System.Numerics;
using System.Text;
using DuckDB.NET.Data;
using DuckDB.NET.Native;
using ExcelDna.Integration;

namespace xlDuckDb;

public static class DuckDbHelper
{
    public static object[,] ExecuteQuery(string query, string dataSource = "")
    {
        if (string.IsNullOrEmpty(query))
            throw new ArgumentException("Value cannot be null or empty.", nameof(query));

        using var duckDbConnection =
            new DuckDBConnection(
                $"Data Source = {(string.IsNullOrEmpty(dataSource) ? ":memory:" : dataSource)}");
        duckDbConnection.Open();

        using var command = duckDbConnection.CreateCommand();
        command.CommandText = query;
        using var reader = command.ExecuteReader();
        if (!reader.HasRows || reader.IsClosed) return new object[,] {{ExcelError.ExcelErrorNA}};
        var rows = new List<object[]>();

        var headerRow = new object[reader.FieldCount];
        // We do special handling for certain data types
        // so we need to know the types of the columns
        var bigIntField = new bool[reader.FieldCount];
        var blobField = new bool[reader.FieldCount];
        var decimalField = new bool[reader.FieldCount];
        var floatField = new bool[reader.FieldCount];
        var timeSpanField = new bool[reader.FieldCount];
        var timeTzField = new bool[reader.FieldCount];
        var dateOnlyField = new bool[reader.FieldCount];
        var uuidField = new bool[reader.FieldCount];

        for (var i = 0; i < reader.FieldCount; i++)
        {
            headerRow[i] = reader.GetName(i);
            var fieldType = reader.GetFieldType(i);
            bigIntField[i] = fieldType == typeof(BigInteger) ||
                             fieldType == typeof(ulong);
            blobField[i] = fieldType == typeof(Stream);
            decimalField[i] = fieldType == typeof(decimal);
            floatField[i] = fieldType == typeof(float);
            timeSpanField[i] = fieldType == typeof(TimeSpan);
            timeTzField[i] = fieldType == typeof(DateTimeOffset);
            dateOnlyField[i] = fieldType == typeof(DateOnly);
            uuidField[i] = fieldType == typeof(Guid);
        }

        // Add the first row of column names
        rows.Add(headerRow);

        // Loop through the results, converting data as necessary
        foreach (var _ in reader)
        {
            var rowData = new object[reader.FieldCount];
            for (var i = 0; i < reader.FieldCount; i++)
                if (bigIntField[i])
                {
                    rowData[i] = reader.GetInt64(i);
                }
                else if (blobField[i])
                {
                    var stream = reader.GetStream(i);
                    using var streamReader = new StreamReader(stream, Encoding.UTF8);
                    rowData[i] = streamReader.ReadToEnd();
                }
                else if (decimalField[i])
                {
                    rowData[i] = decimal.ToDouble(reader.GetDecimal(i));
                }
                else if (floatField[i])
                {
                    rowData[i] = (double) reader.GetFloat(i);
                }
                else if (timeSpanField[i])
                {
                    var tsv = reader.GetValue(i);
                    rowData[i] = tsv switch
                    {
                        DuckDBTimeOnly tsvDdb => new DateTime(1899, 12, 30) + TimeSpan.FromTicks(tsvDdb.Ticks),
                        TimeSpan tsvNet => new DateTime(1899, 12, 30) + tsvNet,
                        _ => tsv
                    };
                }
                else if (timeTzField[i])
                {
                    rowData[i] = new DateTime(1899, 12, 30) +
                                 TimeSpan.FromTicks(((DateTimeOffset) reader.GetValue(i)).Ticks);
                }
                else if (dateOnlyField[i])
                {
                    rowData[i] = reader.GetDateTime(i);
                }
                else if (uuidField[i])
                {
                    rowData[i] = (reader.GetGuid(i)).ToString();
                }
                else
                {
                    rowData[i] = reader.GetValue(i);
                }

            // Convert DBNulls and nan/inf to ExcelNA and ExcelNum
            for (var i = 0; i < reader.FieldCount; i++)
                switch (rowData[i])
                {
                    case DBNull _:
                        rowData[i] = ExcelError.ExcelErrorNA;
                        break;
                    case double d when double.IsNaN(d) || double.IsInfinity(d):
                    case float f when float.IsNaN(f) || float.IsInfinity(f):
                        rowData[i] = ExcelError.ExcelErrorNum;
                        break;
                }

            rows.Add(rowData);
        }

        return rows.AsMultiDimensionalArray();
    }
}