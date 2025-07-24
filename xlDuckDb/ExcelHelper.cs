using ExcelDna.Integration;

namespace xlDuckDb;

internal static class ExcelHelper
{
    internal static object[,] GetRangeValues(string range)
    {
        if (string.IsNullOrWhiteSpace(range))
            throw new ArgumentNullException(nameof(range), "Range string cannot be null or empty.");

        try
        {
            var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            var r = app.Range[range];
            var reference = new ExcelReference(
                r.Row - 1,
                r.Row - 1 + r.Rows.Count - 1,
                r.Column - 1,
                r.Column - 1 + r.Columns.Count - 1,
                r.Worksheet.Name);

            var content = reference.GetValue();
            return ConvertTo2DArray(content);
        }
        catch (Exception e)
        {
            throw new Exception($"Failed to read the range '{range}' from Excel.", e);
        }
    }

    private static object[,] ConvertTo2DArray(object? content)
    {
        switch (content)
        {
            case object[,] values:
                var rows = values.GetLength(0);
                var cols = values.GetLength(1);
                var result = new object[rows, cols];
                for (var i = 0; i < rows; i++)
                {
                    for (var j = 0; j < cols; j++)
                    {
                        result[i, j] = NormalizeExcelValue(values[i, j]);
                    }
                }
                return result;

            case double d:
                return new object[,] { { d } };

            case null:
                return new object[,] { { DBNull.Value } };

            default:
                return new[,] { { NormalizeExcelValue(content) } };
        }
    }

    /// <summary>
    /// Normalizes Excel cell values, converting errors and empty/missing values to DbNull.
    /// </summary>
    private static object NormalizeExcelValue(object? value) => value switch
    {
        double d => d,
        string s => s,
        bool b => b,
        ExcelError => DBNull.Value,
        ExcelEmpty => DBNull.Value,
        ExcelMissing => DBNull.Value,
        null => DBNull.Value,
        _ => value
    };
}