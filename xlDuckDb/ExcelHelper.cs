using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace xlDuckDb;

internal static class ExcelHelper
{
    /// <summary>
    /// Retrieves the values from the specified Excel range as a 2D object array.
    /// Optimization: if the returned array contains only usable CLR values
    /// (no ExcelEmpty/ExcelMissing/ExcelError/null), the original array is returned
    /// directly to avoid an unnecessary copy.
    /// </summary>
    /// <param name="range">The Excel range address (e.g., "Sheet1!A1:C10").</param>
    /// <returns>A 2D object array containing the range values.</returns>
    internal static object[,] GetRangeValues(string range)
    {
        if (string.IsNullOrWhiteSpace(range))
            throw new ArgumentNullException(nameof(range), "Range string cannot be null or empty.");

        var app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
        Microsoft.Office.Interop.Excel.Range? r = null;

        try
        {
            r = app.Range[range] as Microsoft.Office.Interop.Excel.Range
                ?? throw new InvalidOperationException($"Could not resolve range '{range}' to a Range object.");

            // Get the fully qualified address (with sheet name)
            var address = r.Address[true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, true, null];
            
            // Extract sheet name: everything before the last '!'

            var sheetDelimiterIndex = address.LastIndexOf('!');
            if (sheetDelimiterIndex <= 0)
                throw new InvalidOperationException($"Could not extract sheet name from range address '{address}'.");

            var sheetName = address[..sheetDelimiterIndex];
            var sheetRef = (ExcelReference?)XlCall.Excel(XlCall.xlSheetId, sheetName)
                ?? throw new InvalidOperationException($"Could not resolve sheet '{sheetName}'.");

            var reference = new ExcelReference(
                r.Row - 1,
                r.Row - 1 + r.Rows.Count - 1,
                r.Column - 1,
                r.Column - 1 + r.Columns.Count - 1,
                sheetRef.SheetId);

            var content = reference.GetValue();
            return ConvertTo2DArray(content);
        }
        catch (COMException ce)
        {
            throw new Exception($"COM error reading range '{range}' from Excel.", ce);
        }
        catch (Exception e)
        {
            throw new Exception($"Failed to read the range '{range}' from Excel.", e);
        }
        finally
        {
            // Release COM object to avoid Excel leaks
            if (r is not null)
            {
                try { Marshal.FinalReleaseComObject(r); } catch { /* best effort */ }
            }
        }
    }

    private static object[,] ConvertTo2DArray(object? content)
    {
        switch (content)
        {
            case object[,] values:
            {
                var rows = values.GetLength(0);
                var cols = values.GetLength(1);

                // Quick scan: if no cell requires normalization, return original array to avoid copying.
                for (var i = 0; i < rows; i++)
                {
                    for (var j = 0; j < cols; j++)
                    {
                        if (RequiresNormalization(values[i, j]))
                        {
                            // Perform copy+normalize only if a problematic cell is encountered.
                            var result = new object[rows, cols];
                            for (var ii = 0; ii < rows; ii++)
                            {
                                for (var jj = 0; jj < cols; jj++)
                                {
                                    result[ii, jj] = NormalizeExcelValue(values[ii, jj]);
                                }
                            }
                            return result;
                        }
                    }
                }

                // No normalization required; return the original array.
                return values;
            }

            case double d:
                return new object[,] { { d } };

            case null:
                return new object[,] { { DBNull.Value } };

            default:
                return new[,] { { NormalizeExcelValue(content) } };
        }
    }

    private static bool RequiresNormalization(object? v) =>
        v is null || v is ExcelError || v is ExcelEmpty || v is ExcelMissing;

    /// <summary>
    /// Normalizes Excel cell values, converting errors and empty/missing values to DBNull.
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