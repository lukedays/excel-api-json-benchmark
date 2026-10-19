using ExcelDna.Integration;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace CsAddIn;

/// <summary>Formas de gravar a matriz de valores num intervalo já dimensionado.</summary>
public static class BodyWriters
{
    public static void Com(Range target, object?[,] body) => target.Value2 = body;

    /// <summary>Grava via C API (xlSet), sem passar o array pelo COM.</summary>
    public static void CApi(Range target, object?[,] body)
    {
        int rows = body.GetLength(0), cols = body.GetLength(1);
        var values = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < cols; c++)
            {
                values[r, c] = body[r, c] ?? ExcelEmpty.Value;
            }
        }

        var sheet = (Microsoft.Office.Interop.Excel.Worksheet)target.Worksheet;
        var workbook = (Microsoft.Office.Interop.Excel.Workbook)sheet.Parent;
        var reference = new ExcelReference(
            target.Row - 1, target.Row + rows - 2, target.Column - 1, target.Column + cols - 2,
            $"[{workbook.Name}]{sheet.Name}");
        reference.SetValue(values);
    }
}
