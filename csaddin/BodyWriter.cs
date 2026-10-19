using ExcelDna.Integration;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace CsAddIn;

/// <summary>Grava os blocos da tabela via C API (xlSet), cada um assim que o parse dele termina.</summary>
public static class BodyWriter
{
    public static void Write(Range target, ParsedTable table)
    {
        var sheetRef = SheetRef(target);
        int firstRow = target.Row - 1, firstColumn = target.Column - 1;
        var lastColumn = firstColumn + table.Keys.Length - 1;

        for (var i = 0; i < table.Chunks.Count; i++)
        {
            var chunk = table.Chunks[i].GetAwaiter().GetResult();
            var top = firstRow + i * table.ChunkRows;
            new ExcelReference(top, top + chunk.GetLength(0) - 1, firstColumn, lastColumn, sheetRef).SetValue(chunk);
        }
    }

    /// <summary>Apaga os valores do intervalo: xlSet com um único valor vazio vale para todas as células.</summary>
    public static void Clear(Range target)
    {
        int top = target.Row - 1, left = target.Column - 1;
        new ExcelReference(top, top + target.Rows.Count - 1, left, left + target.Columns.Count - 1, SheetRef(target))
            .SetValue(ExcelEmpty.Value);
    }

    private static string SheetRef(Range target)
    {
        var sheet = (Microsoft.Office.Interop.Excel.Worksheet)target.Worksheet;
        var workbook = (Microsoft.Office.Interop.Excel.Workbook)sheet.Parent;
        return $"[{workbook.Name}]{sheet.Name}";
    }
}
