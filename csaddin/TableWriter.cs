using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace CsAddIn;

/// <summary>Grava cabeçalhos e corpo numa tabela nomeada, crescendo ou encolhendo para caber exatamente.</summary>
public static class TableWriter
{
    public static void Write(Worksheet sheet, string tableName, object?[] headers, object?[,] body, Action<Range, object?[,]> writeBody)
    {
        var rowCount = body.GetLength(0);
        var colCount = headers.Length;
        var headerRow = new object?[1, colCount];
        for (var c = 0; c < colCount; c++)
        {
            headerRow[0, c] = headers[c];
        }

        var table = FindTable(sheet, tableName);
        var anchor = table is null ? sheet.Range["A1"] : (Range)table.Range.Cells[1, 1];
        // Tabela exige ao menos uma linha de dados.
        var newRange = anchor.Resize[Math.Max(rowCount, 1) + 1, colCount];

        if (table is null)
        {
            ((Range)newRange.Rows[1]).Value2 = headerRow;
            if (rowCount > 0)
            {
                writeBody(newRange.Offset[1, 0].Resize[rowCount, colCount], body);
            }
            table = sheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, newRange, Type.Missing, XlYesNoGuess.xlYes);
            table.Name = tableName;
            return;
        }

        // Cabeçalho e corpo vão em gravações separadas: um bloco único sobre a tabela inteira a desfaz.
        var oldRange = table.Range;
        int oldRows = oldRange.Rows.Count, oldCols = oldRange.Columns.Count;
        table.Resize(newRange);
        table.HeaderRowRange.Value2 = headerRow;
        if (rowCount == 0)
        {
            table.DataBodyRange.ClearContents();
        }
        else
        {
            writeBody(table.DataBodyRange, body);
        }

        var newRows = newRange.Rows.Count;
        if (oldRows > newRows)
        {
            oldRange.Offset[newRows, 0].Resize[oldRows - newRows, oldCols].Clear();
        }
        if (oldCols > colCount)
        {
            oldRange.Offset[0, colCount].Resize[oldRows, oldCols - colCount].Clear();
        }
    }

    private static ListObject? FindTable(Worksheet sheet, string tableName)
    {
        foreach (ListObject table in sheet.ListObjects)
        {
            if (table.Name == tableName)
            {
                return table;
            }
        }
        return null;
    }
}
