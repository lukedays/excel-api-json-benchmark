using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace CsAddIn;

/// <summary>Dimensiona uma tabela nomeada para caber exatamente nos dados, crescendo ou encolhendo.</summary>
public static class TableWriter
{
    /// <param name="writeBody">Recebe o intervalo do corpo (rowCount x headers.Length) e grava os valores.</param>
    public static void Write(Worksheet sheet, string tableName, object[] headers, int rowCount, Action<Range> writeBody)
    {
        var colCount = headers.Length;
        var headerRow = new object[1, colCount];
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
                writeBody(newRange.Offset[1, 0].Resize[rowCount, colCount]);
            }
            table = sheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, newRange, Type.Missing, XlYesNoGuess.xlYes);
            table.Name = tableName;
            PhaseTimer.Mark("create");
            return;
        }

        // Cabeçalho e corpo vão em gravações separadas: um bloco único sobre a tabela inteira a desfaz.
        var oldRange = table.Range;
        int oldRows = oldRange.Rows.Count, oldCols = oldRange.Columns.Count;
        table.Resize(newRange);
        PhaseTimer.Mark("resize");
        table.HeaderRowRange.Value2 = headerRow;
        PhaseTimer.Mark("header");
        if (rowCount == 0)
        {
            table.DataBodyRange.ClearContents();
        }
        else
        {
            writeBody(table.DataBodyRange);
        }
        PhaseTimer.Mark("body");

        // Fora da tabela sobram só valores; basta apagá-los.
        var newRows = newRange.Rows.Count;
        if (oldRows > newRows)
        {
            BodyWriter.Clear(oldRange.Offset[newRows, 0].Resize[oldRows - newRows, oldCols]);
        }
        if (oldCols > colCount)
        {
            BodyWriter.Clear(oldRange.Offset[0, colCount].Resize[oldRows, oldCols - colCount]);
        }
        PhaseTimer.Mark("clear");
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
