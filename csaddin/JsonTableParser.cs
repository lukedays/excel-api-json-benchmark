using System.Text.Json;
using ExcelDna.Integration;

namespace CsAddIn;

/// <summary>Tabela já fatiada em blocos de linhas; cada bloco é interpretado em paralelo e fica pronto sozinho.</summary>
public sealed record ParsedTable(string[] Keys, int RowCount, int ChunkRows, IReadOnlyList<Task<object[,]>> Chunks);

/// <summary>Converte um array JSON de objetos planos em blocos de matriz prontos para a C API do Excel.</summary>
/// <remarks>
/// Uma varredura linear acha o início de cada linha; depois os blocos são lidos em paralelo com
/// Utf8JsonReader (sem DOM). Células sem valor saem como ExcelEmpty. Assume chaves na mesma ordem em
/// todas as linhas e só cai na busca por nome quando a ordem difere.
/// </remarks>
public static class JsonTableParser
{
    private const int MinChunkRows = 2048;
    private static readonly object BoxedTrue = true;
    private static readonly object BoxedFalse = false;

    public static ParsedTable Parse(byte[] utf8Json)
    {
        var rowStarts = FindRowStarts(utf8Json);
        if (rowStarts.Count == 0)
        {
            return new ParsedTable([], 0, MinChunkRows, []);
        }
        // Cerca de dois blocos por núcleo: parse paralelo e gravação começando cedo. Medido de 4 mil a 65 mil
        // linhas por bloco em 100 mil linhas, a diferença ficou dentro do ruído.
        var chunkRows = Math.Max(MinChunkRows, rowStarts.Count / (Environment.ProcessorCount * 2) + 1);

        var keys = ReadKeys(utf8Json, rowStarts[0]);
        var chunks = new List<Task<object[,]>>();
        for (var first = 0; first < rowStarts.Count; first += chunkRows)
        {
            var start = first;
            var count = Math.Min(chunkRows, rowStarts.Count - first);
            chunks.Add(Task.Run(() => ParseRows(utf8Json, rowStarts, start, count, keys)));
        }
        return new ParsedTable(keys, rowStarts.Count, chunkRows, chunks);
    }

    /// <summary>Posição do '{' de cada objeto no nível 1 do array, respeitando strings e escapes.</summary>
    private static List<int> FindRowStarts(ReadOnlySpan<byte> json)
    {
        var starts = new List<int>(json.Length / 64);
        int depth = 0;
        var inString = false;
        for (var i = 0; i < json.Length; i++)
        {
            var b = json[i];
            if (inString)
            {
                if (b == (byte)'\\') i++;
                else if (b == (byte)'"') inString = false;
                continue;
            }
            switch (b)
            {
                case (byte)'"': inString = true; break;
                case (byte)'[' or (byte)'{':
                    if (b == (byte)'{' && depth == 1) starts.Add(i);
                    depth++;
                    break;
                case (byte)']' or (byte)'}': depth--; break;
            }
        }
        return starts;
    }

    private static string[] ReadKeys(byte[] json, int rowStart)
    {
        var reader = new Utf8JsonReader(json.AsSpan(rowStart));
        reader.Read(); // StartObject
        var keys = new List<string>();
        while (reader.Read() && reader.TokenType == JsonTokenType.PropertyName)
        {
            keys.Add(reader.GetString()!);
            reader.Read();
            reader.Skip();
        }
        return keys.ToArray();
    }

    private static object[,] ParseRows(byte[] json, List<int> rowStarts, int first, int count, string[] keys)
    {
        var matrix = new object[count, keys.Length];
        for (var r = 0; r < count; r++)
        {
            var reader = new Utf8JsonReader(json.AsSpan(rowStarts[first + r]));
            reader.Read(); // StartObject
            ReadRow(ref reader, keys, matrix, r);
        }
        return matrix;
    }

    private static void ReadRow(ref Utf8JsonReader reader, string[] keys, object[,] matrix, int r)
    {
        for (var c = 0; c < keys.Length; c++)
        {
            matrix[r, c] = ExcelEmpty.Value;
        }
        var column = 0;
        while (reader.Read() && reader.TokenType == JsonTokenType.PropertyName)
        {
            if (column >= keys.Length || !reader.ValueTextEquals(keys[column]))
            {
                column = Array.IndexOf(keys, reader.GetString());
            }
            reader.Read();
            var value = ReadValue(ref reader);
            if (column >= 0)
            {
                matrix[r, column] = value;
                column++;
            }
            else
            {
                column = keys.Length; // chave desconhecida: ignora e volta a procurar por nome
            }
        }
    }

    private static object ReadValue(ref Utf8JsonReader reader)
    {
        switch (reader.TokenType)
        {
            case JsonTokenType.String: return reader.GetString()!;
            case JsonTokenType.Number: return reader.GetDouble();
            case JsonTokenType.True: return BoxedTrue;
            case JsonTokenType.False: return BoxedFalse;
            case JsonTokenType.Null: return ExcelEmpty.Value;
            default:
                reader.Skip(); // objeto/array aninhado não vira célula
                return "(aninhado)";
        }
    }
}
