using System.Globalization;
using System.Text;
using Writer.Formats.Xlsx;

namespace Writer.Formats.Compat;

/// <summary>Plain text, CSV and TSV: decoding, and the two shapes they take (paragraphs, or one sheet).</summary>
public static class TextFiles
{
    /// <summary>Text from bytes: a BOM wins; then strict UTF-8; then GB18030, which reads the GBK/GB2312 files
    /// Chinese Windows still writes; Latin-1 as the last resort.</summary>
    public static string Decode(byte[] bytes)
    {
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3);
        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        try { return new UTF8Encoding(false, true).GetString(bytes); }
        catch (DecoderFallbackException) { }
        try { return Encoding.GetEncoding(54936).GetString(bytes); }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or DecoderFallbackException) { return Encoding.Latin1.GetString(bytes); }
    }

    /// <summary>One paragraph per line; tabs become spaces.</summary>
    public static List<Block> Paragraphs(string text) =>
        text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').Select(line => Block.Paragraph(Inline.Esc(line.Replace('\t', ' ')))).ToList();

    /// <summary>A sheet from delimited text. The delimiter is the one given, else the most frequent of tab, semicolon and
    /// comma on the first line. Numbers, TRUE/FALSE and ISO dates are typed; a number with a leading zero stays text.</summary>
    public static SheetModel Csv(string text, char? delimiter)
    {
        var end = text.IndexOfAny(['\r', '\n']);
        var first = end < 0 ? text : text[..end];
        var sep = delimiter ?? new[] { '\t', ';', ',' }.OrderByDescending(c => first.Count(ch => ch == c)).ThenBy(c => c == ',' ? 0 : 1).First();
        var sheet = new SheetModel();
        var rows = XlsxCells.ParseCsv(text, sep);
        for (var r = 0; r < rows.Count; r++)
            for (var c = 0; c < rows[r].Count; c++)
            {
                var raw = rows[r][c];
                if (raw.Length == 0) continue;
                var cell = new CellModel { Row = r + 1, Col = c + 1, Value = Typed(raw) };
                if (cell.Value is DateTime) cell.Format = "yyyy-mm-dd";
                sheet.Cells.Add(cell);
            }
        return sheet;
    }

    internal static object Typed(string raw)
    {
        var s = raw.Trim();
        if (s.Equals("true", StringComparison.OrdinalIgnoreCase)) return true;
        if (s.Equals("false", StringComparison.OrdinalIgnoreCase)) return false;
        var leadingZero = s.Length > 1 && s[0] == '0' && s[1] != '.';
        if (!leadingZero && double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d) && !s.Contains(',')) return d;
        if (DateTime.TryParseExact(s, ["yyyy-MM-dd", "yyyy/MM/dd", "yyyy-M-d", "yyyy/M/d"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) return date;
        return raw;
    }
}
