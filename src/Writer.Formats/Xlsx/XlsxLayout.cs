using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

/// <summary>Sheet-level layout: merged ranges, column widths, row heights, frozen panes and the autofilter.</summary>
static class XlsxLayout
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    public static string? Merges(Worksheet ws)
    {
        var refs = ws.GetFirstChild<MergeCells>()?.Elements<MergeCell>().Select(m => m.Reference?.Value).OfType<string>().ToList();
        if (refs is not { Count: > 0 }) return null;
        return NodeJson.Compact(w =>
        {
            w.WriteStartArray();
            foreach (var r in refs) w.WriteStringValue(r);
            w.WriteEndArray();
        });
    }

    public static void SetMerges(Worksheet ws, string json)
    {
        var boxes = new List<(string Ref, (int Col1, int Row1, int Col2, int Row2) Box)>();
        using (var parsed = JsonDocument.Parse(json))
        {
            if (parsed.RootElement.ValueKind != JsonValueKind.Array || parsed.RootElement.EnumerateArray().Any(e => e.ValueKind != JsonValueKind.String))
                throw new WriterException(ErrorCode.Validation, "merges must be a JSON array of ranges", "Example: [\"A1:C1\",\"B3:B5\"]");
            foreach (var e in parsed.RootElement.EnumerateArray())
            {
                var box = Range(e.GetString()!, "merges");
                if (box.Col1 == box.Col2 && box.Row1 == box.Row2)
                    throw new WriterException(ErrorCode.Validation, $"merges: '{e.GetString()}' is a single cell", "A merge covers at least two cells, e.g. A1:C1.");
                boxes.Add((e.GetString()!, box));
            }
        }
        for (var i = 0; i < boxes.Count; i++)
            for (var j = i + 1; j < boxes.Count; j++)
                if (Overlaps(boxes[i].Box, boxes[j].Box))
                    throw new WriterException(ErrorCode.Validation, $"merges: {boxes[i].Ref} overlaps {boxes[j].Ref}", "Merged ranges cannot share cells; drop or shrink one of them.");
        ws.GetFirstChild<MergeCells>()?.Remove();
        if (boxes.Count == 0) return;
        var merges = new MergeCells { Count = (uint)boxes.Count };
        foreach (var (_, b) in boxes) merges.Append(new MergeCell { Reference = XlsxCells.Reference(b.Col1, b.Row1) + ":" + XlsxCells.Reference(b.Col2, b.Row2) });
        ws.AddChild(merges);
    }

    static bool Overlaps((int Col1, int Row1, int Col2, int Row2) a, (int Col1, int Row1, int Col2, int Row2) b) =>
        a.Col1 <= b.Col2 && b.Col1 <= a.Col2 && a.Row1 <= b.Row2 && b.Row1 <= a.Row2;

    public static string? Widths(Worksheet ws)
    {
        var entries = new List<(int Col, double Width)>();
        foreach (var c in ws.GetFirstChild<Columns>()?.Elements<Column>() ?? [])
        {
            if (c.CustomWidth?.Value != true || c.Width?.Value is not { } width || c.Min?.Value is not { } min || c.Max?.Value is not { } max) continue;
            for (var i = min; i <= max && i <= 16384; i++) entries.Add(((int)i, width));
        }
        if (entries.Count == 0) return null;
        return NodeJson.Compact(w =>
        {
            w.WriteStartObject();
            foreach (var (col, width) in entries.OrderBy(e => e.Col)) w.WriteNumber(XlsxCells.ColumnName(col), width);
            w.WriteEndObject();
        });
    }

    public static void SetWidths(Worksheet ws, string json)
    {
        var cols = ws.GetFirstChild<Columns>();
        foreach (var (col, width) in Entries(json, "widths", "{\"A\":12.5,\"C\":30}", 255, ColumnOf))
        {
            var existing = cols is null ? null : Isolate(cols, col);
            if (width is null)
            {
                if (existing is null) continue;
                existing.Width = null;
                existing.CustomWidth = null;
                existing.BestFit = null;
                if (existing.GetAttributes().All(a => a.LocalName is "min" or "max")) existing.Remove();
                continue;
            }
            existing ??= ColumnAt(ws, ref cols, col);
            existing.Width = width;
            existing.CustomWidth = true;
        }
        if (cols is not null && !cols.HasChildren) cols.Remove();
    }

    /// <summary>The col element for exactly this column: an existing span isolated, or a new one in order (creating cols when needed).</summary>
    internal static Column ColumnAt(Worksheet ws, ref Columns? cols, int col)
    {
        if (cols is not null && Isolate(cols, col) is { } existing) return existing;
        if (cols is null) ws.AddChild(cols = new Columns());
        var column = new Column { Min = (uint)col, Max = (uint)col };
        var next = cols.Elements<Column>().FirstOrDefault(c => c.Min?.Value > (uint)col);
        if (next is null) cols.Append(column);
        else cols.InsertBefore(column, next);
        return column;
    }

    /// <summary>The col element covering exactly this column, splitting a wider span so its other columns keep their settings.</summary>
    internal static Column? Isolate(Columns cols, int col)
    {
        var span = cols.Elements<Column>().FirstOrDefault(c => c.Min?.Value <= (uint)col && (uint)col <= c.Max?.Value);
        if (span is null) return null;
        if (span.Min!.Value < (uint)col)
        {
            var before = (Column)span.CloneNode(true);
            before.Max = (uint)col - 1;
            span.InsertBeforeSelf(before);
        }
        if (span.Max!.Value > (uint)col)
        {
            var after = (Column)span.CloneNode(true);
            after.Min = (uint)col + 1;
            span.InsertAfterSelf(after);
        }
        span.Min = (uint)col;
        span.Max = (uint)col;
        return span;
    }

    public static string? Heights(SheetData data)
    {
        var rows = data.Elements<Row>().Where(r => r.CustomHeight?.Value == true && r.Height?.Value is not null && r.RowIndex?.Value is not null).ToList();
        if (rows.Count == 0) return null;
        return NodeJson.Compact(w =>
        {
            w.WriteStartObject();
            foreach (var r in rows) w.WriteNumber(r.RowIndex!.Value.ToString(Inv), r.Height!.Value);
            w.WriteEndObject();
        });
    }

    public static void SetHeights(SheetData data, string json)
    {
        foreach (var (index, height) in Entries(json, "heights", "{\"3\":24}", 409.5, RowOf))
        {
            if (height is { } h)
            {
                var row = XlsxCells.GetOrCreateRow(data, index);
                row.Height = h;
                row.CustomHeight = true;
                continue;
            }
            if (XlsxCells.FindRow(data, index) is not { } existing) continue;
            existing.Height = null;
            existing.CustomHeight = null;
            if (!existing.HasChildren && existing.GetAttributes().All(a => a.LocalName == "r")) existing.Remove();
        }
    }

    static List<(int Key, double? Value)> Entries(string json, string prop, string example, double max, Func<string, string, int> key)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Object)
            throw new WriterException(ErrorCode.Validation, $"{prop} must be a JSON object", $"Example: {example}");
        var list = new List<(int, double?)>();
        foreach (var p in parsed.RootElement.EnumerateObject())
        {
            double? value = p.Value.ValueKind switch
            {
                JsonValueKind.Null => null,
                JsonValueKind.String when p.Value.GetString() == "default" => null,
                JsonValueKind.Number when p.Value.TryGetDouble(out var d) && d > 0 && d <= max => d,
                _ => throw new WriterException(ErrorCode.Validation, $"{prop}: '{p.Name}' needs a number above 0 and up to {max}, or null to reset it", $"Example: {example}"),
            };
            list.Add((key(p.Name, prop), value));
        }
        return list;
    }

    static int ColumnOf(string key, string prop)
    {
        var k = key.Trim();
        if (k.Length is >= 1 and <= 3 && k.All(char.IsAsciiLetter) && XlsxCells.ColumnIndex(k) <= 16384) return XlsxCells.ColumnIndex(k);
        throw new WriterException(ErrorCode.Validation, $"{prop}: '{key}' is not a column letter", "Keys are column letters like A or AB.");
    }

    static int RowOf(string key, string prop) =>
        int.TryParse(key.Trim(), NumberStyles.None, Inv, out var row) && row is >= 1 and <= 1048576
            ? row
            : throw new WriterException(ErrorCode.Validation, $"{prop}: '{key}' is not a row number", "Keys are row numbers like 3.");

    public static string? Freeze(Worksheet ws)
    {
        var pane = ws.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.Pane;
        if (pane is null || pane.State?.InnerText is not ("frozen" or "frozenSplit")) return null;
        return pane.TopLeftCell?.Value ?? XlsxCells.Reference((int)(pane.HorizontalSplit?.Value ?? 0) + 1, (int)(pane.VerticalSplit?.Value ?? 0) + 1);
    }

    public static void SetFreeze(Worksheet ws, string value)
    {
        var (col, row) = IsNone(value) ? (1, 1) : Cell(value, "freeze");
        var views = ws.GetFirstChild<SheetViews>();
        var view = views?.GetFirstChild<SheetView>();
        if (view is not null)
        {
            view.Pane = null;
            foreach (var s in view.Elements<Selection>().Where(s => s.Pane is not null).ToList()) s.Remove();
        }
        if (col == 1 && row == 1) return;
        if (view is null)
        {
            if (views is null) ws.AddChild(views = new SheetViews());
            view = views.AppendChild(new SheetView { WorkbookViewId = 0U });
        }
        view.Pane = new Pane
        {
            HorizontalSplit = col > 1 ? new DoubleValue(col - 1.0) : null,
            VerticalSplit = row > 1 ? new DoubleValue(row - 1.0) : null,
            TopLeftCell = XlsxCells.Reference(col, row),
            ActivePane = col > 1 && row > 1 ? PaneValues.BottomRight : row > 1 ? PaneValues.BottomLeft : PaneValues.TopRight,
            State = PaneStateValues.Frozen,
        };
    }

    /// <summary>"false" when the sheet hides its grid lines; null when it shows them (Excel's default).</summary>
    public static string? Gridlines(Worksheet ws) => ws.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.ShowGridLines?.Value == false ? "false" : null;

    public static void SetGridlines(Worksheet ws, string value)
    {
        var view = ws.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
        if (value == "true")
        {
            if (view is not null) view.ShowGridLines = null;
            return;
        }
        if (view is null)
        {
            var views = ws.GetFirstChild<SheetViews>();
            if (views is null) ws.AddChild(views = new SheetViews());
            view = views.AppendChild(new SheetView { WorkbookViewId = 0U });
        }
        view.ShowGridLines = false;
    }

    public static string? Filter(Worksheet ws) => ws.GetFirstChild<AutoFilter>()?.Reference?.Value;

    /// <summary>Sets the autofilter and the hidden _xlnm._FilterDatabase name Excel keeps for it.</summary>
    public static void SetFilter(XlsxDocument doc, Sheet sheet, Worksheet ws, string value)
    {
        var box = IsNone(value) ? default : Range(value, "filter");
        var reference = IsNone(value) ? null : XlsxCells.Reference(box.Col1, box.Row1) + ":" + XlsxCells.Reference(box.Col2, box.Row2);
        ws.GetFirstChild<AutoFilter>()?.Remove();
        if (reference is not null) ws.AddChild(new AutoFilter { Reference = reference });
        var index = (uint)sheet.Parent!.Elements<Sheet>().TakeWhile(s => !ReferenceEquals(s, sheet)).Count();
        var workbook = doc.Workbook.Workbook!;
        var names = workbook.DefinedNames;
        foreach (var n in names?.Elements<DefinedName>().Where(n => n.Name?.Value == "_xlnm._FilterDatabase" && n.LocalSheetId?.Value == index).ToList() ?? []) n.Remove();
        if (reference is null) return;
        names ??= workbook.DefinedNames = new DefinedNames();
        names.Append(new DefinedName(XlsxRefs.Formula(sheet.Name!.Value!, box)) { Name = "_xlnm._FilterDatabase", LocalSheetId = index, Hidden = true });
    }

    static bool IsNone(string value) => value.Trim().Length == 0 || value.Trim().Equals("none", StringComparison.OrdinalIgnoreCase);

    static (int Col1, int Row1, int Col2, int Row2) Range(string reference, string prop)
    {
        try { return XlsxCells.ParseRange(reference); }
        catch (WriterException) { throw new WriterException(ErrorCode.Validation, $"{prop}: '{reference}' is not a cell range", "Use a range like A1:D20."); }
    }

    static (int Col, int Row) Cell(string reference, string prop)
    {
        try { return XlsxCells.Parse(reference); }
        catch (WriterException)
        {
            throw new WriterException(ErrorCode.Validation, $"{prop}: '{reference}' is not a cell reference",
                "Name the top-left cell of the scrolling area: A2 freezes row 1, B1 column A, B2 both; none unfreezes.");
        }
    }

    /// <summary>Column widths and row heights of one sheet in EMU, for placing drawings. Excel's own arithmetic:
    /// a column of stored width w is trunc((256w + 18) / 256 * 7) pixels at 96 dpi, a row of h points is h * 12700 EMU.</summary>
    public sealed class Grid
    {
        const double DefaultColumnWidth = 9.140625; // 64 px
        const double DefaultRowHeight = 15; // points
        readonly List<(uint Min, uint Max, long Emu)> _cols = [];
        readonly Dictionary<uint, long> _rows = [];
        readonly long _defaultCol, _defaultRow;

        public Grid(Worksheet ws)
        {
            var format = ws.GetFirstChild<SheetFormatProperties>();
            _defaultCol = ColumnEmu(format?.DefaultColumnWidth?.Value ?? DefaultColumnWidth);
            _defaultRow = RowEmu(format?.DefaultRowHeight?.Value ?? DefaultRowHeight);
            foreach (var c in ws.GetFirstChild<Columns>()?.Elements<Column>() ?? [])
                if (c.Min?.Value is { } min && c.Max?.Value is { } max)
                    _cols.Add((min, max, c.Hidden?.Value == true ? 0 : c.Width?.Value is { } w ? ColumnEmu(w) : _defaultCol));
            foreach (var r in ws.GetFirstChild<SheetData>()?.Elements<Row>() ?? [])
                if (r.RowIndex?.Value is { } i && (r.Hidden?.Value == true || r.Height?.Value is not null))
                    _rows[i] = r.Hidden?.Value == true ? 0 : RowEmu(r.Height!.Value);
        }

        static long ColumnEmu(double width) => (long)Math.Truncate((256 * width + 18) / 256 * 7) * Units.EmuPerPx;
        static long RowEmu(double points) => (long)Math.Round(points * Units.EmuPerPt);

        /// <summary>Width of the 0-based column.</summary>
        public long Column(int index)
        {
            foreach (var c in _cols) if (c.Min <= index + 1 && index + 1 <= c.Max) return c.Emu;
            return _defaultCol;
        }

        /// <summary>Height of the 0-based row.</summary>
        public long Row(int index) => _rows.GetValueOrDefault((uint)index + 1, _defaultRow);

        public long X(int col, long offset) => Sum(col, Column) + offset;
        public long Y(int row, long offset) => Sum(row, Row) + offset;
        public (int Col, long Offset) ColumnAt(long x) => Locate(x, Column, 16384);
        public (int Row, long Offset) RowAt(long y) => Locate(y, Row, 1048576);

        static long Sum(int count, Func<int, long> size)
        {
            long total = 0;
            for (var i = 0; i < count; i++) total += size(i);
            return total;
        }

        static (int, long) Locate(long emu, Func<int, long> size, int limit)
        {
            long start = 0;
            for (var i = 0; i < limit - 1; i++)
            {
                var s = size(i);
                if (emu < start + s) return (i, emu - start);
                start += s;
            }
            return (limit - 1, emu - start);
        }
    }
}

/// <summary>Cell references in formulas: Sheet1!$A$2:$A$6 in the file, A2:A6 or 'Other sheet'!A2:A6 for users.</summary>
static class XlsxRefs
{
    /// <summary>Parses "A2:A6", "B1" or "Sheet1!A2:A6" against the current sheet. The sheet must exist.</summary>
    public static (string Sheet, (int Col1, int Row1, int Col2, int Row2) Box) Parse(XlsxDocument doc, string currentSheet, string input, string prop)
    {
        var text = input.Trim();
        var bang = text.LastIndexOf('!');
        var sheet = bang < 0 ? currentSheet : Unquote(text[..bang]);
        var range = bang < 0 ? text : text[(bang + 1)..];
        var actual = doc.Sheets.Select(s => s.Sheet.Name?.Value).FirstOrDefault(n => string.Equals(n, sheet, StringComparison.OrdinalIgnoreCase))
            ?? throw new WriterException(ErrorCode.Validation, $"{prop}: there is no sheet called '{sheet}'",
                $"Sheets: {string.Join(", ", doc.Sheets.Select(s => s.Sheet.Name?.Value))}. Write A2:A6 for this sheet or Sheet1!A2:A6 for another.");
        try { return (actual, XlsxCells.ParseRange(range)); }
        catch (WriterException)
        {
            throw new WriterException(ErrorCode.Validation, $"{prop}: '{input}' is not a cell range", "Use a cell or a range like B1 or B2:B6, with a sheet if needed: Sheet1!B2:B6.");
        }
    }

    public static string Formula(string sheet, (int Col1, int Row1, int Col2, int Row2) b) =>
        Quote(sheet) + "!" + Absolute(b.Col1, b.Row1) + (b.Col1 == b.Col2 && b.Row1 == b.Row2 ? "" : ":" + Absolute(b.Col2, b.Row2));

    /// <summary>A stored formula the way users write it: no $ signs, and no sheet name when it is the current sheet.</summary>
    public static string Display(string formula, string currentSheet)
    {
        var bang = formula.LastIndexOf('!');
        var range = (bang < 0 ? formula : formula[(bang + 1)..]).Replace("$", "");
        if (bang < 0) return range;
        var sheet = Unquote(formula[..bang]);
        return string.Equals(sheet, currentSheet, StringComparison.OrdinalIgnoreCase) ? range : Quote(sheet) + "!" + range;
    }

    /// <summary>True when the text names one cell (B1, $B$1, Sheet1!B1) on a sheet that exists.</summary>
    public static bool IsCell(XlsxDocument doc, string text)
    {
        var t = text.Trim();
        var bang = t.LastIndexOf('!');
        var range = bang < 0 ? t : t[(bang + 1)..];
        var r = range.Replace("$", "");
        var letters = r.TakeWhile(char.IsAsciiLetter).Count();
        if (letters is 0 or > 3 || letters == r.Length || !r[letters..].All(char.IsAsciiDigit)) return false;
        if (bang < 0) return true;
        var sheet = Unquote(t[..bang]);
        return doc.Sheets.Any(s => string.Equals(s.Sheet.Name?.Value, sheet, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Rewrites references to a renamed sheet in the formula text.</summary>
    public static string Renamed(string formula, string oldName, string newName)
    {
        var bang = formula.LastIndexOf('!');
        if (bang < 0 || !string.Equals(Unquote(formula[..bang]), oldName, StringComparison.OrdinalIgnoreCase)) return formula;
        return Quote(newName) + formula[bang..];
    }

    /// <summary>A1-style references in a formula, outside string literals; the groups are the optional $ signs, the column and the row.</summary>
    static readonly System.Text.RegularExpressions.Regex Ref = new(@"(?<![\w.$])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(])|""[^""]*""", System.Text.RegularExpressions.RegexOptions.Compiled);

    /// <summary>The formula as Excel's fill handle writes it dRow rows down and dCol columns right: relative references move, $ pins one.</summary>
    public static string Shift(string formula, int dRow, int dCol)
    {
        if (dRow == 0 && dCol == 0) return formula;
        return Ref.Replace(formula, m =>
        {
            if (m.Groups[2].Length == 0) return m.Value; // a string literal
            var col = m.Groups[1].Length > 0 ? m.Groups[2].Value : XlsxCells.ColumnName(Math.Max(1, XlsxCells.ColumnIndex(m.Groups[2].Value.ToUpperInvariant()) + dCol));
            var row = m.Groups[3].Length > 0 ? m.Groups[4].Value : Math.Max(1, int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture) + dRow).ToString(CultureInfo.InvariantCulture);
            return m.Groups[1].Value + col + m.Groups[3].Value + row;
        });
    }

    static string Absolute(int col, int row) => "$" + XlsxCells.ColumnName(col) + "$" + row.ToString(CultureInfo.InvariantCulture);

    static string Quote(string sheet) =>
        sheet.Length > 0 && !char.IsDigit(sheet[0]) && sheet.All(ch => char.IsLetterOrDigit(ch) || ch == '_') ? sheet : "'" + sheet.Replace("'", "''") + "'";

    static string Unquote(string s) => s.Length >= 2 && s[0] == '\'' && s[^1] == '\'' ? s[1..^1].Replace("''", "'") : s;
}
