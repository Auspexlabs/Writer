using System.Globalization;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

/// <summary>Cell references, values and types.</summary>
static class XlsxCells
{
    static readonly string[] DateFormats = ["yyyy-MM-dd", "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm:ss.fff", "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ss.fff"];

    public static string ColumnName(int index)
    {
        var name = "";
        while (index > 0)
        {
            var remainder = (index - 1) % 26;
            name = (char)('A' + remainder) + name;
            index = (index - 1) / 26;
        }
        return name;
    }

    public static int ColumnIndex(ReadOnlySpan<char> letters)
    {
        var n = 0;
        foreach (var ch in letters) n = n * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
        return n;
    }

    public static (int Col, int Row) Parse(string reference)
    {
        var r = reference.Trim().Replace("$", "");
        var split = 0;
        while (split < r.Length && char.IsLetter(r[split])) split++;
        if (split is 0 or > 3 || split == r.Length || !int.TryParse(r.AsSpan(split), NumberStyles.None, CultureInfo.InvariantCulture, out var row) || row < 1)
            throw new WriterException(ErrorCode.Usage, $"'{reference}' is not a cell reference", "Use a column and row, e.g. B3.");
        return (ColumnIndex(r.AsSpan(0, split)), row);
    }

    public static (int Col1, int Row1, int Col2, int Row2) ParseRange(string reference)
    {
        var parts = reference.Split(':');
        if (parts.Length > 2) throw new WriterException(ErrorCode.Usage, $"'{reference}' is not a range", "Use A1:C10.");
        var (c1, r1) = Parse(parts[0]);
        var (c2, r2) = parts.Length == 2 ? Parse(parts[1]) : (c1, r1);
        return (Math.Min(c1, c2), Math.Min(r1, r2), Math.Max(c1, c2), Math.Max(r1, r2));
    }

    public static string Reference(int col, int row) => ColumnName(col) + row.ToString(CultureInfo.InvariantCulture);

    public static (int Col, int Row) Position(Cell cell) => Parse(cell.CellReference?.Value ?? "A1");

    public static Row? FindRow(SheetData data, int index) => data.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)index);

    public static Cell? FindCell(Row row, int col) => row.Elements<Cell>().FirstOrDefault(c => Position(c).Col == col);

    public static Row GetOrCreateRow(SheetData data, int index)
    {
        if (FindRow(data, index) is { } existing) return existing;
        var row = new Row { RowIndex = (uint)index };
        var next = data.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value > (uint)index);
        if (next is null) data.Append(row);
        else data.InsertBefore(row, next);
        return row;
    }

    public static Cell GetOrCreateCell(SheetData data, int col, int rowIndex)
    {
        var row = GetOrCreateRow(data, rowIndex);
        if (FindCell(row, col) is { } existing) return existing;
        var cell = new Cell { CellReference = Reference(col, rowIndex) };
        var next = row.Elements<Cell>().FirstOrDefault(c => Position(c).Col > col);
        if (next is null) row.Append(cell);
        else row.InsertBefore(cell, next);
        return cell;
    }

    /// <summary>The date an Excel serial stands for. 1900 system: 1 is 1900-01-01 and serials up to 60 sit a day off the calendar
    /// (Excel counts a 29 February 1900); 1904 system: 0 is 1904-01-01. A serial below 1 is a time of day.</summary>
    public static DateTime FromSerial(double serial, bool date1904) =>
        date1904 ? new DateTime(1904, 1, 1).AddMilliseconds(Math.Round(serial * 86400000)) : DateTime.FromOADate(serial >= 1 && serial < 61 ? serial + 1 : serial);

    public static double ToSerial(DateTime date, bool date1904)
    {
        if (date1904) return (date - new DateTime(1904, 1, 1)).TotalDays;
        var oa = date.ToOADate();
        return oa >= 1 && oa < 61 ? oa - 1 : oa;
    }

    public static string Display(XlsxDocument doc, Cell cell)
    {
        var raw = cell.CellValue?.Text ?? "";
        switch (cell.DataType?.InnerText)
        {
            case "s": return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i) ? doc.Strings.Get(i) : "";
            case "inlineStr": return cell.InlineString?.InnerText ?? "";
            case "str": return raw;
            case "b": return raw == "1" ? "true" : "false";
            case "e": return raw;
        }
        if (raw.Length > 0 && doc.Styles.IsDate(doc.Styles.FormatOf(cell)) && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial)
            && serial is > -1 and < 2958466)
        {
            var date = FromSerial(serial, doc.Date1904);
            return date.ToString(date.TimeOfDay == TimeSpan.Zero ? "yyyy-MM-dd" : date.Millisecond == 0 ? "yyyy-MM-dd HH:mm:ss" : "yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
        }
        return raw;
    }

    public static string? TypeOf(XlsxDocument doc, Cell cell)
    {
        var raw = cell.CellValue?.Text;
        switch (cell.DataType?.InnerText)
        {
            case "s" or "inlineStr" or "str": return "string";
            case "b": return "bool";
            case "e": return "error";
        }
        if (string.IsNullOrEmpty(raw)) return null;
        return doc.Styles.IsDate(doc.Styles.FormatOf(cell)) ? "date" : "number";
    }

    /// <summary>The cell's formula as a user would type it: a shared group's dependents (<c>&lt;f t="shared" si="n"/&gt;</c>) get the
    /// master's text shifted to their own position, as Excel shows it.</summary>
    public static string? Formula(XlsxSheet sheet, Cell cell)
    {
        if (cell.CellFormula is not { } f) return null;
        if (f.Text.Length > 0 || f.FormulaType?.InnerText != "shared") return f.Text;
        if (f.SharedIndex?.Value is not { } si || sheet.SharedMaster(si) is not { } master) return "";
        var (mc, mr) = Position(master);
        var (c, r) = Position(cell);
        return XlsxRefs.Shift(master.CellFormula!.Text, r - mr, c - mc);
    }

    /// <summary>Before a cell of a shared-formula group changes: every member gets its own formula (the master's, shifted), so the
    /// others keep computing and keep their cached values. Called wherever a formula is replaced, removed or moved.</summary>
    public static void Unshare(XlsxSheet sheet, Cell cell)
    {
        if (cell.CellFormula is not { } f || f.FormulaType?.InnerText != "shared" || f.SharedIndex?.Value is not { } si) return;
        var master = sheet.SharedMaster(si);
        var members = sheet.Data.Descendants<CellFormula>().Where(x => x.FormulaType?.InnerText == "shared" && x.SharedIndex?.Value == si && x.Parent is Cell).ToList();
        var text = master?.CellFormula?.Text;
        var (mc, mr) = master is null ? (0, 0) : Position(master);
        foreach (var member in members)
        {
            var owner = (Cell)member.Parent!;
            if (text is null) { member.Remove(); continue; } // a group without its master: the cells keep their values as constants
            var (c, r) = Position(owner);
            owner.CellFormula = new CellFormula(XlsxRefs.Shift(text, r - mr, c - mc));
        }
        sheet.ForgetShared();
    }

    /// <summary>Takes the formula off a cell that is getting a constant (or none). Excel's calc chain is dropped with it; Excel rebuilds it.</summary>
    static void DropFormula(XlsxDocument doc, XlsxSheet sheet, Cell cell)
    {
        if (cell.CellFormula is null) return;
        Unshare(sheet, cell);
        cell.RemoveAllChildren<CellFormula>();
        doc.DropCalcChain();
    }

    /// <summary>Stores a value, typing it from its shape: true/false, numbers and ISO dates are typed; everything else is text.
    /// A value the cell already holds, in the same type, leaves the cell untouched.</summary>
    public static void SetValue(XlsxDocument doc, XlsxSheet sheet, Cell cell, string value)
    {
        var trimmed = value.Trim();
        double number = 0;
        DateTime date = default;
        var isBool = trimmed.Equals("true", StringComparison.OrdinalIgnoreCase) || trimmed.Equals("false", StringComparison.OrdinalIgnoreCase);
        var isNumber = !isBool && double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number) && !trimmed.Contains(',');
        var isDate = !isBool && !isNumber && DateTime.TryParseExact(trimmed, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        var kind = trimmed.Length == 0 ? null : isBool ? "bool" : isNumber ? "number" : isDate ? "date" : "string";
        if (kind is not null && cell.CellFormula is null && TypeOf(doc, cell) == kind && Display(doc, cell) == value) return;
        DropFormula(doc, sheet, cell);
        if (kind is null)
        {
            cell.RemoveAllChildren<InlineString>();
            cell.CellValue = null;
            cell.DataType = null;
        }
        else if (isBool) SetBool(cell, trimmed.Equals("true", StringComparison.OrdinalIgnoreCase));
        else if (isNumber) SetNumber(cell, number);
        else if (isDate) SetDate(doc, cell, date);
        else SetString(doc, cell, value);
    }

    public static void SetTyped(XlsxDocument doc, XlsxSheet sheet, Cell cell, string type, string current)
    {
        switch (type)
        {
            case "string":
                DropFormula(doc, sheet, cell);
                SetString(doc, cell, current);
                break;
            case "number":
                if (!double.TryParse(current.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                    throw new WriterException(ErrorCode.Validation, $"'{current}' is not a number", "Write a numeric value first.");
                SetNumber(cell, number);
                break;
            case "bool":
                var t = current.Trim().ToLowerInvariant();
                if (t is not ("true" or "false" or "1" or "0"))
                    throw new WriterException(ErrorCode.Validation, $"'{current}' is not a boolean", "Use true or false.");
                SetBool(cell, t is "true" or "1");
                break;
            case "date":
                if (DateTime.TryParseExact(current.Trim(), DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) SetDate(doc, cell, date);
                else if (double.TryParse(current.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var serial) && serial is > -1 and < 2958466) SetDate(doc, cell, FromSerial(serial, doc.Date1904));
                else throw new WriterException(ErrorCode.Validation, $"'{current}' is not a date", "Use yyyy-mm-dd or yyyy-mm-dd hh:mm.");
                break;
        }
    }

    /// <summary>Writes text the way the cell already stores text: an inline string stays inline, a formula-string cell keeps its
    /// <c>t="str"</c>, anything else goes to the shared string table (reusing an entry with the same text).</summary>
    internal static void SetString(XlsxDocument doc, Cell cell, string value)
    {
        var kind = cell.DataType?.InnerText;
        if (kind is "s" or "str" or "inlineStr" && Display(doc, cell) == value) return;
        switch (kind)
        {
            case "inlineStr":
                cell.InlineString = new InlineString(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                cell.CellValue = null;
                return;
            case "str":
                cell.CellValue = new CellValue(value);
                return;
        }
        cell.RemoveAllChildren<InlineString>();
        cell.DataType = CellValues.SharedString;
        cell.CellValue = new CellValue(doc.Strings.Add(value).ToString(CultureInfo.InvariantCulture));
    }

    internal static void SetNumber(Cell cell, double number)
    {
        cell.RemoveAllChildren<InlineString>();
        if (cell.DataType?.InnerText != "n") cell.DataType = null; // an explicit t="n" is kept as written
        cell.CellValue = new CellValue(number.ToString("R", CultureInfo.InvariantCulture));
    }

    internal static void SetBool(Cell cell, bool value)
    {
        cell.RemoveAllChildren<InlineString>();
        cell.DataType = CellValues.Boolean;
        cell.CellValue = new CellValue(value ? "1" : "0");
    }

    internal static void SetDate(XlsxDocument doc, Cell cell, DateTime date)
    {
        cell.RemoveAllChildren<InlineString>();
        cell.DataType = null;
        cell.CellValue = new CellValue(ToSerial(date, doc.Date1904).ToString("R", CultureInfo.InvariantCulture));
        if (!doc.Styles.IsDate(doc.Styles.FormatOf(cell)))
            doc.Styles.Apply(cell, xf => { xf.NumberFormatId = doc.Styles.NumberFormatId(date.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss"); xf.ApplyNumberFormat = true; });
    }

    /// <summary>Writes a formula. The same text as the cell already computes changes nothing (a shared group stays shared, the cached
    /// result stays). A new text first gives every cell of a shared group its own formula; an array formula keeps its range. The
    /// cell's last result is kept for readers that do not calculate, and Excel is asked to recalculate on load.</summary>
    public static void SetFormula(XlsxDocument doc, XlsxSheet sheet, Cell cell, string formula)
    {
        var text = formula.Trim().TrimStart('=');
        if (text.Length == 0)
        {
            DropFormula(doc, sheet, cell);
            return;
        }
        if (text == Formula(sheet, cell)) return;
        Unshare(sheet, cell);
        var old = cell.CellFormula;
        if (old is null || cell.DataType?.InnerText is "s" or "inlineStr")
        {
            // a constant becoming a formula: its text was never this formula's result
            cell.RemoveAllChildren<InlineString>();
            cell.CellValue = null;
            cell.DataType = null;
        }
        cell.CellFormula = old?.FormulaType?.InnerText == "array"
            ? new CellFormula(text) { FormulaType = CellFormulaValues.Array, Reference = old.Reference?.Value }
            : new CellFormula(text);
        doc.DropCalcChain();
        doc.RecalculateOnLoad();
    }

    /// <summary>Minimal CSV: commas (or the given delimiter), quotes, doubled quotes, CRLF or LF lines.</summary>
    public static List<List<string>> ParseCsv(string text, char delimiter = ',')
    {
        var rows = new List<List<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        var quoted = false;
        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (quoted)
            {
                if (ch == '"' && i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                else if (ch == '"') quoted = false;
                else field.Append(ch);
                continue;
            }
            switch (ch)
            {
                case '"': quoted = true; break;
                case var d when d == delimiter: row.Add(field.ToString()); field.Clear(); break;
                case '\r': break;
                case '\n': row.Add(field.ToString()); field.Clear(); rows.Add(row); row = []; break;
                default: field.Append(ch); break;
            }
        }
        if (field.Length > 0 || row.Count > 0)
        {
            row.Add(field.ToString());
            rows.Add(row);
        }
        return rows;
    }
}

/// <summary>Cell styles: reading fonts, fills, borders, alignment and number formats, and deriving new formats for edits.
/// Every write goes through the style tables with de-duplication; reads never create a styles part.</summary>
sealed class XlsxStyles(WorkbookPart workbook)
{
    public const string MinimalXml =
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        + "<fonts count=\"1\"><font><sz val=\"11\"/><color rgb=\"FF000000\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>"
        + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
        + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
        + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        + "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
        + "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        + "</styleSheet>";

    /// <summary>Formatting properties shared by the cell and range kinds; all written through <see cref="Set"/>.</summary>
    public static readonly string[] Props =
        ["bold", "italic", "underline", "strike", "size", "font", "color", "fill", "format", "align", "valign", "wrap", "indent", "rotate", "border", "borders", "borderColor"];

    static readonly string[] Sides = ["top", "right", "bottom", "left"];

    static readonly Prop BorderProp = Registry.Find("cell")!.Props.First(p => p.Name == "border");

    static readonly Dictionary<uint, string> Builtin = new()
    {
        [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00", [9] = "0%", [10] = "0.00%", [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??",
        [14] = "m/d/yyyy", [15] = "d-mmm-yy", [16] = "d-mmm", [17] = "mmm-yy", [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM", [20] = "h:mm", [21] = "h:mm:ss",
        [22] = "m/d/yyyy h:mm", [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)", [39] = "#,##0.00;(#,##0.00)", [40] = "#,##0.00;[Red](#,##0.00)",
        [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@",
    };

    XlsxColors? _colors;

    /// <summary>Theme, indexed and rgb colours as RRGGBB (and back to the file's theme references on writes).</summary>
    public XlsxColors Colors => _colors ??= new XlsxColors(workbook);

    /// <summary>What a fill shows: a solid pattern's colour, or a gradient's first stop. Other patterns are left out.</summary>
    public string? FillColor(Fill? fill) =>
        fill?.PatternFill is { } pattern ? pattern.PatternType?.InnerText == "solid" ? Colors.Resolve(pattern.ForegroundColor) ?? Colors.Resolve(pattern.BackgroundColor) : null
        : Colors.Resolve(fill?.GradientFill?.GetFirstChild<GradientStop>()?.Color);

    /// <summary>The workbook's default font (the first one, as Normal uses): name and size, as cells without their own show them.</summary>
    public (string? Name, double? Size) Normal => Existing?.Fonts?.GetFirstChild<Font>() is { } f ? (f.FontName?.Val?.Value, f.FontSize?.Val?.Value) : (null, null);

    /// <summary>The stylesheet for reading; null when the workbook has none. Reads never create parts.</summary>
    Stylesheet? Existing => workbook.WorkbookStylesPart?.Stylesheet;

    /// <summary>The stylesheet for writing, created on demand.</summary>
    Stylesheet Sheet
    {
        get
        {
            var part = workbook.WorkbookStylesPart ?? workbook.AddNewPart<WorkbookStylesPart>();
            return part.Stylesheet ??= new Stylesheet(MinimalXml);
        }
    }

    public CellFormat? FormatOf(Cell cell)
    {
        var index = (int)(cell.StyleIndex?.Value ?? 0);
        return Existing?.CellFormats?.Elements<CellFormat>().ElementAtOrDefault(index);
    }

    public Font? FontOf(CellFormat? format) =>
        format?.FontId?.Value is { } id ? Existing?.Fonts?.Elements<Font>().ElementAtOrDefault((int)id) : null;

    public Fill? FillOf(CellFormat? format) =>
        format?.FillId?.Value is { } id ? Existing?.Fills?.Elements<Fill>().ElementAtOrDefault((int)id) : null;

    public Border? BorderOf(CellFormat? format) =>
        format?.BorderId?.Value is { } id ? Existing?.Borders?.Elements<Border>().ElementAtOrDefault((int)id) : null;

    public string? NumberFormatCode(CellFormat? format)
    {
        var id = format?.NumberFormatId?.Value ?? 0;
        if (id == 0) return null;
        if (Builtin.TryGetValue(id, out var code)) return code;
        return Existing?.NumberingFormats?.Elements<NumberingFormat>().FirstOrDefault(n => n.NumberFormatId?.Value == id)?.FormatCode?.Value;
    }

    public bool IsDate(CellFormat? format)
    {
        var id = format?.NumberFormatId?.Value ?? 0;
        if (id is >= 14 and <= 22 or >= 45 and <= 47) return true;
        if (id < 164) return false;
        var code = NumberFormatCode(format) ?? "";
        var stripped = System.Text.RegularExpressions.Regex.Replace(code, "\\[[^\\]]*\\]|\"[^\"]*\"|[\\\\_*].", ""); // brackets, quoted text, escaped, padding and fill characters
        return stripped.IndexOfAny(['y', 'd', 'h', 'Y', 'D', 'H']) >= 0 || (stripped.Contains('m') && !stripped.Contains('0') && !stripped.Contains('#'));
    }

    /// <summary>Adds the cell's formatting to props: set attributes only, and size/font only when they differ from the workbook's default font.</summary>
    public void Read(Cell cell, Dictionary<string, string> props)
    {
        var format = FormatOf(cell);
        var font = FontOf(format);
        var normal = Existing?.Fonts?.GetFirstChild<Font>();
        if (On(font?.Bold)) props["bold"] = "true";
        if (On(font?.Italic)) props["italic"] = "true";
        if (font?.Underline is { } u && u.Val?.InnerText != "none") props["underline"] = "true";
        if (On(font?.Strike)) props["strike"] = "true";
        if (font?.FontSize?.Val?.Value is { } size && size != (normal?.FontSize?.Val?.Value ?? 11)) props["size"] = size.ToString("0.##", CultureInfo.InvariantCulture);
        if (font?.FontName?.Val?.Value is { } name && name != normal?.FontName?.Val?.Value) props["font"] = name;
        if (Colors.Resolve(font?.Color) is { } color && color != "000000") props["color"] = color;
        if (FillColor(FillOf(format)) is { } fill) props["fill"] = fill;
        if (NumberFormatCode(format) is { } code) props["format"] = code;
        if (format?.Alignment is { } al)
        {
            if (al.Horizontal?.InnerText is ("left" or "center" or "right" or "justify") and { } h) props["align"] = h;
            if (al.Vertical?.InnerText is ("top" or "center" or "bottom") and { } v) props["valign"] = v == "center" ? "middle" : v;
            if (al.WrapText?.Value == true) props["wrap"] = "true";
            if (al.Indent?.Value is > 0 and var indent) props["indent"] = indent.ToString(CultureInfo.InvariantCulture);
            // Excel: 0–90 counter-clockwise, 91–180 clockwise (90 + degrees), 255 stacked; the engine speaks -90..90
            if (al.TextRotation?.Value is > 0 and <= 180 and var rot) props["rotate"] = (rot <= 90 ? (int)rot : 90 - (int)rot).ToString(CultureInfo.InvariantCulture);
        }
        var styled = SidesOf(BorderOf(format)).Where(s => s.Style != "none").ToList();
        if (styled.Count == 0) return;
        var main = styled.GroupBy(s => s.Style).OrderByDescending(g => g.Count()).First().Key;
        props["border"] = main;
        if (styled.Count < 4 || styled.Any(s => s.Style != main))
            props["borders"] = NodeJson.Compact(w =>
            {
                w.WriteStartObject();
                foreach (var s in styled) w.WriteString(s.Side, s.Style);
                w.WriteEndObject();
            });
        if (styled.Select(s => s.Color).Distinct().ToList() is [{ } one]) props["borderColor"] = one;
    }

    /// <summary>An edit's props split into the cell's style ones and the rest, each in the order given.</summary>
    public static (List<KeyValuePair<string, string>> Style, List<KeyValuePair<string, string>> Other) Split(IEnumerable<KeyValuePair<string, string>> props)
    {
        var list = props.ToList();
        return (list.Where(p => Props.Contains(p.Key)).ToList(), list.Where(p => !Props.Contains(p.Key)).ToList());
    }

    /// <summary>Writes formatting properties (see <see cref="Props"/>) as one new format for the cell, keeping its other style attributes:
    /// the font, border and format each go through their tables once, whatever the number of props. borderColor is the colour
    /// pending for border writes made in the same edit.</summary>
    public void Set(Cell cell, IReadOnlyList<KeyValuePair<string, string>> changes, string? borderColor = null)
    {
        var xf = (CellFormat)(FormatOf(cell) ?? new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U }).CloneNode(true);
        Font? font = null;
        Border? border = null;
        foreach (var (name, value) in changes)
            switch (name)
            {
                case "bold" or "italic" or "underline" or "strike" or "size" or "font" or "color":
                    font ??= (Font)(FontOf(xf) ?? Sheet.Fonts?.GetFirstChild<Font>() ?? new Font()).CloneNode(true);
                    switch (name)
                    {
                        case "bold": font.Bold = value == "true" ? new Bold() : null; break;
                        case "italic": font.Italic = value == "true" ? new Italic() : null; break;
                        case "underline": font.Underline = value == "true" ? new Underline() : null; break;
                        case "strike": font.Strike = value == "true" ? new Strike() : null; break;
                        case "size": font.FontSize = new FontSize { Val = double.Parse(value, CultureInfo.InvariantCulture) }; break;
                        case "font": font.FontName = new FontName { Val = value }; break;
                        default: font.Color = value == "none" ? null : Colors.Like<Color>(value); break;
                    }
                    break;
                case "fill":
                    xf.FillId = value == "none"
                        ? 0u
                        : Index(Sheet.Fills ??= new Fills(), new Fill(new PatternFill(Colors.Like<ForegroundColor>(value), new BackgroundColor { Indexed = 64U }) { PatternType = PatternValues.Solid }), n => Sheet.Fills!.Count = n);
                    xf.ApplyFill = value != "none" ? true : null;
                    break;
                case "format":
                    xf.NumberFormatId = NumberFormatId(value.Trim());
                    xf.ApplyNumberFormat = true;
                    break;
                case "align" or "valign" or "wrap" or "indent" or "rotate":
                    var al = xf.Alignment ??= new Alignment();
                    switch (name)
                    {
                        case "align": al.Horizontal = new HorizontalAlignmentValues(value); break;
                        case "valign": al.Vertical = new VerticalAlignmentValues(value == "middle" ? "center" : value); break;
                        case "wrap": al.WrapText = value == "true" ? true : null; break;
                        case "rotate": var deg = int.Parse(value, CultureInfo.InvariantCulture); al.TextRotation = deg == 0 ? null : (uint)(deg > 0 ? deg : 90 - deg); break;
                        default: al.Indent = value == "0" ? null : uint.Parse(value, CultureInfo.InvariantCulture); break;
                    }
                    xf.ApplyAlignment = true;
                    break;
                case "border": border = SetSides(border ?? BorderClone(xf), Sides.Select(s => new KeyValuePair<string, string>(s, value)), borderColor); break;
                case "borders": border = SetSides(border ?? BorderClone(xf), ParseSides(value), borderColor); break;
                case "borderColor":
                    border ??= BorderClone(xf);
                    border = SetSides(border, SidesOf(border).Where(s => s.Style != "none").Select(s => new KeyValuePair<string, string>(s.Side, s.Style)), value);
                    break;
            }
        if (xf.Alignment is { } alignment)
        {
            // Excel's defaults (general, bottom) are no alignment at all: a cell set back to them is the cell it was
            if (alignment.Horizontal?.InnerText == "general") alignment.Horizontal = null;
            if (alignment.Vertical?.InnerText == "bottom") alignment.Vertical = null;
            if (!alignment.HasAttributes && !alignment.HasChildren) (xf.Alignment, xf.ApplyAlignment) = (null, null);
        }
        if (font is not null)
        {
            xf.FontId = Index(Sheet.Fonts ??= new Fonts(), font, n => Sheet.Fonts!.Count = n);
            xf.ApplyFont = true;
        }
        if (border is not null)
        {
            xf.BorderId = Index(Sheet.Borders ??= new Borders(), border, n => Sheet.Borders!.Count = n);
            xf.ApplyBorder = true;
        }
        cell.StyleIndex = Index(Sheet.CellFormats ??= new CellFormats(), xf, n => Sheet.CellFormats!.Count = n);
        Touched = true;
    }

    Border BorderClone(CellFormat xf) => (Border)(BorderOf(xf) ?? Sheet.Borders?.GetFirstChild<Border>() ?? new Border()).CloneNode(true);

    /// <summary>Sets the style of the given sides ("none" clears one) and, when a color is given, their color.</summary>
    Border SetSides(Border border, IEnumerable<KeyValuePair<string, string>> sides, string? color)
    {
        foreach (var (side, style) in sides)
        {
            var element = side switch
            {
                "top" => border.TopBorder ??= new TopBorder(),
                "right" => border.RightBorder ??= new RightBorder(),
                "bottom" => border.BottomBorder ??= new BottomBorder(),
                _ => (BorderPropertiesType)(border.LeftBorder ??= new LeftBorder()),
            };
            element.Style = style == "none" ? null : new BorderStyleValues(style);
            if (style == "none" || color == "none") element.Color = null;
            else if (color is not null) element.Color = Colors.Like<Color>(color);
        }
        return border;
    }

    static List<KeyValuePair<string, string>> ParseSides(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Object)
            throw new WriterException(ErrorCode.Validation, "borders must be a JSON object of sides", "Example: {\"bottom\":\"medium\",\"top\":\"thin\"}");
        var sides = new List<KeyValuePair<string, string>>();
        foreach (var p in parsed.RootElement.EnumerateObject())
        {
            var side = p.Name.Trim().ToLowerInvariant();
            if (!Sides.Contains(side))
                throw new WriterException(ErrorCode.Validation, $"borders: '{p.Name}' is not a side", "Sides: top, right, bottom, left.");
            sides.Add(new(side, Registry.NormalizeValue(BorderProp, p.Value.ToString())));
        }
        return sides;
    }

    (string Side, string Style, string? Color)[] SidesOf(Border? border) =>
    [
        Side("top", border?.TopBorder), Side("right", border?.RightBorder), Side("bottom", border?.BottomBorder), Side("left", border?.LeftBorder),
    ];

    (string Side, string Style, string? Color) Side(string side, BorderPropertiesType? element) =>
        (side, element?.Style?.InnerText ?? "none", Colors.Resolve(element?.Color));

    /// <summary>Gives the cell the format that is its current one with the change applied, reusing an identical entry when one exists.</summary>
    public void Apply(Cell cell, Action<CellFormat> change)
    {
        var source = FormatOf(cell) ?? new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };
        var xf = (CellFormat)source.CloneNode(true);
        change(xf);
        cell.StyleIndex = Index(Sheet.CellFormats ??= new CellFormats(), xf, n => Sheet.CellFormats!.Count = n);
        Touched = true;
    }

    /// <summary>A cell's format was written since the workbook was opened: <see cref="Trim"/> has work at save.</summary>
    public bool Touched { get; private set; }

    /// <summary>At save, after style writes: the trailing formats no cell, row or column uses any more go, and then the trailing fonts,
    /// fills and borders no format uses. Earlier entries keep their places (every index after them would move), so the tables end
    /// where the workbook's styles end instead of growing with each edit that left an entry behind. Worksheets not loaded are read
    /// as streams, not parsed.</summary>
    public void Trim(IEnumerable<WorksheetPart> sheets)
    {
        if (!Touched || Existing?.CellFormats is not { } xfs) return;
        var used = new HashSet<uint>();
        foreach (var part in sheets)
        {
            if (part.IsRootElementLoaded)
            {
                foreach (var e in part.Worksheet!.Descendants())
                    if (e is Cell { StyleIndex.Value: var s }) used.Add(s);
                    else if (e is Row { StyleIndex.Value: var rs }) used.Add(rs);
                    else if (e is Column { Style.Value: var cs }) used.Add(cs);
                continue;
            }
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = System.Xml.XmlReader.Create(stream);
            while (reader.Read())
                if (reader.NodeType == System.Xml.XmlNodeType.Element && (reader.LocalName is "c" or "row" ? reader.GetAttribute("s") : reader.LocalName == "col" ? reader.GetAttribute("style") : null) is { } v
                    && uint.TryParse(v, out var index)) used.Add(index);
        }
        Shrink(xfs, 1, i => used.Contains(i), n => xfs.Count = n);
        var formats = xfs.Elements<CellFormat>().Concat(Existing.CellStyleFormats?.Elements<CellFormat>() ?? []).ToList();
        if (Existing.Fonts is { } fonts) Shrink(fonts, 1, i => formats.Any(f => f.FontId?.Value == i), n => fonts.Count = n);
        if (Existing.Fills is { } fills) Shrink(fills, 2, i => formats.Any(f => f.FillId?.Value == i), n => fills.Count = n);
        if (Existing.Borders is { } borders) Shrink(borders, 1, i => formats.Any(f => f.BorderId?.Value == i), n => borders.Count = n);
    }

    static void Shrink(OpenXmlCompositeElement table, int keep, Func<uint, bool> used, Action<uint> setCount)
    {
        var entries = table.ChildElements.ToList();
        var count = entries.Count;
        while (count > keep && !used((uint)count - 1)) entries[--count].Remove();
        if (count < entries.Count) setCount((uint)count);
    }

    /// <summary>The index of an entry identical to the given one in a style table, appending it when there is none.</summary>
    static uint Index(OpenXmlCompositeElement table, OpenXmlElement entry, Action<uint> setCount)
    {
        var wanted = Signature(entry);
        var i = 0u;
        foreach (var existing in table.ChildElements)
        {
            if (Signature(existing) == wanted) return i;
            i++;
        }
        table.Append(entry);
        setCount(i + 1);
        return i;
    }

    /// <summary>Names, attributes and text of an element tree, independent of namespace prefixes.</summary>
    static string Signature(OpenXmlElement element)
    {
        var sb = new StringBuilder();
        Write(element);
        return sb.ToString();

        void Write(OpenXmlElement e)
        {
            sb.Append('<').Append(e.LocalName);
            // applyFont="1" and the like say what a format's own ids already say: an entry with or without them is the same entry
            foreach (var a in e.GetAttributes().Where(a => !(a.LocalName.StartsWith("apply", StringComparison.Ordinal) && a.Value is "1" or "true")).OrderBy(a => a.LocalName, StringComparer.Ordinal))
                sb.Append(' ').Append(a.LocalName).Append('=').Append(a.Value);
            sb.Append('>');
            if (e is OpenXmlLeafTextElement leaf) sb.Append(leaf.Text);
            foreach (var child in e.ChildElements) Write(child);
            sb.Append("</>");
        }
    }

    /// <summary>What a conditional format's dxf changes: fill, color, bold, italic.</summary>
    public IEnumerable<(string Name, string Value)> Dxf(uint id)
    {
        var dxf = Existing?.DifferentialFormats?.Elements<DifferentialFormat>().ElementAtOrDefault((int)id);
        if (dxf is null) yield break;
        var pattern = dxf.Fill?.PatternFill;
        if ((Colors.Resolve(pattern?.BackgroundColor) ?? Colors.Resolve(pattern?.ForegroundColor)) is { } fill) yield return ("fill", fill);
        if (Colors.Resolve(dxf.Font?.Color) is { } color) yield return ("color", color);
        if (On(dxf.Font?.Bold)) yield return ("bold", "true");
        if (On(dxf.Font?.Italic)) yield return ("italic", "true");
    }

    /// <summary>The dxf index for a conditional format's look, reusing an identical entry when one exists.</summary>
    public uint DxfId(string? fill, string? color, bool bold, bool italic)
    {
        var dxf = new DifferentialFormat();
        if (color is not null || bold || italic)
        {
            var font = new Font();
            if (bold) font.Append(new Bold());
            if (italic) font.Append(new Italic());
            if (color is not null) font.Append(Colors.Like<Color>(color));
            dxf.Append(font);
        }
        if (fill is not null) dxf.Append(new Fill(new PatternFill(Colors.Like<BackgroundColor>(fill)) { PatternType = PatternValues.Solid }));
        return Index(Sheet.DifferentialFormats ??= new DifferentialFormats(), dxf, n => Sheet.DifferentialFormats!.Count = n);
    }

    public uint NumberFormatId(string code)
    {
        if (code.Length == 0 || code.Equals("General", StringComparison.OrdinalIgnoreCase)) return 0;
        foreach (var (id, builtin) in Builtin)
            if (builtin == code) return id;
        var formats = Sheet.NumberingFormats;
        if (formats is null)
        {
            formats = new NumberingFormats();
            Sheet.PrependChild(formats);
        }
        if (formats.Elements<NumberingFormat>().FirstOrDefault(n => n.FormatCode?.Value == code)?.NumberFormatId?.Value is { } existing) return existing;
        var next = Math.Max(164u, formats.Elements<NumberingFormat>().Select(n => n.NumberFormatId?.Value ?? 0).DefaultIfEmpty(0u).Max() + 1);
        formats.Append(new NumberingFormat { NumberFormatId = next, FormatCode = code });
        formats.Count = (uint)formats.Elements<NumberingFormat>().Count();
        return next;
    }

    static bool On(BooleanPropertyType? flag) => flag is not null && (flag.Val is null || flag.Val.Value);

    public static string? Rgb(string? argb) => argb is { Length: 8 } ? argb[2..].ToUpperInvariant() : argb is { Length: 6 } ? argb.ToUpperInvariant() : null;
}
