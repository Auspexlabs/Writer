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
    static readonly string[] DateFormats = ["yyyy-MM-dd", "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss"];

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
        if (raw.Length > 0 && doc.Styles.IsDate(doc.Styles.FormatOf(cell)) && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
        {
            var date = DateTime.FromOADate(serial);
            return date.TimeOfDay == TimeSpan.Zero ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
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

    /// <summary>Stores a value, typing it from its shape: true/false, numbers and ISO dates are typed; everything else is text.</summary>
    public static void SetValue(XlsxDocument doc, Cell cell, string value)
    {
        cell.RemoveAllChildren<CellFormula>();
        cell.RemoveAllChildren<InlineString>();
        var trimmed = value.Trim();
        if (trimmed.Length == 0)
        {
            cell.CellValue = null;
            cell.DataType = null;
            return;
        }
        if (trimmed.Equals("true", StringComparison.OrdinalIgnoreCase) || trimmed.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            SetBool(cell, trimmed.Equals("true", StringComparison.OrdinalIgnoreCase));
            return;
        }
        if (double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var number) && !trimmed.Contains(','))
        {
            SetNumber(cell, number);
            return;
        }
        if (DateTime.TryParseExact(trimmed, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            SetDate(doc, cell, date);
            return;
        }
        SetString(doc, cell, value);
    }

    public static void SetTyped(XlsxDocument doc, Cell cell, string type, string current)
    {
        switch (type)
        {
            case "string":
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
                else if (double.TryParse(current.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var serial)) SetDate(doc, cell, DateTime.FromOADate(serial));
                else throw new WriterException(ErrorCode.Validation, $"'{current}' is not a date", "Use yyyy-mm-dd or yyyy-mm-dd hh:mm.");
                break;
        }
    }

    static void SetString(XlsxDocument doc, Cell cell, string value)
    {
        cell.RemoveAllChildren<CellFormula>();
        cell.DataType = CellValues.SharedString;
        cell.CellValue = new CellValue(doc.Strings.Add(value).ToString(CultureInfo.InvariantCulture));
    }

    static void SetNumber(Cell cell, double number)
    {
        cell.DataType = null;
        cell.CellValue = new CellValue(number.ToString("R", CultureInfo.InvariantCulture));
    }

    static void SetBool(Cell cell, bool value)
    {
        cell.DataType = CellValues.Boolean;
        cell.CellValue = new CellValue(value ? "1" : "0");
    }

    static void SetDate(XlsxDocument doc, Cell cell, DateTime date)
    {
        cell.DataType = null;
        cell.CellValue = new CellValue(date.ToOADate().ToString("R", CultureInfo.InvariantCulture));
        if (!doc.Styles.IsDate(doc.Styles.FormatOf(cell)))
            doc.Styles.Apply(cell, xf => { xf.NumberFormatId = doc.Styles.NumberFormatId(date.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss"); xf.ApplyNumberFormat = true; });
    }

    public static void SetFormula(XlsxDocument doc, Cell cell, string formula)
    {
        var text = formula.Trim().TrimStart('=');
        cell.RemoveAllChildren<CellFormula>();
        cell.RemoveAllChildren<InlineString>();
        if (text.Length == 0) return;
        cell.CellValue = null;
        cell.DataType = null;
        cell.CellFormula = new CellFormula(text);
        doc.RecalculateOnLoad();
    }

    /// <summary>Minimal CSV: commas, quotes, doubled quotes, CRLF or LF lines.</summary>
    public static List<List<string>> ParseCsv(string text)
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
                case ',': row.Add(field.ToString()); field.Clear(); break;
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
        ["bold", "italic", "underline", "strike", "size", "font", "color", "fill", "format", "align", "valign", "wrap", "indent", "border", "borders", "borderColor"];

    static readonly string[] Sides = ["top", "right", "bottom", "left"];

    static readonly Prop BorderProp = Registry.Find("cell")!.Props.First(p => p.Name == "border");

    static readonly Dictionary<uint, string> Builtin = new()
    {
        [1] = "0", [2] = "0.00", [3] = "#,##0", [4] = "#,##0.00", [9] = "0%", [10] = "0.00%", [11] = "0.00E+00", [12] = "# ?/?", [13] = "# ??/??",
        [14] = "m/d/yyyy", [15] = "d-mmm-yy", [16] = "d-mmm", [17] = "mmm-yy", [18] = "h:mm AM/PM", [19] = "h:mm:ss AM/PM", [20] = "h:mm", [21] = "h:mm:ss",
        [22] = "m/d/yyyy h:mm", [37] = "#,##0 ;(#,##0)", [38] = "#,##0 ;[Red](#,##0)", [39] = "#,##0.00;(#,##0.00)", [40] = "#,##0.00;[Red](#,##0.00)",
        [45] = "mm:ss", [46] = "[h]:mm:ss", [47] = "mmss.0", [48] = "##0.0E+0", [49] = "@",
    };

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
        var stripped = System.Text.RegularExpressions.Regex.Replace(code, "\\[[^\\]]*\\]|\"[^\"]*\"", "");
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
        if (Rgb(font?.Color?.Rgb?.Value) is { } color && color != "000000") props["color"] = color;
        if (FillOf(format)?.PatternFill is { } pattern && pattern.PatternType?.InnerText == "solid" && Rgb(pattern.ForegroundColor?.Rgb?.Value) is { } fill)
            props["fill"] = fill;
        if (NumberFormatCode(format) is { } code) props["format"] = code;
        if (format?.Alignment is { } al)
        {
            if (al.Horizontal?.InnerText is ("left" or "center" or "right" or "justify") and { } h) props["align"] = h;
            if (al.Vertical?.InnerText is ("top" or "center" or "bottom") and { } v) props["valign"] = v == "center" ? "middle" : v;
            if (al.WrapText?.Value == true) props["wrap"] = "true";
            if (al.Indent?.Value is > 0 and var indent) props["indent"] = indent.ToString(CultureInfo.InvariantCulture);
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

    /// <summary>Writes one formatting property (see <see cref="Props"/>), keeping the cell's other style attributes.
    /// borderColor is the color pending for border writes made in the same edit.</summary>
    public void Set(Cell cell, string name, string value, string? borderColor = null)
    {
        switch (name)
        {
            case "bold" or "italic" or "underline" or "strike" or "size" or "font" or "color":
                var font = (Font)(FontOf(FormatOf(cell)) ?? Sheet.Fonts?.GetFirstChild<Font>() ?? new Font()).CloneNode(true);
                switch (name)
                {
                    case "bold": font.Bold = value == "true" ? new Bold() : null; break;
                    case "italic": font.Italic = value == "true" ? new Italic() : null; break;
                    case "underline": font.Underline = value == "true" ? new Underline() : null; break;
                    case "strike": font.Strike = value == "true" ? new Strike() : null; break;
                    case "size": font.FontSize = new FontSize { Val = double.Parse(value, CultureInfo.InvariantCulture) }; break;
                    case "font": font.FontName = new FontName { Val = value }; break;
                    default: font.Color = value == "none" ? null : new Color { Rgb = "FF" + value }; break;
                }
                var fontId = Index(Sheet.Fonts ??= new Fonts(), font, n => Sheet.Fonts!.Count = n);
                Apply(cell, xf => { xf.FontId = fontId; xf.ApplyFont = true; });
                break;
            case "fill":
                var fillId = value == "none"
                    ? 0u
                    : Index(Sheet.Fills ??= new Fills(), new Fill(new PatternFill(new ForegroundColor { Rgb = "FF" + value }, new BackgroundColor { Indexed = 64U }) { PatternType = PatternValues.Solid }), n => Sheet.Fills!.Count = n);
                Apply(cell, xf => { xf.FillId = fillId; xf.ApplyFill = value != "none"; });
                break;
            case "format":
                var formatId = NumberFormatId(value.Trim());
                Apply(cell, xf => { xf.NumberFormatId = formatId; xf.ApplyNumberFormat = true; });
                break;
            case "align" or "valign" or "wrap" or "indent":
                Apply(cell, xf =>
                {
                    var al = xf.Alignment ??= new Alignment();
                    switch (name)
                    {
                        case "align": al.Horizontal = new HorizontalAlignmentValues(value); break;
                        case "valign": al.Vertical = new VerticalAlignmentValues(value == "middle" ? "center" : value); break;
                        case "wrap": al.WrapText = value == "true" ? true : null; break;
                        default: al.Indent = value == "0" ? null : uint.Parse(value, CultureInfo.InvariantCulture); break;
                    }
                    xf.ApplyAlignment = true;
                });
                break;
            case "border": SetBorder(cell, Sides.Select(s => new KeyValuePair<string, string>(s, value)), borderColor); break;
            case "borders": SetBorder(cell, ParseSides(value), borderColor); break;
            case "borderColor":
                SetBorder(cell, SidesOf(BorderOf(FormatOf(cell))).Where(s => s.Style != "none").Select(s => new KeyValuePair<string, string>(s.Side, s.Style)), value);
                break;
        }
    }

    /// <summary>Sets the style of the given sides ("none" clears one) and, when a color is given, their color.</summary>
    void SetBorder(Cell cell, IEnumerable<KeyValuePair<string, string>> sides, string? color)
    {
        var border = (Border)(BorderOf(FormatOf(cell)) ?? Sheet.Borders?.GetFirstChild<Border>() ?? new Border()).CloneNode(true);
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
            else if (color is not null) element.Color = new Color { Rgb = "FF" + color };
        }
        var id = Index(Sheet.Borders ??= new Borders(), border, n => Sheet.Borders!.Count = n);
        Apply(cell, xf => { xf.BorderId = id; xf.ApplyBorder = true; });
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

    static (string Side, string Style, string? Color)[] SidesOf(Border? border) =>
    [
        Side("top", border?.TopBorder), Side("right", border?.RightBorder), Side("bottom", border?.BottomBorder), Side("left", border?.LeftBorder),
    ];

    static (string Side, string Style, string? Color) Side(string side, BorderPropertiesType? element) =>
        (side, element?.Style?.InnerText ?? "none", Rgb(element?.Color?.Rgb?.Value));

    /// <summary>Gives the cell the format that is its current one with the change applied, reusing an identical entry when one exists.</summary>
    public void Apply(Cell cell, Action<CellFormat> change)
    {
        var source = FormatOf(cell) ?? new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };
        var xf = (CellFormat)source.CloneNode(true);
        change(xf);
        cell.StyleIndex = Index(Sheet.CellFormats ??= new CellFormats(), xf, n => Sheet.CellFormats!.Count = n);
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
            foreach (var a in e.GetAttributes().OrderBy(a => a.LocalName, StringComparer.Ordinal)) sb.Append(' ').Append(a.LocalName).Append('=').Append(a.Value);
            sb.Append('>');
            if (e is OpenXmlLeafTextElement leaf) sb.Append(leaf.Text);
            foreach (var child in e.ChildElements) Write(child);
            sb.Append("</>");
        }
    }

    public uint NumberFormatId(string code)
    {
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
