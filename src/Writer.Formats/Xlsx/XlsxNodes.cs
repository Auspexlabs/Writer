using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
using Writer.Formats.Common;

namespace Writer.Formats.Xlsx;

sealed class XlsxRoot(XlsxDocument doc) : Node
{
    public override string Kind => "document";
    public override string Format => "xlsx";
    public override object Anchor => doc.Package;

    protected override IEnumerable<Node> ProjectChildren() => doc.Sheets.Select(s => (Node)new XlsxSheet(doc, s.Sheet, s.Part));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["format"] = "xlsx", ["sheets"] = doc.Sheets.Count.ToString(CultureInfo.InvariantCulture) };
        if (doc.Package.PackageProperties.Title is { Length: > 0 } title) props["title"] = title;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        if (name == "title") doc.Package.PackageProperties.Title = value.Length > 0 ? value : null;
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var part = doc.AddSheet(props.GetValueOrDefault("name"), index);
        var entry = doc.Sheets.First(s => ReferenceEquals(s.Part, part));
        return new XlsxSheet(doc, entry.Sheet, entry.Part);
    }
}

sealed class XlsxSheet(XlsxDocument doc, Sheet sheet, WorksheetPart part) : Node
{
    public override string Kind => "sheet";
    public override object Anchor => part;
    public override string? Name => sheet.Name?.Value;

    internal SheetData Data => part.Worksheet!.GetFirstChild<SheetData>() ?? part.Worksheet.AppendChild(new SheetData());
    internal XlsxDocument Doc => doc;
    internal WorksheetPart Part => part;
    internal Worksheet Ws => part.Worksheet!;
    internal string SheetName => sheet.Name?.Value ?? "";
    internal XlsxLayout.Grid Grid => new(Ws);

    protected override IEnumerable<Node> ProjectChildren() =>
        Data.Elements<Row>().Where(r => r.RowIndex?.Value is not null).Select(r => (Node)new XlsxRow(doc, this, (int)r.RowIndex!.Value)).Concat(XlsxCharts.Of(doc, this)).Concat(XlsxImages.Of(doc, this));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["name"] = sheet.Name?.Value ?? "" };
        var cells = Data.Elements<Row>().SelectMany(r => r.Elements<Cell>()).Where(c => c.CellValue is not null || c.CellFormula is not null || c.InlineString is not null)
            .Select(XlsxCells.Position).ToList();
        if (cells.Count > 0)
            props["range"] = XlsxCells.Reference(cells.Min(c => c.Col), cells.Min(c => c.Row)) + ":" + XlsxCells.Reference(cells.Max(c => c.Col), cells.Max(c => c.Row));
        if (sheet.SheetId?.Value is { } id) props["id"] = id.ToString(CultureInfo.InvariantCulture);
        if (XlsxLayout.Merges(Ws) is { } merges) props["merges"] = merges;
        if (XlsxLayout.Widths(Ws) is { } widths) props["widths"] = widths;
        if (XlsxLayout.Heights(Data) is { } heights) props["heights"] = heights;
        if (XlsxLayout.Freeze(Ws) is { } freeze) props["freeze"] = freeze;
        if (XlsxLayout.Filter(Ws) is { } filter) props["filter"] = filter;
        return props;
    }

    protected override Node? ProjectVirtual(string kind, string key) => kind switch
    {
        "row" when int.TryParse(key, NumberStyles.None, CultureInfo.InvariantCulture, out var index) && index >= 1 => new XlsxRow(doc, this, index),
        "cell" => Cell(key),
        "range" => new XlsxRange(doc, this, key),
        _ => null,
    };

    XlsxCell Cell(string reference)
    {
        var (col, row) = XlsxCells.Parse(reference);
        return new XlsxCell(doc, this, col, row);
    }

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "name":
                var old = SheetName;
                sheet.Name = XlsxDocument.ValidName(value);
                if (sheet.Name.Value != old) XlsxCharts.RenameSheet(doc, old, sheet.Name.Value!);
                break;
            case "merges": XlsxLayout.SetMerges(Ws, value); break;
            case "widths": XlsxLayout.SetWidths(Ws, value); break;
            case "heights": XlsxLayout.SetHeights(Data, value); break;
            case "freeze": XlsxLayout.SetFreeze(Ws, value); break;
            case "filter": XlsxLayout.SetFilter(doc, sheet, Ws, value); break;
        }
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        if (kind == "chart") return XlsxCharts.Add(doc, this, props);
        if (kind == "image") return XlsxImages.Add(doc, this, props);
        if (kind != "row")
            throw new WriterException(ErrorCode.Validation, "Cells are addressed, not added", "Write directly: writer set file.xlsx /sheet[1]/cell[B3] --prop value=42");
        var last = Data.Elements<Row>().Select(r => (int)(r.RowIndex?.Value ?? 0)).DefaultIfEmpty(0).Max();
        var node = new XlsxRow(doc, this, last + 1);
        XlsxCells.GetOrCreateRow(Data, last + 1);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override string GetRaw() => part.Worksheet!.OuterXml;

    public override void SetRaw(string raw)
    {
        try
        {
            part.Worksheet = new Worksheet(RawXml.WithNamespaces(raw, doc.Namespaces));
        }
        catch (Exception ex) when (ex is not WriterException)
        {
            throw new WriterException(ErrorCode.Validation, $"Raw worksheet XML could not be parsed: {ex.Message}", "Start from 'get --raw' output.");
        }
    }

    public override void Remove() => doc.RemoveSheet(sheet, part);

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent.Kind != "document")
            throw new WriterException(ErrorCode.Validation, "Sheets can only be reordered within the workbook", "Use --to / with --index.");
        doc.MoveSheet(sheet, index);
    }
}

sealed class XlsxRow(XlsxDocument doc, XlsxSheet sheet, int index) : Node
{
    readonly object _placeholder = new();

    public override string Kind => "row";
    public override string Key => index.ToString(CultureInfo.InvariantCulture);
    public override object Anchor => (object?)Current ?? _placeholder;

    Row? Current => XlsxCells.FindRow(sheet.Data, index);

    protected override IEnumerable<Node> ProjectChildren() =>
        Current?.Elements<Cell>().Select(c => (Node)new XlsxCell(doc, sheet, XlsxCells.Position(c).Col, index)) ?? [];

    public override IReadOnlyDictionary<string, string> GetProps() =>
        new Dictionary<string, string> { ["data"] = NodeJson.Compact(w => WriteData(w)) };

    void WriteData(Utf8JsonWriter w)
    {
        w.WriteStartArray();
        var cells = Current?.Elements<Cell>().ToDictionary(c => XlsxCells.Position(c).Col) ?? new Dictionary<int, Cell>();
        var last = cells.Count == 0 ? 0 : cells.Keys.Max();
        for (var col = 1; col <= last; col++)
            w.WriteStringValue(cells.TryGetValue(col, out var cell) ? XlsxCells.Display(doc, cell) : "");
        w.WriteEndArray();
    }

    public override void SetProp(string name, string value)
    {
        if (name != "data") return;
        var values = ParseValues(value);
        for (var col = 1; col <= values.Count; col++)
            XlsxCells.SetValue(doc, XlsxCells.GetOrCreateCell(sheet.Data, col, index), values[col - 1]);
    }

    internal static List<string> ParseValues(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Array)
            throw new WriterException(ErrorCode.Validation, "data must be a JSON array", "Example: [\"Ann\",90]");
        return parsed.RootElement.EnumerateArray().Select(CellText).ToList();
    }

    internal static string CellText(JsonElement e) => e.ValueKind switch
    {
        JsonValueKind.String => e.GetString()!,
        JsonValueKind.Null => "",
        _ => e.GetRawText(),
    };

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index2)
    {
        var last = Current?.Elements<Cell>().Select(c => XlsxCells.Position(c).Col).DefaultIfEmpty(0).Max() ?? 0;
        var node = new XlsxCell(doc, sheet, last + 1, index);
        XlsxCells.GetOrCreateCell(sheet.Data, last + 1, index);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override string GetRaw() => Current?.OuterXml ?? $"<row r=\"{index}\"/>";
    public override void SetRaw(string raw) => RawXml.Replace(Current ?? XlsxCells.GetOrCreateRow(sheet.Data, index), raw, doc.Namespaces);
    public override void Remove() => Current?.Remove();
}

sealed class XlsxCell(XlsxDocument doc, XlsxSheet sheet, int col, int row) : Node
{
    readonly object _placeholder = new();
    string? _written;
    string? _borderColor;

    public override string Kind => "cell";
    public override string Key => XlsxCells.Reference(col, row);
    public override object Anchor => (object?)Current ?? _placeholder;

    Cell? Current => XlsxCells.FindRow(sheet.Data, row) is { } r ? XlsxCells.FindCell(r, col) : null;
    Cell Ensure() => XlsxCells.GetOrCreateCell(sheet.Data, col, row);
    WorksheetPart Part => (WorksheetPart)sheet.Anchor;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string>();
        var cell = Current;
        var display = cell is null ? "" : XlsxCells.Display(doc, cell);
        props["text"] = display;
        props["value"] = display;
        if (cell is not null)
        {
            if (XlsxCells.TypeOf(doc, cell) is { } type) props["type"] = type;
            if (cell.CellFormula?.Text is { Length: > 0 } formula) props["formula"] = formula;
            doc.Styles.Read(cell, props);
        }
        if (XlsxNotes.Link(Part, Key) is { } link) props["link"] = link;
        if (XlsxNotes.Note(Part, Key) is { } note) props["note"] = note;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        var cell = Ensure();
        switch (name)
        {
            case "value":
                _written = value;
                XlsxCells.SetValue(doc, cell, value);
                break;
            case "type": XlsxCells.SetTyped(doc, cell, value, _written ?? XlsxCells.Display(doc, cell)); break;
            case "formula": XlsxCells.SetFormula(doc, cell, value); break;
            case "link": XlsxNotes.SetLink(Part, Key, value); break;
            case "note": XlsxNotes.SetNote(Part, Key, value); break;
            default:
                if (name == "borderColor") _borderColor = value;
                doc.Styles.Set(cell, name, value, _borderColor);
                break;
        }
    }

    public override string GetRaw() => Current?.OuterXml ?? $"<c r=\"{Key}\"/>";
    public override void SetRaw(string raw) => RawXml.Replace(Ensure(), raw, doc.Namespaces);

    public override void Remove()
    {
        Current?.Remove();
        XlsxNotes.RemoveLink(Part, Key);
        XlsxNotes.RemoveNote(Part, Key);
    }
}

sealed class XlsxRange(XlsxDocument doc, XlsxSheet sheet, string reference) : Node
{
    readonly (int Col1, int Row1, int Col2, int Row2) _box = XlsxCells.ParseRange(reference);
    string? _borderColor;

    public override string Kind => "range";
    public override string Key => XlsxCells.Reference(_box.Col1, _box.Row1) + ":" + XlsxCells.Reference(_box.Col2, _box.Row2);
    public override object Anchor => Key;

    IEnumerable<(int Col, int Row)> Cells()
    {
        for (var r = _box.Row1; r <= _box.Row2; r++)
            for (var c = _box.Col1; c <= _box.Col2; c++)
                yield return (c, r);
    }

    /// <summary>The values, plus every formatting property that all cells in the range share.</summary>
    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string>
        {
            ["values"] = NodeJson.Compact(w =>
            {
                w.WriteStartArray();
                for (var r = _box.Row1; r <= _box.Row2; r++)
                {
                    w.WriteStartArray();
                    var row = XlsxCells.FindRow(sheet.Data, r);
                    for (var c = _box.Col1; c <= _box.Col2; c++)
                    {
                        var cell = row is null ? null : XlsxCells.FindCell(row, c);
                        w.WriteStringValue(cell is null ? "" : XlsxCells.Display(doc, cell));
                    }
                    w.WriteEndArray();
                }
                w.WriteEndArray();
            }),
        };
        Dictionary<string, string>? shared = null;
        foreach (var (c, r) in Cells())
        {
            var cell = new XlsxCell(doc, sheet, c, r).GetProps();
            shared = (shared is null ? cell.Where(p => XlsxStyles.Props.Contains(p.Key)) : shared.Where(p => cell.GetValueOrDefault(p.Key) == p.Value)).ToDictionary();
        }
        foreach (var (name, value) in shared ?? new Dictionary<string, string>()) props[name] = value;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        if (name != "values")
        {
            if (name == "borderColor") _borderColor = value;
            foreach (var (c, r) in Cells()) doc.Styles.Set(XlsxCells.GetOrCreateCell(sheet.Data, c, r), name, value, _borderColor);
            return;
        }
        var rows = value.TrimStart().StartsWith('[') ? ParseRows(value) : XlsxCells.ParseCsv(value);
        for (var r = 0; r < rows.Count; r++)
            for (var c = 0; c < rows[r].Count; c++)
                XlsxCells.SetValue(doc, XlsxCells.GetOrCreateCell(sheet.Data, _box.Col1 + c, _box.Row1 + r), rows[r][c]);
    }

    static List<List<string>> ParseRows(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Array)
            throw new WriterException(ErrorCode.Validation, "values must be a JSON array of rows or CSV text", "Example: [[\"Name\",\"Score\"],[\"Ann\",90]]");
        return parsed.RootElement.EnumerateArray()
            .Select(row => row.ValueKind == JsonValueKind.Array ? row.EnumerateArray().Select(XlsxRow.CellText).ToList() : [XlsxRow.CellText(row)])
            .ToList();
    }

    public override void Remove()
    {
        var part = (WorksheetPart)sheet.Anchor;
        foreach (var (c, r) in Cells())
        {
            XlsxNotes.RemoveLink(part, XlsxCells.Reference(c, r));
            XlsxNotes.RemoveNote(part, XlsxCells.Reference(c, r));
        }
        for (var r = _box.Row1; r <= _box.Row2; r++)
        {
            var row = XlsxCells.FindRow(sheet.Data, r);
            if (row is null) continue;
            foreach (var cell in row.Elements<Cell>().Where(c => XlsxCells.Position(c).Col is var col && col >= _box.Col1 && col <= _box.Col2).ToList()) cell.Remove();
        }
    }
}
