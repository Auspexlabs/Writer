using System.Globalization;
using System.Security;
using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
namespace Writer.Formats.Xlsx;

static class XlsxPivots
{
    const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    static string Esc(string? s) => SecurityElement.Escape(s ?? "") ?? "";
    static int Index(OpenXmlElement e, string key) => int.TryParse(e.GetAttribute(key, "").Value, out var i) ? i : -1;
    internal static JsonArray Read(XlsxSheet sheet)
    {
        var result = new JsonArray();
        foreach (var part in sheet.Part.PivotTableParts)
        {
            var pivot = part.PivotTableDefinition; if (pivot is null) continue;
            var cache = part.PivotTableCacheDefinitionPart?.PivotCacheDefinition; var source = cache?.CacheSource?.GetFirstChild<WorksheetSource>();
            var rows = pivot.GetFirstChild<RowFields>()?.ChildElements.Select(x => (JsonNode?)JsonValue.Create(Index(x, "x"))).ToArray() ?? [];
            var cols = pivot.GetFirstChild<ColumnFields>()?.ChildElements.Select(x => (JsonNode?)JsonValue.Create(Index(x, "x"))).ToArray() ?? [];
            var model = new JsonObject { ["id"] = part.Uri.ToString(), ["name"] = pivot.Name?.Value, ["sourceSheet"] = source?.Sheet?.Value,
                ["sourceRange"] = source?.Reference?.Value, ["sourceTable"] = source?.Name?.Value, ["range"] = pivot.Location?.Reference?.Value,
                ["target"] = pivot.Location?.Reference?.Value?.Split(':')[0] ?? "A1", ["rows"] = new JsonArray(rows), ["cols"] = new JsonArray(cols),
                ["values"] = new JsonArray((pivot.GetFirstChild<DataFields>()?.Elements<DataField>() ?? []).Select(v => (JsonNode)new JsonObject { ["field"] = v.Field is { } field ? (int)field.Value : -1, ["fn"] = v.Subtotal?.InnerText ?? "sum", ["name"] = v.Name?.Value }).ToArray()),
                ["fields"] = new JsonArray((cache?.CacheFields?.Elements<CacheField>() ?? []).Select(f => (JsonNode?)JsonValue.Create(f.Name?.Value)).ToArray()),
                ["style"] = pivot.GetFirstChild<PivotTableStyle>()?.Name?.Value ?? "PivotStyleMedium9" };
            if (source?.Sheet?.Value is { } sourceName && source.Reference?.Value is { } sourceRange)
            {
                var owner = sheet.Doc.Sheets.FirstOrDefault(s => string.Equals(s.Sheet.Name?.Value, sourceName, StringComparison.OrdinalIgnoreCase));
                if (owner.Part is not null)
                {
                    var box = XlsxCells.ParseRange(sourceRange); var header = owner.Part.Worksheet!.GetFirstChild<SheetData>()?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == box.Row1);
                    var fields = Enumerable.Range(box.Col1, box.Col2 - box.Col1 + 1).Select(c => header is not null && XlsxCells.FindCell(header, c) is { } cell ? XlsxCells.Display(sheet.Doc, cell) : "").ToArray();
                    var cached = model["fields"]!.AsArray().Select(f => f?.GetValue<string>() ?? "").ToArray();
                    int Map(int index) => index >= 0 && index < cached.Length ? Array.FindIndex(fields, f => string.Equals(f, cached[index], StringComparison.OrdinalIgnoreCase)) : index;
                    foreach (var key in new[] { "rows", "cols" }) model[key] = new JsonArray(model[key]!.AsArray().Select(v => (JsonNode?)JsonValue.Create(Map(v!.GetValue<int>()))).ToArray());
                    foreach (var v in model["values"]!.AsArray()) v!["field"] = Map(v["field"]!.GetValue<int>());
                    model["fields"] = new JsonArray(fields.Select(f => (JsonNode?)JsonValue.Create(f)).ToArray());
                }
            }
            result.Add((JsonNode)model);
        }
        return result;
    }
    sealed record Value(string Type, string Text)
    {
        public string Xml => Type == "m" ? "<m/>" : $"<{Type} v=\"{Esc(Text)}\"/>";
        public double? Number => Type == "n" && double.TryParse(Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n : null;
    }
    internal static void Write(XlsxDocument doc, XlsxSheet sheet, string json)
    {
        var list = JsonNode.Parse(json) as JsonArray ?? throw Bad("pivots must be an array");
        var old = sheet.Part.PivotTableParts.ToList(); var keep = new HashSet<PivotTablePart>();
        var before = Read(sheet);
        foreach (var node in list)
        {
            var n = node as JsonObject ?? throw Bad("Each pivot must be an object");
            var name = n["name"]?.GetValue<string>() ?? "";
            if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || doc.Sheets.Where(s => s.Part != sheet.Part).SelectMany(s => s.Part.PivotTableParts).Any(p => string.Equals(p.PivotTableDefinition?.Name?.Value, name, StringComparison.OrdinalIgnoreCase))) throw Bad("Invalid or duplicate pivot name");
            var part = old.FirstOrDefault(p => n["id"]?.GetValue<string>() is { } id ? p.Uri.ToString() == id : p.PivotTableDefinition?.Name?.Value == name);
            if (part is not null && keep.Contains(part)) throw Bad("Duplicate pivot table");
            if (part is not null && JsonNode.DeepEquals(n, before.FirstOrDefault(p => p!["id"]!.GetValue<string>() == part.Uri.ToString()))) { keep.Add(part); continue; }
            if (part is not null && before.FirstOrDefault(p => p!["id"]!.GetValue<string>() == part.Uri.ToString()) is { } prior)
            {
                var a = n.DeepClone().AsObject(); var b = prior.DeepClone().AsObject(); foreach (var key in new[] { "sourceSheet", "sourceRange", "target", "range", "name", "sourceChanged" }) { a.Remove(key); b.Remove(key); }
                if (n["sourceChanged"]?.GetValue<bool>() == true || JsonNode.DeepEquals(a, b)) { part.PivotTableDefinition!.Name = name; if (part.PivotTableDefinition.Location is { } location && n["range"] is { } range) location.Reference = range.GetValue<string>(); if (part.PivotTableCacheDefinitionPart?.PivotCacheDefinition?.CacheSource?.GetFirstChild<WorksheetSource>() is { } source) { source.Sheet = n["sourceSheet"]?.GetValue<string>(); source.Reference = n["sourceRange"]?.GetValue<string>(); } keep.Add(part); continue; }
            }
            Build(doc, sheet, n, part, out var created); keep.Add(created);
        }
        foreach (var part in old.Where(p => !keep.Contains(p))) { ClearOutput(sheet, part.PivotTableDefinition?.Location?.Reference?.Value); sheet.Part.DeletePart(part); }
        var used = doc.Sheets.SelectMany(s => s.Part.PivotTableParts).Select(p => p.PivotTableCacheDefinitionPart).ToHashSet();
        foreach (var cache in doc.Workbook.PivotTableCacheDefinitionParts.Where(c => !used.Contains(c) && c.PivotCacheDefinition?.RefreshedBy?.Value == "Writer").ToList()) { var rel = doc.Workbook.GetIdOfPart(cache); foreach (var item in doc.Workbook.Workbook!.PivotCaches?.Elements<PivotCache>().Where(c => c.Id?.Value == rel).ToList() ?? []) item.Remove(); doc.Workbook.DeletePart(cache); }
        if (doc.Workbook.Workbook!.PivotCaches?.HasChildren == false) doc.Workbook.Workbook.PivotCaches.Remove();
        doc.RecalculateOnLoad();
    }
    static void ClearOutput(XlsxSheet sheet, string? reference)
    {
        if (reference is null) return; var box = XlsxCells.ParseRange(reference);
        foreach (var row in sheet.Data.Elements<Row>().Where(r => r.RowIndex?.Value >= box.Row1 && r.RowIndex?.Value <= box.Row2)) foreach (var cell in row.Elements<Cell>().ToList()) { var (c, _) = XlsxCells.Position(cell); if (c >= box.Col1 && c <= box.Col2) { XlsxCells.SetValue(sheet.Doc, sheet, cell, ""); if (cell.StyleIndex is null) cell.Remove(); } }
    }
    sealed class AggregateState
    {
        int count, numbers; double sum, min = double.PositiveInfinity, max = double.NegativeInfinity; string? error;
        internal void Add(Value value) { if (value.Type != "m") count++; if (value.Type == "e") error = value.Text; if (value.Number is { } n) { numbers++; sum += n; min = Math.Min(min, n); max = Math.Max(max, n); } }
        internal string Result(string fn) => fn == "count" ? count.ToString(CultureInfo.InvariantCulture) : fn == "countNums" ? numbers.ToString(CultureInfo.InvariantCulture) : error ?? (numbers == 0 ? "" : (fn switch { "sum" => sum, "average" => sum / numbers, "min" => min, _ => max }).ToString("G15", CultureInfo.InvariantCulture));
    }
    static void Build(XlsxDocument doc, XlsxSheet target, JsonObject spec, PivotTablePart? previous, out PivotTablePart part)
    {
        var sourceName = spec["sourceSheet"]?.GetValue<string>(); var reference = spec["sourceRange"]?.GetValue<string>();
        if (spec["sourceTable"]?.GetValue<string>() is { Length: > 0 } tableName)
            foreach (var s in doc.Sheets) if (s.Part.TableDefinitionParts.Select(p => p.Table).FirstOrDefault(t => string.Equals(t?.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase)) is { } table) { sourceName = s.Sheet.Name!.Value; var b = XlsxCells.ParseRange(table.Reference!.Value!); reference = XlsxCells.Reference(b.Col1, b.Row1) + ":" + XlsxCells.Reference(b.Col2, b.Row2 - (int)(table.TotalsRowCount?.Value ?? 0)); break; }
        var source = doc.Sheets.FirstOrDefault(s => string.Equals(s.Sheet.Name?.Value, sourceName, StringComparison.OrdinalIgnoreCase));
        if (source.Part is null || reference is null) throw Bad("The pivot needs a worksheet source");
        var box = XlsxCells.ParseRange(reference); var width = box.Col2 - box.Col1 + 1; var rowFields = spec["rows"]?.AsArray().Select(n => n!.GetValue<int>()).ToArray() ?? [];
        var colFields = spec["cols"]?.AsArray().Select(n => n!.GetValue<int>()).ToArray() ?? []; var values = spec["values"] as JsonArray ?? [];
        if (rowFields.Length != 1 || colFields.Length > 1 || values.Count != 1) throw Bad("Choose one row field, an optional column field and one value field");
        var rowField = rowFields[0]; var colField = colFields.Length == 0 ? -1 : colFields[0]; var valueField = values[0]?["field"]?.GetValue<int>() ?? -1; var fn = values[0]?["fn"]?.GetValue<string>() ?? "sum";
        if (rowField < 0 || rowField >= width || colField >= width || valueField < 0 || valueField >= width || fn is not ("sum" or "count" or "countNums" or "average" or "min" or "max")) throw Bad("Invalid pivot field or aggregation");
        if (rowField == colField || rowField == valueField || colField == valueField) throw Bad("Choose different row, column and value fields");
        doc.Calculate();
        var sourceSheet = new XlsxSheet(doc, source.Sheet, source.Part); var sourceRows = sourceSheet.Data.Elements<Row>().Where(r => r.RowIndex?.Value >= box.Row1 && r.RowIndex?.Value <= box.Row2).ToDictionary(r => (int)r.RowIndex!.Value, r => r.Elements<Cell>().ToDictionary(c => XlsxCells.Position(c).Col));
        Value Get(int r, int c) { var cell = sourceRows.GetValueOrDefault(r)?.GetValueOrDefault(c); if (cell is null) return new("m", ""); var text = XlsxCells.Display(doc, cell); if (text.Length == 0) return new("m", ""); var type = XlsxCells.TypeOf(doc, cell); return type == "number" || type == "date" ? new("n", cell.CellValue?.Text ?? text) : type == "bool" ? new("b", text == "true" ? "1" : "0") : type == "error" ? new("e", text) : new("s", text); }
        var fields = Enumerable.Range(box.Col1, width).Select(c => Get(box.Row1, c).Text).ToArray();
        if (fields.Any(string.IsNullOrWhiteSpace) || fields.Distinct(StringComparer.OrdinalIgnoreCase).Count() != fields.Length) throw Bad("Pivot source headers must be nonempty and unique");
        var data = Enumerable.Range(box.Row1 + 1, Math.Max(0, box.Row2 - box.Row1)).Select(r => Enumerable.Range(box.Col1, width).Select(c => Get(r, c)).ToArray()).ToArray();
        var shared = Enumerable.Range(0, width).Select(c => data.Select(r => r[c]).Distinct().ToArray()).ToArray(); var indices = shared.Select(items => items.Select((v, i) => (v, i)).ToDictionary(x => x.v, x => x.i)).ToArray();
        var rowKeys = shared[rowField]; var colKeys = colField >= 0 ? shared[colField] : [new Value("s", "")];
        var (targetCol, targetRow) = XlsxCells.Parse(spec["target"]?.GetValue<string>() ?? "A1");
        var endCol = targetCol + colKeys.Length + (colField >= 0 ? 1 : 0); var endRow = targetRow + rowKeys.Length + 1;
        if (endCol > 16384 || endRow > 1048576) throw Bad("Pivot output exceeds the worksheet");
        var outputRange = XlsxCells.Reference(targetCol, targetRow) + ":" + XlsxCells.Reference(endCol, endRow);
        if (source.Part == target.Part && targetCol <= box.Col2 && endCol >= box.Col1 && targetRow <= box.Row2 && endRow >= box.Row1) throw Bad("Pivot output overlaps its source");
        var oldRange = previous?.PivotTableDefinition?.Location?.Reference?.Value; var oldBox = oldRange is null ? (Col1: 0, Row1: 0, Col2: 0, Row2: 0) : XlsxCells.ParseRange(oldRange);
        foreach (var r in target.Data.Elements<Row>().Where(r => r.RowIndex?.Value >= targetRow && r.RowIndex?.Value <= endRow)) foreach (var cell in r.Elements<Cell>()) { var (c, rr) = XlsxCells.Position(cell); if (c >= targetCol && c <= endCol && !(c >= oldBox.Col1 && c <= oldBox.Col2 && rr >= oldBox.Row1 && rr <= oldBox.Row2) && (cell.CellValue is not null || cell.InlineString is not null || cell.CellFormula is not null)) throw Bad("Pivot output would overwrite existing cells"); }
        var caches = doc.Workbook.Workbook!.PivotCaches ??= new PivotCaches(); var cacheId = caches.Elements<PivotCache>().Select(c => c.CacheId?.Value ?? 0).DefaultIfEmpty().Max() + 1;
        var cachePart = doc.Workbook.AddNewPart<PivotTableCacheDefinitionPart>(); var recordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
        var fieldXml = string.Concat(fields.Select((f, c) => $"<cacheField name=\"{Esc(f)}\"><sharedItems count=\"{shared[c].Length}\" containsString=\"{(shared[c].Any(v => v.Type == "s") ? 1 : 0)}\" containsNumber=\"{(shared[c].Any(v => v.Type == "n") ? 1 : 0)}\" containsBlank=\"{(shared[c].Any(v => v.Type == "m") ? 1 : 0)}\" containsSemiMixedTypes=\"{(shared[c].Select(v => v.Type).Distinct().Count() > 1 ? 1 : 0)}\">{string.Concat(shared[c].Select(v => v.Xml))}</sharedItems></cacheField>"));
        cachePart.PivotCacheDefinition = new PivotCacheDefinition($"<pivotCacheDefinition xmlns=\"{Ns}\" xmlns:r=\"{Rel}\" r:id=\"{cachePart.GetIdOfPart(recordsPart)}\" recordCount=\"{data.Length}\" refreshedBy=\"Writer\" createdVersion=\"3\" refreshedVersion=\"3\" minRefreshableVersion=\"3\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"{Esc(sourceName)}\" ref=\"{Esc(reference)}\"/></cacheSource><cacheFields count=\"{width}\">{fieldXml}</cacheFields></pivotCacheDefinition>");
        recordsPart.PivotCacheRecords = new PivotCacheRecords($"<pivotCacheRecords xmlns=\"{Ns}\" count=\"{data.Length}\">{string.Concat(data.Select(row => "<r>" + string.Concat(row.Select((v,c) => $"<x v=\"{indices[c][v]}\"/>")) + "</r>"))}</pivotCacheRecords>");
        caches.Append(new PivotCache { CacheId = cacheId, Id = doc.Workbook.GetIdOfPart(cachePart) });
        part = previous ?? target.Part.AddNewPart<PivotTablePart>(); if (part.PivotTableCacheDefinitionPart is { } oldCache) part.DeletePart(oldCache); part.AddPart(cachePart);
        var fieldDefs = string.Concat(fields.Select((_, c) => { var axis = c == rowField ? "axisRow" : c == colField ? "axisCol" : null; return axis is null ? $"<pivotField dataField=\"{(c == valueField ? 1 : 0)}\" showAll=\"0\"/>" : $"<pivotField axis=\"{axis}\" showAll=\"0\" defaultSubtotal=\"0\"><items count=\"{shared[c].Length}\">{string.Concat(shared[c].Select((v,i) => $"<item x=\"{i}\"/>"))}</items></pivotField>"; }));
        string Items(string tag, int n) => $"<{tag} count=\"{n+1}\">{string.Concat(Enumerable.Range(0,n).Select(i=>$"<i><x v=\"{i}\"/></i>"))}<i t=\"grand\"><x/></i></{tag}>";
        var caption = values[0]?["name"]?.GetValue<string>() ?? fn + " " + fields[valueField];
        part.PivotTableDefinition = new PivotTableDefinition($"<pivotTableDefinition xmlns=\"{Ns}\" name=\"{Esc(spec["name"]!.GetValue<string>())}\" cacheId=\"{cacheId}\" dataCaption=\"Values\" rowGrandTotals=\"1\" colGrandTotals=\"1\" compact=\"0\" compactData=\"0\" outline=\"0\" gridDropZones=\"1\" createdVersion=\"3\" updatedVersion=\"3\" minRefreshableVersion=\"3\"><location ref=\"{outputRange}\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"1\"/><pivotFields count=\"{width}\">{fieldDefs}</pivotFields><rowFields count=\"1\"><field x=\"{rowField}\"/></rowFields>{Items("rowItems",rowKeys.Length)}{(colField >= 0 ? $"<colFields count=\"1\"><field x=\"{colField}\"/></colFields>{Items("colItems",colKeys.Length)}" : "")}<dataFields count=\"1\"><dataField name=\"{Esc(caption)}\" fld=\"{valueField}\" subtotal=\"{fn}\"/></dataFields><pivotTableStyleInfo name=\"{Esc(spec["style"]?.GetValue<string>() ?? "PivotStyleMedium9")}\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"1\" showColStripes=\"0\" showLastColumn=\"1\"/></pivotTableDefinition>");
        ClearOutput(target, oldRange);
        void Cell(int r, int c, string value, bool text = false) { var cell = XlsxCells.GetOrCreateCell(target.Data, c, r); XlsxCells.SetValue(doc, target, cell, value); if (text) XlsxCells.SetString(doc, cell, value); else if (value.StartsWith('#')) { cell.DataType = CellValues.Error; cell.CellValue = new CellValue(value); }
            if (cell.StyleIndex is null && (r == targetRow || r == endRow || colField >= 0 && c == endCol)) { var node = new XlsxCell(doc, target, c, r); node.SetProp("bold", "true"); node.SetProp("fill", "E8EEF7"); }
        }
        var groups = new Dictionary<(int R, int C), AggregateState>();
        void Add(int r, int c, Value value) { if (!groups.TryGetValue((r,c), out var state)) groups[(r,c)] = state = new AggregateState(); state.Add(value); }
        foreach (var row in data) { var r = indices[rowField][row[rowField]]; var c = colField >= 0 ? indices[colField][row[colField]] : 0; Add(r,c,row[valueField]); Add(r,-1,row[valueField]); Add(-1,c,row[valueField]); Add(-1,-1,row[valueField]); }
        string Aggregate(int r, int c) => groups.GetValueOrDefault((r,c))?.Result(fn) ?? "";
        Cell(targetRow, targetCol, fields[rowField], true);
        for (var j = 0; j < colKeys.Length; j++) Cell(targetRow, targetCol + j + 1, colField >= 0 ? colKeys[j].Text : caption, true);
        if (colField >= 0) Cell(targetRow, endCol, "Grand Total", true);
        for (var i = 0; i < rowKeys.Length; i++) { Cell(targetRow + i + 1, targetCol, rowKeys[i].Text, true); for (var j = 0; j < colKeys.Length; j++) Cell(targetRow + i + 1, targetCol + j + 1, Aggregate(i,j)); if (colField >= 0) Cell(targetRow + i + 1, endCol, Aggregate(i,-1)); }
        Cell(endRow, targetCol, "Grand Total", true); for (var j = 0; j < colKeys.Length; j++) Cell(endRow, targetCol + j + 1, Aggregate(-1,j)); if (colField >= 0) Cell(endRow, endCol, Aggregate(-1,-1));
    }
    static WriterException Bad(string message) => new(ErrorCode.Validation, message, "Use a named pivot, worksheet source, target, rows:[field], cols:[] or [field], values:[{field,fn,name}]. Field indices start at zero.");
}
