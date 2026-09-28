using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
namespace Writer.Formats.Xlsx;

static class XlsxTables
{
    internal static string Read(XlsxSheet sheet)
    {
        var result = new JsonArray();
        foreach (var part in sheet.Part.TableDefinitionParts)
        {
            var t = part.Table; if (t is null) continue;
            result.Add((JsonNode)new JsonObject { ["id"] = t.Id?.Value, ["name"] = t.DisplayName?.Value ?? t.Name?.Value,
                ["range"] = t.Reference?.Value, ["header"] = t.HeaderRowCount?.Value != 0, ["totals"] = t.TotalsRowCount?.Value == 1,
                ["style"] = t.TableStyleInfo?.Name?.Value ?? "TableStyleMedium2", ["stripes"] = t.TableStyleInfo?.ShowRowStripes?.Value ?? true,
                ["columns"] = new JsonArray((t.TableColumns?.Elements<TableColumn>() ?? []).Select(c => (JsonNode)new JsonObject {
                    ["id"] = c.Id?.Value, ["name"] = c.Name?.Value, ["formula"] = c.CalculatedColumnFormula?.Text,
                    ["total"] = c.TotalsRowFunction?.InnerText, ["label"] = c.TotalsRowLabel?.Value }).ToArray()) });
        }
        return result.ToJsonString();
    }
    internal static void Write(XlsxDocument doc, XlsxSheet sheet, string json)
    {
        var list = JsonNode.Parse(json) as JsonArray ?? throw Bad("tables must be an array");
        var elsewhere = doc.Sheets.Where(s => s.Part != sheet.Part).SelectMany(s => s.Part.TableDefinitionParts).Select(p => p.Table).OfType<Table>().ToList();
        var names = elsewhere.Select(t => t.DisplayName?.Value ?? "").ToHashSet(StringComparer.OrdinalIgnoreCase);
        var boxes = new List<(int Col1, int Row1, int Col2, int Row2)>();
        // Validate the complete request before touching parts.
        foreach (var node in list)
        {
            var name = node?["name"]?.GetValue<string>() ?? "";
            if (name.Length > 255 || !Regex.IsMatch(name, @"^[\p{L}_\\][\p{L}\p{N}_.\\]*$") || Regex.IsMatch(name, @"^(?:[A-Za-z]{1,3}[1-9]\d*|[RrCc])$") || !names.Add(name)) throw Bad("Invalid or duplicate table name: " + name);
            var box = XlsxCells.ParseRange(node!["range"]!.GetValue<string>());
            var header = node["header"]?.GetValue<bool>() != false; var totals = node["totals"]?.GetValue<bool>() == true;
            if (box.Row2 - box.Row1 + 1 <= (header ? 1 : 0) + (totals ? 1 : 0)) throw Bad("A table needs at least one data row");
            if (boxes.Any(b => box.Col1 <= b.Col2 && box.Col2 >= b.Col1 && box.Row1 <= b.Row2 && box.Row2 >= b.Row1)) throw Bad("Tables cannot overlap");
            if ((sheet.Ws.GetFirstChild<MergeCells>()?.Elements<MergeCell>() ?? []).Any(m => { var b = XlsxCells.ParseRange(m.Reference!.Value!); return box.Col1 <= b.Col2 && box.Col2 >= b.Col1 && box.Row1 <= b.Row2 && box.Row2 >= b.Row1; })) throw Bad("Tables cannot overlap merged cells");
            boxes.Add(box);
        }
        var old = sheet.Part.TableDefinitionParts.ToList(); var keep = new HashSet<TableDefinitionPart>();
        uint next = elsewhere.Concat(old.Select(p => p.Table).OfType<Table>()).Select(t => t.Id?.Value ?? 0).DefaultIfEmpty().Max() + 1;
        var refs = sheet.Ws.GetFirstChild<TableParts>() ?? new TableParts(); if (refs.Parent is null) sheet.Ws.AddChild(refs, true);
        foreach (var node in list)
        {
            var n = node!; var name = n["name"]!.GetValue<string>(); var id = n["id"]?.GetValue<uint>();
            var part = old.FirstOrDefault(p => !keep.Contains(p) && (id is not null ? p.Table!.Id?.Value == id : string.Equals(p.Table!.DisplayName?.Value, name, StringComparison.OrdinalIgnoreCase)));
            if (part is null) { part = sheet.Part.AddNewPart<TableDefinitionPart>(); part.Table = new Table { Id = next++ }; refs.Append(new TablePart { Id = sheet.Part.GetIdOfPart(part) }); }
            keep.Add(part); var t = part.Table!; var reference = n["range"]!.GetValue<string>(); var box = XlsxCells.ParseRange(reference);
            t.Name = name; t.DisplayName = name; t.Reference = reference; t.HeaderRowCount = n["header"]?.GetValue<bool>() == false ? 0U : 1U;
            t.TotalsRowCount = n["totals"]?.GetValue<bool>() == true ? 1U : 0U; t.TotalsRowShown = t.TotalsRowCount.Value == 1;
            var columns = t.TableColumns ??= new TableColumns(); var previous = columns.Elements<TableColumn>().ToList(); columns.RemoveAllChildren();
            var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase); var input = n["columns"] as JsonArray;
            for (var i = 0; i <= box.Col2 - box.Col1; i++)
            {
                var item = input is not null && i < input.Count ? input[i] : null;
                var cname = item?["name"]?.GetValue<string>() ?? (t.HeaderRowCount.Value == 1 ? XlsxCells.Display(doc, XlsxCells.GetOrCreateCell(sheet.Data, box.Col1 + i, box.Row1)) : "");
                if (string.IsNullOrWhiteSpace(cname)) cname = "Column" + (i + 1);
                var stem = cname; for (var suffix = 2; !unique.Add(cname); suffix++) cname = stem + suffix;
                var cid = item?["id"]?.GetValue<uint>(); var col = previous.FirstOrDefault(c => cid is not null ? c.Id?.Value == cid : c.Name?.Value == cname)?.CloneNode(true) as TableColumn ?? new TableColumn();
                col.Id = (uint)i + 1; col.Name = cname;
                if (item is not null && item.AsObject().ContainsKey("formula")) col.CalculatedColumnFormula = item["formula"] is { } f ? new CalculatedColumnFormula(f.GetValue<string>().TrimStart('=')) : null;
                if (item is not null && item.AsObject().ContainsKey("total")) col.TotalsRowFunction = item["total"] is { } total ? new DocumentFormat.OpenXml.EnumValue<TotalsRowFunctionValues> { InnerText = total.GetValue<string>() } : null;
                if (item is not null && item.AsObject().ContainsKey("label")) col.TotalsRowLabel = item["label"]?.GetValue<string>();
                columns.Append(col);
                if (t.HeaderRowCount.Value == 1) { var cell = XlsxCells.GetOrCreateCell(sheet.Data, box.Col1 + i, box.Row1); XlsxCells.SetValue(doc, sheet, cell, cname); XlsxCells.SetString(doc, cell, cname); }
            }
            columns.Count = (uint)columns.ChildElements.Count;
            if (t.HeaderRowCount.Value == 1) { var af = t.AutoFilter ??= new AutoFilter(); af.Reference = XlsxCells.Reference(box.Col1, box.Row1) + ":" + XlsxCells.Reference(box.Col2, box.Row2 - (int)t.TotalsRowCount.Value); }
            else t.AutoFilter = null;
            var style = t.TableStyleInfo ??= new TableStyleInfo(); style.Name = n["style"]?.GetValue<string>() ?? "TableStyleMedium2"; style.ShowRowStripes = n["stripes"]?.GetValue<bool>() != false;
            style.ShowColumnStripes ??= false; style.ShowFirstColumn ??= false; style.ShowLastColumn ??= false;
        }
        foreach (var part in old.Where(p => !keep.Contains(p))) { XlsxCalculation.ConvertTableReferences(doc, doc.Sheets.FindIndex(s => s.Part == sheet.Part), part.Table!.DisplayName!.Value!); var rel = sheet.Part.GetIdOfPart(part); foreach (var r in refs.Elements<TablePart>().Where(r => r.Id?.Value == rel).ToList()) r.Remove(); sheet.Part.DeletePart(part); }
        refs.Count = (uint)refs.ChildElements.Count; if (!refs.HasChildren) refs.Remove();
        doc.RecalculateOnLoad();
    }
    internal static string ColumnReference(string formula, string table, string from, string to, bool local)
    {
        var result = new System.Text.StringBuilder(); var i = 0;
        while (i < formula.Length)
        {
            if (formula[i] == '"') { var start = i++; while (i < formula.Length) if (formula[i++] == '"') { if (i < formula.Length && formula[i] == '"') i++; else break; } result.Append(formula[start..i]); continue; }
            var match = Regex.Match(formula[i..], @"^([\p{L}_\\][\p{L}\p{N}_.\\]*)?\[");
            if (!match.Success) { result.Append(formula[i++]); continue; }
            var end = i + match.Length - 1; var depth = 0;
            for (; end < formula.Length; end++) { if (formula[end] == '\'' && end + 1 < formula.Length && "[]#@'".Contains(formula[end + 1])) { end++; continue; } if (formula[end] == '[') depth++; else if (formula[end] == ']' && --depth == 0) break; }
            end = Math.Min(end + 1, formula.Length); var token = formula[i..end]; var name = match.Groups[1].Value;
            if (name.Length == 0 ? local : name.Equals(table, StringComparison.OrdinalIgnoreCase))
                token = Regex.Replace(token, @"\[((?:'[\s\S]|[^\[\]])*)\]", m => {
                    var raw = m.Groups[1].Value; var prefix = raw.StartsWith('@') ? "@" : ""; var text = Regex.Replace(raw[prefix.Length..], @"'([\[\]#@'])", "$1");
                    return text.Equals(from, StringComparison.OrdinalIgnoreCase) ? "[" + prefix + Regex.Replace(to, @"[\[\]#@']", "'$0") + "]" : m.Value;
                });
            result.Append(token); i = end;
        }
        return result.ToString();
    }
    internal static void HeaderChanged(XlsxDocument doc, XlsxSheet sheet, Cell cell, string value)
    {
        var (c, r) = XlsxCells.Position(cell);
        foreach (var part in sheet.Part.TableDefinitionParts)
        {
            var t = part.Table; if (t?.Reference?.Value is not { } reference || t.HeaderRowCount?.Value == 0) continue;
            var box = XlsxCells.ParseRange(reference); if (r != box.Row1 || c < box.Col1 || c > box.Col2) continue;
            var col = t.TableColumns?.Elements<TableColumn>().ElementAtOrDefault(c - box.Col1); if (col is null || col.Name?.Value == value) continue;
            var old = col.Name?.Value ?? ""; var name = string.IsNullOrWhiteSpace(value) ? "Column" + (c - box.Col1 + 1) : value; var stem = name; var n = 2;
            while (t.TableColumns!.Elements<TableColumn>().Any(x => x != col && string.Equals(x.Name?.Value, name, StringComparison.OrdinalIgnoreCase))) name = stem + n++;
            col.Name = name; XlsxCells.SetString(doc, cell, name);
            foreach (var owner in doc.Sheets) foreach (var item in owner.Part.Worksheet!.Descendants<Cell>().Where(x => x.CellFormula is not null))
            {
                var (cc, rr) = XlsxCells.Position(item); var local = owner.Part == sheet.Part && cc >= box.Col1 && cc <= box.Col2 && rr >= box.Row1 && rr <= box.Row2;
                item.CellFormula!.Text = ColumnReference(item.CellFormula.Text, t.DisplayName!.Value!, old, name, local);
            }
            foreach (var calculated in t.TableColumns!.Elements<TableColumn>().Select(x => x.CalculatedColumnFormula).OfType<CalculatedColumnFormula>()) calculated.Text = ColumnReference(calculated.Text, t.DisplayName!.Value!, old, name, true);
            doc.RecalculateOnLoad();
        }
    }
    static WriterException Bad(string message) => new(ErrorCode.Validation, message, "Use unique table names, a non-overlapping range with a header and data, and optional columns/style.");
}
