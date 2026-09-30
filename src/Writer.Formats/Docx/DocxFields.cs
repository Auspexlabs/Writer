using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

static class DocxFields
{
    public static bool Editable(string? instruction) => Regex.IsMatch(instruction ?? "", @"^\s*(?:(REF|MERGEFIELD)\s+|=)", RegexOptions.IgnoreCase);
    public static W.SimpleField Create(string instruction, OpenXmlElement content)
    {
        if (!instruction.TrimStart().StartsWith("=", StringComparison.Ordinal) && !Regex.IsMatch(instruction, @"^\s*(REF|MERGEFIELD)\s+[\p{L}_][\p{L}\p{N}_]{0,63}(?:\s+\\h)?\s*$", RegexOptions.IgnoreCase))
            throw new WriterException(ErrorCode.Validation, "Unsupported document field", "Use REF BookmarkName \\h or MERGEFIELD FieldName.");
        return new W.SimpleField(content) { Instruction = " " + instruction.Trim() + " " };
    }
    static string Argument(W.SimpleField field) => Regex.Match(field.Instruction?.Value ?? "", @"^\s*\w+\s+([^\s]+)").Groups[1].Value.Trim('"');
    public static void Refresh(DocxDocument doc)
    {
        DocxMarks.RefreshCaptions(doc);
        var body = doc.Main.Document!.Body!;
        var all = body.Descendants().ToList();
        var bookmarks = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var i = 0; i < all.Count; i++)
        {
            if (all[i] is not W.BookmarkStart start || start.Name?.Value is not { } name) continue;
            var texts = new List<string>();
            for (var j = i + 1; j < all.Count; j++)
            {
                if (all[j] is W.BookmarkEnd end && end.Id?.Value == start.Id?.Value) break;
                if (all[j] is W.Text text && !text.Ancestors<W.DeletedRun>().Any()) texts.Add(text.Text);
            }
            bookmarks[name] = string.Concat(texts);
        }
        foreach (var field in body.Descendants<W.SimpleField>().Where(f => Regex.IsMatch(f.Instruction?.Value ?? "", @"^\s*REF\s+", RegexOptions.IgnoreCase)).ToList())
            Result(field, bookmarks.GetValueOrDefault(Argument(field), "Error! Reference source not found."));
        RefreshTableFormulas(doc, bookmarks);
        DocxToc.UpdateFieldsOnOpen(doc);
    }
    static void RefreshTableFormulas(DocxDocument doc, Dictionary<string,string> bookmarks)
    {
        var roots = new List<OpenXmlElement> { doc.Main.Document! };
        roots.AddRange(doc.Main.HeaderParts.Select(p => (OpenXmlElement)p.Header!));
        roots.AddRange(doc.Main.FooterParts.Select(p => (OpenXmlElement)p.Footer!));
        var tables = roots.SelectMany(root => root.Descendants<W.Table>()).ToList();
        if (!tables.Any(t=>t.Descendants<W.SimpleField>().Any(f=>(f.Instruction?.Value??"").TrimStart().StartsWith("=")))) return;
        var aliases = new Dictionary<W.Table,List<string>>();
        foreach(var root in roots)
        {
            var elements=root.Descendants().ToList();var positions=elements.Select((e,i)=>(e,i)).ToDictionary(x=>x.e,x=>x.i);
            var ends=elements.OfType<W.BookmarkEnd>().Where(e=>e.Id?.Value is not null).GroupBy(e=>e.Id!.Value!).ToDictionary(g=>g.Key,g=>positions[g.First()]);
            foreach(var start in elements.OfType<W.BookmarkStart>())
            {
                if(start.Name?.Value is not {} name||start.Id?.Value is not {} id||!ends.TryGetValue(id,out var end))continue;
                var begin=positions[start];
                foreach(var table in tables.Where(positions.ContainsKey))
                {
                    var text=table.Descendants<W.Text>().ToList();
                    if(text.Count==0||positions[text[0]]<=begin||positions[text[^1]]>=end)continue;
                    if(!aliases.TryGetValue(table,out var names))aliases[table]=names=[];names.Add(name);
                }
            }
        }
        var input = new JsonArray(); var targets = new List<Dictionary<(int, int), W.SimpleField>>();
        foreach (var table in tables)
        {
            var rows = new JsonArray(); var fields = new Dictionary<(int,int), W.SimpleField>(); var ri = 0;
            foreach (var row in table.Elements<W.TableRow>())
            {
                var cells = new JsonArray(); var header = row.TableRowProperties?.GetFirstChild<W.TableHeader>() is { } h && (h.Val is null || h.Val.Value == W.OnOffOnlyValues.On);
                var before = row.TableRowProperties?.GetFirstChild<W.GridBefore>()?.Val?.Value ?? 0;
                for (var i = 0; i < before; i++) cells.Add(null);
                foreach (var cell in row.Elements<W.TableCell>())
                {
                    var field = cell.Descendants<W.SimpleField>().FirstOrDefault(f => f.Ancestors<W.TableCell>().First() == cell && (f.Instruction?.Value ?? "").TrimStart().StartsWith("="));
                    if (field is not null && !DocxTable.IsContinue(cell)) fields[(ri,cells.Count)] = field;
                    cells.Add(DocxTable.IsContinue(cell) ? null : (JsonNode)new JsonObject { ["text"] = string.Concat(cell.Descendants<W.Text>().Where(t => t.Ancestors<W.TableCell>().First() == cell).Select(t => t.Text)), ["formula"] = field?.Instruction?.Value, ["header"] = header });
                    for (var c = 1; c < DocxTable.Span(cell); c++) cells.Add(null);
                }
                rows.Add((JsonNode)cells); ri++;
            }
            var names=new JsonObject();foreach(var (key,value) in bookmarks)names[key]=value;
            input.Add((JsonNode)new JsonObject{["rows"]=rows,["bookmarks"]=names,["names"]=new JsonArray((aliases.GetValueOrDefault(table)??[]).Select(s=>(JsonNode?)JsonValue.Create(s)).ToArray())}); targets.Add(fields);
        }
        using var values = Xlsx.XlsxCalculation.CalculateWordTables(input.ToJsonString()); var index = 0;
        foreach (var table in values.RootElement.EnumerateArray())
        {
            foreach (var result in table.EnumerateArray()) if (targets[index].TryGetValue((result.GetProperty("r").GetInt32(),result.GetProperty("c").GetInt32()),out var field)) { var error=result.GetProperty("error").GetString(); if(error is null || error != "#NAME?" && !error.StartsWith("Unsupported",StringComparison.Ordinal)) Result(field,result.GetProperty("text").GetString() ?? ""); }
            index++;
        }
    }

    static void Result(W.SimpleField field, string text)
    {
        if (string.Concat(field.Descendants<W.Text>().Select(t => t.Text)) == text) return;
        var properties = field.Descendants<W.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
        field.RemoveAllChildren(); var run = new W.Run(); if (properties is not null) run.Append(properties);
        run.Append(DocxRuns.TextElements(text)); field.Append(run);
    }
    public static void Merge(DocxDocument doc, string json)
    {
        using var data = JsonDocument.Parse(json);
        if (data.RootElement.ValueKind != JsonValueKind.Object) throw new WriterException(ErrorCode.Validation, "Merge data must be an object", "Provide one record with named fields.");
        var roots = new List<OpenXmlElement> { doc.Main.Document! };
        roots.AddRange(doc.Main.HeaderParts.Select(p => (OpenXmlElement)p.Header!));
        roots.AddRange(doc.Main.FooterParts.Select(p => (OpenXmlElement)p.Footer!));
        if (doc.Main.FootnotesPart?.Footnotes is { } footnotes) roots.Add(footnotes);
        if (doc.Main.EndnotesPart?.Endnotes is { } endnotes) roots.Add(endnotes);
        // Word normally writes MERGEFIELD as begin/instruction/separate/result/end runs.
        foreach (var paragraph in roots.SelectMany(r => r.Descendants<W.Paragraph>()))
        {
            var children = paragraph.ChildElements.ToList();
            for (var i = 0; i < children.Count; i++)
            {
                if (children[i] is not W.Run begin || begin.GetFirstChild<W.FieldChar>()?.FieldCharType?.Value != W.FieldCharValues.Begin) continue;
                var depth = 0; var end = i; var instruction = "";
                for (; end < children.Count; end++)
                {
                    foreach (var ch in children[end].Descendants<W.FieldChar>()) { if (ch.FieldCharType?.Value == W.FieldCharValues.Begin) depth++; else if (ch.FieldCharType?.Value == W.FieldCharValues.End) depth--; }
                    instruction += string.Concat(children[end].Descendants<W.FieldCode>().Select(c => c.Text));
                    if (depth == 0) break;
                }
                var match = Regex.Match(instruction, "^\\s*MERGEFIELD\\s+(?:\"([^\"]+)\"|([^\\s]+))", RegexOptions.IgnoreCase);
                if (end >= children.Count || !match.Success || !data.RootElement.TryGetProperty(match.Groups[1].Success ? match.Groups[1].Value : match.Groups[2].Value, out var value)) continue;
                var range = children.Skip(i).Take(end - i + 1).ToList();
                var cached = range.OfType<W.Run>().FirstOrDefault(r => r.Elements<W.Text>().Any());
                var run = new W.Run(); if (cached?.RunProperties is { } rp) run.Append(rp.CloneNode(true));
                run.Append(DocxRuns.TextElements(value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : value.ToString()));
                begin.InsertBeforeSelf(run); foreach (var child in range) child.Remove(); i = end;
            }
        }
        foreach (var field in roots.SelectMany(r => r.Descendants<W.SimpleField>()).Where(f => Regex.IsMatch(f.Instruction?.Value ?? "", @"^\s*MERGEFIELD\s+", RegexOptions.IgnoreCase)).ToList())
        {
            if (!data.RootElement.TryGetProperty(Argument(field), out var value)) continue;
            Result(field, value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : value.ToString());
            foreach (var child in field.ChildElements.ToList()) { child.Remove(); field.InsertBeforeSelf(child); } field.Remove();
        }
    }
}
