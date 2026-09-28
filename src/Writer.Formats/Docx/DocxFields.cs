using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

static class DocxFields
{
    public static bool Editable(string? instruction) => Regex.IsMatch(instruction ?? "", @"^\s*(REF|MERGEFIELD)\s+", RegexOptions.IgnoreCase);
    public static W.SimpleField Create(string instruction, OpenXmlElement content)
    {
        if (!Regex.IsMatch(instruction, @"^\s*(REF|MERGEFIELD)\s+[\p{L}_][\p{L}\p{N}_]{0,63}(?:\s+\\h)?\s*$", RegexOptions.IgnoreCase))
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
        DocxToc.UpdateFieldsOnOpen(doc);
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
