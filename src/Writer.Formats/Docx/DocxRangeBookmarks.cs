using System.Globalization;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

static class DocxRangeBookmarks
{
    static IEnumerable<Node> Paragraphs(Node node)
    {
        foreach (var child in node.Children)
            if (child.Anchor is W.Paragraph) yield return child;
            else if (child.Kind is "body" or "table" or "row" or "cell")
                foreach (var p in Paragraphs(child)) yield return p;
    }

    public static string Read(DocxDocument doc)
    {
        var body = doc.Main.Document!.Body!;
        var starts = body.Descendants<W.BookmarkStart>().Where(b => DocxMarks.Shown(b.Name?.Value) && (!DocxMarks.WholeParagraph(b) || b.Ancestors<W.TableCell>().Any())).ToList();
        if (starts.Count == 0) return "[]";
        var paths = Paragraphs(doc.Root).ToDictionary(n => (W.Paragraph)n.Anchor, n => {
            var p = (W.Paragraph)n.Anchor;
            return n.Parent?.Kind == "cell" ? (Path: n.Parent.Path, At: p.ElementsBefore().OfType<W.Paragraph>().Sum(x => DocxRuns.ParagraphText(x).Length + 1)) : (Path: n.Path, At: 0);
        });
        var ends = body.Descendants<W.BookmarkEnd>().GroupBy(e => e.Id?.Value ?? "").ToDictionary(g => g.Key, g => g.First());
        var list = new JsonArray();
        foreach (var b in starts)
        {
            if (!DocxMarks.Shown(b.Name?.Value) || DocxMarks.WholeParagraph(b) && !b.Ancestors<W.TableCell>().Any()) continue;
            if (!ends.TryGetValue(b.Id?.Value ?? "", out var e)) continue;
            var start = b.Ancestors<W.Paragraph>().FirstOrDefault(); var end = e.Ancestors<W.Paragraph>().FirstOrDefault();
            if (start is null || end is null || !paths.TryGetValue(start, out var sp) || !paths.TryGetValue(end, out var ep)) continue;
            list.Add((JsonNode)new JsonObject { ["name"] = b.Name!.Value, ["start"] = sp.Path, ["startOffset"] = sp.At + DocxFootnotes.OffsetOf(start, b), ["end"] = ep.Path, ["endOffset"] = ep.At + DocxFootnotes.OffsetOf(end, e) });
        }
        return list.ToJsonString();
    }

    public static void Write(DocxDocument doc, string json)
    {
        var patch = JsonNode.Parse(json)?.AsObject() ?? throw Invalid("Bookmark patch must be an object");
        var body = doc.Main.Document!.Body!;
        var paras = body.Descendants<W.Paragraph>().ToList();
        var plans = new List<(string Name, W.Paragraph Start, int From, W.Paragraph End, int To)>();
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in patch["set"]?.AsArray() ?? [])
        {
            var name = item?["name"]?.GetValue<string>() ?? "";
            if (!DocxMarks.BookmarkName().IsMatch(name) || !names.Add(name)) throw Invalid("Invalid or duplicate bookmark name");
            (W.Paragraph? Paragraph, int Offset) Endpoint(string path, int at)
            {
                var anchor = PathResolver.Single(doc.Root, path).Anchor;
                if (anchor is W.Paragraph paragraph) return (paragraph, at);
                if (anchor is W.TableCell cell) foreach (var p in cell.Elements<W.Paragraph>()) { var length = DocxRuns.ParagraphText(p).Length; if (at <= length) return (p, at); at -= length + 1; }
                return (null, at);
            }
            var (start, from) = Endpoint(item?["start"]?.GetValue<string>() ?? "", item?["startOffset"]?.GetValue<int>() ?? 0);
            var (end, to) = Endpoint(item?["end"]?.GetValue<string>() ?? "", item?["endOffset"]?.GetValue<int>() ?? 0);
            if (start is null || end is null || !paras.Contains(start) || !paras.Contains(end) || from < 0 || to < 0 || from > DocxRuns.ParagraphText(start).Length || to > DocxRuns.ParagraphText(end).Length || paras.IndexOf(start) > paras.IndexOf(end) || start == end && from > to)
                throw Invalid("Bookmark endpoints must follow document order and lie within paragraph text");
            plans.Add((name, start, from, end, to));
        }
        foreach (var item in patch["remove"]?.AsArray() ?? []) names.Add(item?.GetValue<string>() ?? "");
        // Validate the entire patch before touching native markers. Hidden Word anchors are never removed by this API.
        foreach (var b in body.Descendants<W.BookmarkStart>().Where(b => DocxMarks.Shown(b.Name?.Value) && names.Contains(b.Name!.Value!)).ToList())
        {
            foreach (var e in body.Descendants<W.BookmarkEnd>().Where(e => e.Id?.Value == b.Id?.Value).ToList()) e.Remove();
            b.Remove();
        }
        var id = body.Descendants<W.BookmarkStart>().Select(b => int.TryParse(b.Id?.Value, out var n) ? n : 0).DefaultIfEmpty().Max();
        foreach (var plan in plans)
        {
            var key = (++id).ToString(CultureInfo.InvariantCulture);
            var start = new W.BookmarkStart { Id = key, Name = plan.Name }; var end = new W.BookmarkEnd { Id = key };
            DocxFootnotes.Place(plan.Start, start, plan.From);
            if (plan.Start == plan.End && plan.From == plan.To) start.InsertAfterSelf(end);
            else DocxFootnotes.Place(plan.End, end, plan.To);
        }
    }

    static WriterException Invalid(string message) => new(ErrorCode.Validation, message, "Use valid paragraph paths and visible character offsets.");
}
