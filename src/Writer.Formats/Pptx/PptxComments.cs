using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

static class PptxComments
{
    public static string Read(PptxDocument doc, SlidePart slide) => NodeJson.Compact(w =>
    {
        var authors = doc.Presentation.CommentAuthorsPart?.CommentAuthorList?.Elements<P.CommentAuthor>().ToDictionary(a => a.Id?.Value ?? 0, a => a.Name?.Value ?? "Writer") ?? [];
        w.WriteStartArray();
        foreach (var c in slide.SlideCommentsPart?.CommentList?.Elements<P.Comment>() ?? [])
        {
            w.WriteStartObject(); w.WriteString("id", (c.AuthorId?.Value ?? 0) + ":" + (c.Index?.Value ?? 0));
            w.WriteString("author", authors.GetValueOrDefault(c.AuthorId?.Value ?? 0, "Writer")); w.WriteString("text", c.GetFirstChild<P.Text>()?.Text ?? "");
            if (c.DateTime?.Value is { } date) w.WriteString("date", date.ToString("O", CultureInfo.InvariantCulture));
            w.WriteEndObject();
        }
        foreach (var part in slide.GetPartsOfType<PowerPointCommentPart>())
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            var xml = XDocument.Load(stream);
            foreach (var comment in xml.Descendants().Where(e => e.Name.LocalName == "cm"))
            {
                w.WriteStartObject(); w.WriteString("id", "modern:" + (string?)comment.Attribute("id")); w.WriteString("author", (string?)comment.Attribute("authorId") ?? "Writer");
                w.WriteString("text", string.Concat(comment.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value))); w.WriteBoolean("readonly", true); w.WriteEndObject();
            }
        }
        w.WriteEndArray();
    });
    public static void Write(PptxDocument doc, SlidePart slide, string json)
    {
        using var data = JsonDocument.Parse(json);
        var list = data.RootElement.EnumerateArray().ToList();
        var part = slide.SlideCommentsPart ?? slide.AddNewPart<SlideCommentsPart>();
        var comments = part.CommentList ??= new P.CommentList();
        var old = comments.Elements<P.Comment>().ToDictionary(c => (c.AuthorId?.Value ?? 0) + ":" + (c.Index?.Value ?? 0));
        var authorsPart = doc.Presentation.CommentAuthorsPart ?? doc.Presentation.AddNewPart<CommentAuthorsPart>();
        var authors = authorsPart.CommentAuthorList ??= new P.CommentAuthorList();
        comments.RemoveAllChildren<P.Comment>();
        foreach (var row in list)
        {
            var id = row.TryGetProperty("id", out var ident) ? ident.ToString() : "";
            if (id.StartsWith("modern:", StringComparison.Ordinal)) continue;
            var name = row.TryGetProperty("author", out var authorValue) ? authorValue.GetString() ?? "Writer" : "Writer";
            var author = authors.Elements<P.CommentAuthor>().FirstOrDefault(a => a.Name?.Value == name);
            if (author is null) { var next = authors.Elements<P.CommentAuthor>().Select(a => a.Id?.Value ?? 0).DefaultIfEmpty().Max() + 1; author = new P.CommentAuthor { Id = next, Name = name, Initials = name[..Math.Min(2, name.Length)], LastIndex = 0, ColorIndex = next % 8 }; authors.Append(author); }
            var comment = old.TryGetValue(id, out var kept) ? kept : new P.Comment(new P.Position { X = 0, Y = 0 }, new P.Text()) { AuthorId = author.Id, Index = (author.LastIndex?.Value ?? 0) + 1, DateTime = DateTime.UtcNow };
            author.LastIndex = Math.Max(author.LastIndex?.Value ?? 0, comment.Index?.Value ?? 0);
            comment.GetFirstChild<P.Text>()!.Text = row.GetProperty("text").GetString() ?? "";
            comments.Append(comment);
        }
    }
}
