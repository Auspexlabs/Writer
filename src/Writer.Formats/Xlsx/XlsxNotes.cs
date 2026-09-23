using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

/// <summary>Cell hyperlinks and legacy comments (notes), stored the way Excel stores them:
/// hyperlinks as a &lt;hyperlinks&gt; entry plus a relationship, notes as a comments part plus a VML drawing.</summary>
static class XlsxNotes
{
    public static string? Link(WorksheetPart part, string reference)
    {
        var link = FindLink(part, reference);
        if (link is null) return null;
        if (link.Location?.Value is { Length: > 0 } location) return "#" + location;
        return link.Id?.Value is { } id ? part.HyperlinkRelationships.FirstOrDefault(r => r.Id == id)?.Uri.OriginalString : null;
    }

    /// <summary>A URL, or #Sheet!A1 for a place in the workbook. Empty removes the link.</summary>
    public static void SetLink(WorksheetPart part, string reference, string target)
    {
        RemoveLink(part, reference);
        if (target.Length == 0) return;
        var link = new Hyperlink { Reference = reference };
        if (target.StartsWith('#')) link.Location = target[1..];
        else
        {
            Uri uri;
            try { uri = new Uri(target, UriKind.RelativeOrAbsolute); }
            catch (UriFormatException) { throw new WriterException(ErrorCode.Validation, $"'{target}' is not a valid link", "Use a URL like https://example.com or #Sheet2!A1."); }
            link.Id = part.AddHyperlinkRelationship(uri, true).Id;
        }
        var links = part.Worksheet!.GetFirstChild<Hyperlinks>();
        if (links is null)
            InsertBeforeAny(part.Worksheet, links = new Hyperlinks(), e => e is PrintOptions or PageMargins or PageSetup or HeaderFooter or RowBreaks or ColumnBreaks
                or CustomProperties or CellWatches or IgnoredErrors or Drawing or LegacyDrawing or LegacyDrawingHeaderFooter or DrawingHeaderFooter
                or Picture or OleObjects or Controls or WebPublishItems or TableParts or WorksheetExtensionList or AlternateContent);
        links.Append(link);
    }

    public static void RemoveLink(WorksheetPart part, string reference)
    {
        if (FindLink(part, reference) is not { } link) return;
        var links = link.Parent!;
        if (link.Id?.Value is { } id && part.HyperlinkRelationships.FirstOrDefault(r => r.Id == id) is { } relationship
            && !links.Elements<Hyperlink>().Any(other => !ReferenceEquals(other, link) && other.Id?.Value == id))
            part.DeleteReferenceRelationship(relationship);
        link.Remove();
        if (!links.HasChildren) links.Remove();
    }

    static Hyperlink? FindLink(WorksheetPart part, string reference)
    {
        var (col, row) = XlsxCells.Parse(reference);
        return part.Worksheet?.GetFirstChild<Hyperlinks>()?.Elements<Hyperlink>().FirstOrDefault(h =>
            h.Reference?.Value is { } r && XlsxCells.ParseRange(r) is var box && col >= box.Col1 && col <= box.Col2 && row >= box.Row1 && row <= box.Row2);
    }

    public static string? Note(WorksheetPart part, string reference) => FindNote(part, reference)?.CommentText?.InnerText;

    /// <summary>Creates or updates the cell's comment; empty removes it. New comments get a VML shape so Excel shows them.</summary>
    public static void SetNote(WorksheetPart part, string reference, string text)
    {
        if (text.Length == 0)
        {
            RemoveNote(part, reference);
            return;
        }
        var body = new CommentText(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        if (FindNote(part, reference) is { } existing)
        {
            existing.CommentText = body;
            return;
        }
        var comments = (part.WorksheetCommentsPart ?? part.AddNewPart<WorksheetCommentsPart>()).Comments ??= new Comments(new Authors(), new CommentList());
        var authors = comments.Authors ??= comments.PrependChild(new Authors());
        if (!authors.HasChildren) authors.Append(new Author("Writer"));
        (comments.CommentList ??= comments.AppendChild(new CommentList())).Append(new Comment(body) { Reference = reference, AuthorId = 0U });
        var (col, row) = XlsxCells.Parse(reference);
        EditVml(part, create: true, xml =>
        {
            var next = xml.Descendants(V + "shape").Select(s => (s.Attribute("id")?.Value ?? "").Split('s')[^1])
                .Select(n => int.TryParse(n, NumberStyles.None, CultureInfo.InvariantCulture, out var i) ? i : 0).DefaultIfEmpty(1024).Max() + 1;
            xml.Add(Shape(next, col - 1, row - 1));
        });
    }

    public static void RemoveNote(WorksheetPart part, string reference)
    {
        if (FindNote(part, reference) is not { } comment) return;
        var list = comment.Parent!;
        comment.Remove();
        if (!list.HasChildren) part.DeletePart(part.WorksheetCommentsPart!);
        var (col, row) = XlsxCells.Parse(reference);
        EditVml(part, create: false, xml =>
        {
            foreach (var shape in xml.Elements(V + "shape").Where(s => s.Element(X + "ClientData") is { } data && data.Attribute("ObjectType")?.Value == "Note"
                && data.Element(X + "Row")?.Value.Trim() == (row - 1).ToString(CultureInfo.InvariantCulture)
                && data.Element(X + "Column")?.Value.Trim() == (col - 1).ToString(CultureInfo.InvariantCulture)).ToList())
                shape.Remove();
        });
    }

    static Comment? FindNote(WorksheetPart part, string reference) =>
        part.WorksheetCommentsPart?.Comments?.CommentList?.Elements<Comment>().FirstOrDefault(c => c.Reference?.Value == reference);

    static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
    static readonly XNamespace X = "urn:schemas-microsoft-com:office:excel";

    /// <summary>Edits the VML drawing behind &lt;legacyDrawing&gt;, creating it when asked; a drawing left without shapes is deleted.</summary>
    static void EditVml(WorksheetPart part, bool create, Action<XElement> edit)
    {
        var worksheet = part.Worksheet!;
        var legacy = worksheet.GetFirstChild<LegacyDrawing>();
        var vml = legacy?.Id?.Value is { } id && part.TryGetPartById(id, out var existing) ? existing as VmlDrawingPart : null;
        XElement xml;
        if (vml is null)
        {
            if (!create) return;
            vml = part.AddNewPart<VmlDrawingPart>();
            legacy?.Remove();
            InsertBeforeAny(worksheet, new LegacyDrawing { Id = part.GetIdOfPart(vml) },
                e => e is LegacyDrawingHeaderFooter or DrawingHeaderFooter or Picture or OleObjects or Controls or WebPublishItems or TableParts or WorksheetExtensionList or AlternateContent);
            xml = new XElement("xml", new XAttribute(XNamespace.Xmlns + "v", V), new XAttribute(XNamespace.Xmlns + "o", O), new XAttribute(XNamespace.Xmlns + "x", X),
                new XElement(O + "shapelayout", new XAttribute(V + "ext", "edit"), new XElement(O + "idmap", new XAttribute(V + "ext", "edit"), new XAttribute("data", "1"))),
                new XElement(V + "shapetype", new XAttribute("id", "_x0000_t202"), new XAttribute("coordsize", "21600,21600"), new XAttribute(O + "spt", "202"),
                    new XAttribute("path", "m,l,21600r21600,l21600,xe"),
                    new XElement(V + "stroke", new XAttribute("joinstyle", "miter")),
                    new XElement(V + "path", new XAttribute("gradientshapeok", "t"), new XAttribute(O + "connecttype", "rect"))));
        }
        else
        {
            using var reader = vml.GetStream(FileMode.Open, FileAccess.Read);
            xml = XElement.Load(reader);
        }
        edit(xml);
        if (!xml.Elements(V + "shape").Any())
        {
            part.DeletePart(vml);
            worksheet.GetFirstChild<LegacyDrawing>()?.Remove();
            return;
        }
        using var stream = vml.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>A hidden note box anchored beside the cell, as Excel draws it.</summary>
    static XElement Shape(int id, int col, int row) => new(V + "shape",
        new XAttribute("id", $"_x0000_s{id}"), new XAttribute("type", "#_x0000_t202"),
        new XAttribute("style", "position:absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:60pt;z-index:1;visibility:hidden"),
        new XAttribute("fillcolor", "#ffffe1"), new XAttribute(O + "insetmode", "auto"),
        new XElement(V + "fill", new XAttribute("color2", "#ffffe1")),
        new XElement(V + "shadow", new XAttribute("on", "t"), new XAttribute("color", "black"), new XAttribute("obscured", "t")),
        new XElement(V + "path", new XAttribute(O + "connecttype", "none")),
        new XElement(V + "textbox", new XAttribute("style", "mso-direction-alt:auto"), new XElement("div", new XAttribute("style", "text-align:left"), "")),
        new XElement(X + "ClientData", new XAttribute("ObjectType", "Note"),
            new XElement(X + "MoveWithCells"), new XElement(X + "SizeWithCells"),
            new XElement(X + "Anchor", $"{col + 1}, 15, {Math.Max(row - 1, 0)}, 2, {col + 3}, 15, {Math.Max(row - 1, 0) + 4}, 8"),
            new XElement(X + "AutoFill", "False"),
            new XElement(X + "Row", row), new XElement(X + "Column", col)));

    static void InsertBeforeAny(OpenXmlCompositeElement parent, OpenXmlElement element, Func<OpenXmlElement, bool> after)
    {
        var anchor = parent.ChildElements.FirstOrDefault(after);
        if (anchor is null) parent.Append(element);
        else parent.InsertBefore(element, anchor);
    }
}
