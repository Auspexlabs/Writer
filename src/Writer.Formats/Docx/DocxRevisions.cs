using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>Tracked changes: the w:ins / w:del wrappers around runs, the trackRevisions setting, the default author, and accepting
/// or rejecting them (DocxReplace records edits as revisions). Format changes (rPrChange, pPrChange), moves and table-row markers
/// are left as they are and not counted.</summary>
static class DocxRevisions
{
    /// <summary>The default author lives in a document variable in settings.xml: Word keeps it and never shows it. The file's
    /// creator is deliberately not used, or a reviewer's changes would carry the original author's name.</summary>
    const string AuthorVariable = "WriterAuthor";

    public static string? DefaultAuthor(DocxDocument doc) =>
        doc.Main.DocumentSettingsPart?.Settings?.GetFirstChild<W.DocumentVariables>()?.Elements<W.DocumentVariable>()
            .FirstOrDefault(v => v.Name?.Value == AuthorVariable)?.Val?.Value is { Length: > 0 } a ? a : null;

    public static string Author(DocxDocument doc) => DefaultAuthor(doc) ?? "Writer";

    public static void SetDefaultAuthor(DocxDocument doc, string author)
    {
        var settings = Settings(doc);
        var variables = settings.GetFirstChild<W.DocumentVariables>();
        variables?.Elements<W.DocumentVariable>().Where(v => v.Name?.Value == AuthorVariable).ToList().ForEach(v => v.Remove());
        if (author.Length > 0)
        {
            if (variables is null) settings.AddChild(variables = new W.DocumentVariables());
            variables.Append(new W.DocumentVariable { Name = AuthorVariable, Val = author });
        }
        if (variables is { HasChildren: false }) variables.Remove();
    }

    static W.Settings Settings(DocxDocument doc) => (doc.Main.DocumentSettingsPart ?? doc.Main.AddNewPart<DocumentSettingsPart>()).Settings ??= new W.Settings();

    public static string Now() => DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

    /// <summary>The run (or hyperlink) inside a w:ins or w:del carrying the spec's author and date.</summary>
    public static OpenXmlElement Wrap(DocxDocument doc, RunSpec spec, OpenXmlElement element)
    {
        var author = spec.Author ?? Author(doc);
        var date = new DateTimeValue { InnerText = spec.Date ?? Now() };
        return spec.Deleted
            ? new W.DeletedRun(element) { Author = author, Date = date, Id = doc.NextId() }
            : new W.InsertedRun(element) { Author = author, Date = date, Id = doc.NextId() };
    }

    public static bool Tracking(DocxDocument doc) =>
        doc.Main.DocumentSettingsPart?.Settings?.GetFirstChild<W.TrackRevisions>() is { } t && (t.Val is null || t.Val.Value);

    public static void SetTracking(DocxDocument doc, bool on)
    {
        var settings = Settings(doc);
        settings.GetFirstChild<W.TrackRevisions>()?.Remove();
        if (on) settings.AddChild(new W.TrackRevisions());
    }

    /// <summary>Marks a paragraph added while tracking as inserted, mark included, so rejecting it removes the whole paragraph.</summary>
    public static void MarkInserted(DocxDocument doc, W.Paragraph p)
    {
        var pp = p.ParagraphProperties ??= new W.ParagraphProperties();
        var mark = pp.ParagraphMarkRunProperties ??= new W.ParagraphMarkRunProperties();
        mark.AddChild(new W.Inserted { Author = Author(doc), Date = new DateTimeValue { InnerText = Now() }, Id = doc.NextId() });
    }

    static bool IsMark(OpenXmlElement e) => e is W.Inserted or W.Deleted && e.Parent is W.ParagraphMarkRunProperties;

    /// <summary>Pending insertions and deletions of text and paragraph marks in the body.</summary>
    public static int Count(DocxDocument doc) =>
        doc.Main.Document!.Body!.Descendants().Count(e => e is W.InsertedRun or W.DeletedRun || IsMark(e));

    /// <summary>Accepts or rejects every revision in the paragraphs of a container (a paragraph itself, the body, a cell).</summary>
    public static void Resolve(OpenXmlElement container, bool accept)
    {
        var paragraphs = container is W.Paragraph p ? [p] : container.Descendants<W.Paragraph>().ToList();
        foreach (var paragraph in paragraphs) ResolveParagraph(paragraph, accept);
    }

    static void ResolveParagraph(W.Paragraph p, bool accept)
    {
        foreach (var ins in p.Descendants<W.InsertedRun>().ToList())
        {
            if (ins.Parent is null) continue;
            if (accept) Unwrap(ins);
            else ins.Remove();
        }
        foreach (var del in p.Descendants<W.DeletedRun>().ToList())
        {
            if (del.Parent is null) continue;
            if (accept) { del.Remove(); continue; }
            foreach (var t in del.Descendants<W.DeletedText>().ToList())
                t.Parent!.ReplaceChild(new W.Text(t.Text) { Space = SpaceProcessingModeValues.Preserve }, t);
            Unwrap(del);
        }
        var mark = p.ParagraphProperties?.ParagraphMarkRunProperties;
        var inserted = mark?.GetFirstChild<W.Inserted>();
        var deleted = mark?.GetFirstChild<W.Deleted>();
        inserted?.Remove();
        deleted?.Remove();
        var gone = accept ? deleted is not null : inserted is not null;
        if (!gone || p.Parent is not { } parent || !parent.Elements<W.Paragraph>().Skip(1).Any()) return;
        // The paragraph mark goes: an emptied paragraph disappears, remaining content joins the next paragraph.
        if (p.NextSibling() is W.Paragraph next)
        {
            var cursor = (OpenXmlElement?)next.ParagraphProperties;
            foreach (var child in p.ChildElements.Where(c => c is not W.ParagraphProperties).ToList())
            {
                child.Remove();
                if (cursor is null) next.PrependChild(child);
                else next.InsertAfter(child, cursor);
                cursor = child;
            }
        }
        else if (p.ChildElements.Any(c => c is not W.ParagraphProperties)) return;
        DocxBlocks.Detach(p);
    }

    static void Unwrap(OpenXmlElement wrapper)
    {
        foreach (var child in wrapper.ChildElements.ToList())
        {
            child.Remove();
            wrapper.Parent!.InsertBefore(child, wrapper);
        }
        wrapper.Remove();
    }
}
