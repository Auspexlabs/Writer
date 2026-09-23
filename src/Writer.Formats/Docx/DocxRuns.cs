using System.Text;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>Reading and writing the text of paragraphs. Insertions count as text, deletions do not, though deleted runs
/// are still projected (with change=deleted) so a reader can see what a tracked change removed.</summary>
static class DocxRuns
{
    /// <summary>Text-bearing runs in reading order, each with the hyperlink that wraps it, if any. Deleted runs only when asked for.</summary>
    public static IEnumerable<(W.Run Run, W.Hyperlink? Link)> Walk(OpenXmlElement container, W.Hyperlink? link = null, bool deleted = false)
    {
        foreach (var child in container.ChildElements)
        {
            switch (child)
            {
                case W.Run run:
                    if (HasText(run)) yield return (run, link);
                    break;
                case W.Hyperlink h:
                    foreach (var x in Walk(h, h, deleted)) yield return x;
                    break;
                case W.DeletedRun when deleted:
                case W.InsertedRun or W.MoveToRun or W.SimpleField or W.SdtRun or W.SdtContentRun or W.CustomXmlRun:
                    foreach (var x in Walk(child, link, deleted)) yield return x;
                    break;
            }
        }
    }

    public static bool HasText(W.Run run) => run.ChildElements.Any(IsTextElement);

    public static bool IsTextElement(OpenXmlElement e) => e is W.Text or W.DeletedText or W.TabChar or W.Break or W.CarriageReturn or W.NoBreakHyphen;

    public static string RunText(W.Run run)
    {
        var sb = new StringBuilder();
        foreach (var e in run.ChildElements)
        {
            switch (e)
            {
                case W.Text t: sb.Append(t.Text); break;
                case W.DeletedText d: sb.Append(d.Text); break;
                case W.TabChar: sb.Append('\t'); break;
                case W.Break b when b.Type?.Value == W.BreakValues.Page: sb.Append('\f'); break; // a page break, as Word's own text has it
                case W.Break or W.CarriageReturn: sb.Append('\n'); break;
                case W.NoBreakHyphen: sb.Append('-'); break;
            }
        }
        return sb.ToString();
    }

    public static string ParagraphText(W.Paragraph p) => string.Concat(Walk(p).Select(x => RunText(x.Run)));

    /// <summary>The w:ins or w:del wrapping a run, if any.</summary>
    public static OpenXmlElement? Revision(W.Run run)
    {
        for (var e = run.Parent; e is not null and not W.Paragraph; e = e.Parent)
            if (e is W.InsertedRun or W.DeletedRun) return e;
        return null;
    }

    /// <summary>Elements for a string: w:t for text (w:delText inside a deletion), w:tab for tabs, w:br for newlines and a page
    /// break (w:br w:type="page") for a form feed.</summary>
    public static IEnumerable<OpenXmlElement> TextElements(string text, bool deleted = false)
    {
        var pages = text.Split('\f'); // before ReplaceLineEndings, which would take a form feed for a newline
        for (var k = 0; k < pages.Length; k++)
        {
            if (k > 0) yield return new W.Break { Type = W.BreakValues.Page };
            var lines = pages[k].ReplaceLineEndings("\n").Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                if (i > 0) yield return new W.Break();
                var parts = lines[i].Split('\t');
                for (var j = 0; j < parts.Length; j++)
                {
                    if (j > 0) yield return new W.TabChar();
                    if (parts[j].Length == 0) continue;
                    yield return deleted
                        ? new W.DeletedText(parts[j]) { Space = SpaceProcessingModeValues.Preserve }
                        : new W.Text(parts[j]) { Space = SpaceProcessingModeValues.Preserve };
                }
            }
        }
    }

    /// <summary>Replaces a run's text, keeping its formatting.</summary>
    public static void SetRunText(W.Run run, string text)
    {
        foreach (var e in run.ChildElements.Where(IsTextElement).ToList()) e.Remove();
        foreach (var e in TextElements(text, Revision(run) is W.DeletedRun)) run.Append(e);
    }

    /// <summary>Things that survive a text replacement: properties, bookmarks, comment and permission ranges, reference-only runs.</summary>
    static bool IsMarker(OpenXmlElement e) => e switch
    {
        W.ParagraphProperties or W.BookmarkStart or W.BookmarkEnd or W.CommentRangeStart or W.CommentRangeEnd
            or W.PermStart or W.PermEnd or W.ProofError => true,
        W.Run r => r.ChildElements.All(c => c is W.RunProperties or W.CommentReference or W.FootnoteReference or W.EndnoteReference
            or W.AnnotationReferenceMark or W.FootnoteReferenceMark or W.EndnoteReferenceMark or W.LastRenderedPageBreak),
        _ => false,
    };

    /// <summary>Replaces the paragraph's content with new runs at the position of the old content. Table cells still write through
    /// this; paragraphs use DocxReplace, which keeps fields, equations and pictures in place.</summary>
    public static void ReplaceContent(W.Paragraph p, IEnumerable<OpenXmlElement> content)
    {
        var first = p.ChildElements.FirstOrDefault(c => !IsMarker(c));
        var cursor = first?.PreviousSibling() ?? (OpenXmlElement?)p.ParagraphProperties;
        foreach (var child in p.ChildElements.Where(c => !IsMarker(c)).ToList()) child.Remove();
        foreach (var element in content)
        {
            if (cursor is null) p.PrependChild(element);
            else p.InsertAfter(element, cursor);
            cursor = element;
        }
    }

    /// <summary>The formatting of the first plain (non-link) run, cloned, so replacement text keeps the paragraph's look.
    /// A tracked format change on that run is not copied: it belongs to the old text, not to every new run.</summary>
    public static W.RunProperties? BaseRunProperties(W.Paragraph p)
    {
        var rp = Walk(p).Where(x => x.Link is null).Select(x => x.Run).FirstOrDefault()?.RunProperties?.CloneNode(true) as W.RunProperties;
        rp?.RunPropertiesChange?.Remove();
        return rp;
    }

    /// <summary>Runs for the specs; a spec with a change becomes a run wrapped in w:ins or w:del, attributed to its author
    /// (else the document's, else Writer) at its date (else now).</summary>
    public static List<OpenXmlElement> MakeRuns(DocxDocument doc, IEnumerable<RunSpec> specs, W.RunProperties? baseProperties)
    {
        var result = new List<OpenXmlElement>();
        foreach (var spec in specs)
        {
            var run = new W.Run();
            var rp = baseProperties?.CloneNode(true) as W.RunProperties ?? new W.RunProperties();
            if (spec.Bold) rp.Bold = new W.Bold();
            if (spec.Italic) rp.Italic = new W.Italic();
            if (spec.Strike) rp.Strike = new W.Strike();
            if (spec.Code) rp.RunFonts = new W.RunFonts { Ascii = "Consolas", HighAnsi = "Consolas", ComplexScript = "Consolas" };
            if (spec.Underline) rp.Underline = new W.Underline { Val = W.UnderlineValues.Single };
            if (spec.Color is not null) rp.Color = new W.Color { Val = spec.Color };
            if (spec.HalfPoints is { } half)
            {
                rp.FontSize = new W.FontSize { Val = half };
                rp.FontSizeComplexScript = new W.FontSizeComplexScript { Val = half };
            }
            if (spec.Font is not null) rp.RunFonts = new W.RunFonts { Ascii = spec.Font, HighAnsi = spec.Font, EastAsia = spec.Font, ComplexScript = spec.Font };
            if (spec.Highlight is not null) rp.Shading = new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = spec.Highlight };
            if (spec.Link is not null) rp.RunStyle = new W.RunStyle { Val = doc.Styles.ResolveStyle("Hyperlink", "character") };
            if (rp.HasChildren) run.RunProperties = rp;
            foreach (var e in TextElements(spec.Text, spec.Deleted)) run.Append(e);
            var element = spec.Link is null ? run : (OpenXmlElement)MakeHyperlink(doc, spec.Link, run);
            result.Add(spec.Change is "inserted" or "deleted" ? DocxRevisions.Wrap(doc, spec, element) : element);
        }
        return result;
    }

    public static W.Hyperlink MakeHyperlink(DocxDocument doc, string target, W.Run run)
    {
        if (target.StartsWith('#')) return new W.Hyperlink(run) { Anchor = target[1..], History = true };
        Uri uri;
        try { uri = new Uri(target, UriKind.RelativeOrAbsolute); }
        catch (UriFormatException) { throw new WriterException(ErrorCode.Validation, $"'{target}' is not a valid link", "Use a URL like https://example.com or #bookmark."); }
        var relationship = doc.Main.AddHyperlinkRelationship(uri, true);
        return new W.Hyperlink(run) { Id = relationship.Id, History = true };
    }
}
