using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>A footnote or endnote, projected under the paragraph holding its reference mark. The mark is a run with
/// w:footnoteReference (w:endnoteReference) at a character offset; the text lives in word/footnotes.xml (endnotes.xml) as
/// paragraphs that open with the note's own mark (w:footnoteRef). Word numbers notes by their order in the text.</summary>
sealed class DocxFootnote(DocxDocument doc, W.Paragraph p, OpenXmlElement reference) : Node
{
    public override string Kind => "footnote";
    public override object Anchor => reference;
    bool Endnote => reference is W.EndnoteReference;
    string Id => ((reference as W.FootnoteReference)?.Id ?? ((W.EndnoteReference)reference).Id)?.Value.ToString(CultureInfo.InvariantCulture) ?? "";

    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string>
    {
        ["id"] = Endnote ? "e" + Id : Id, // footnotes and endnotes count from 1 each: an endnote's id says which it is
        ["kind"] = Endnote ? "endnote" : "footnote",
        ["text"] = DocxFootnotes.Note(doc, Endnote, Id) is { } note ? DocxFootnotes.TextOf(note) : "",
        ["at"] = DocxFootnotes.OffsetOf(p, reference.Parent!).ToString(CultureInfo.InvariantCulture),
    };

    public override string GetRaw() => DocxFootnotes.Note(doc, Endnote, Id)?.OuterXml ?? reference.Parent!.OuterXml;

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "text": DocxFootnotes.SetText(doc, Endnote, Id, value); break;
            case "at":
                var run = (W.Run)reference.Parent!;
                var (previous, next) = (run.PreviousSibling(), run.NextSibling());
                run.Remove();
                DocxRuns.Rejoin(previous, next);
                DocxFootnotes.Place(p, run, int.Parse(value, CultureInfo.InvariantCulture));
                break;
        }
    }

    public override void Remove() => DocxFootnotes.Remove(doc, Endnote, Id, reference);
}

static class DocxFootnotes
{
    /// <summary>The notes whose marks sit in the paragraph, in reading order.</summary>
    public static IEnumerable<Node> In(DocxDocument doc, W.Paragraph p) =>
        p.Descendants().Where(e => e is W.FootnoteReference or W.EndnoteReference).Select(e => (Node)new DocxFootnote(doc, p, e));

    public static Node Add(DocxDocument doc, W.Paragraph p, IReadOnlyDictionary<string, string> props)
    {
        var endnote = props.GetValueOrDefault("kind") == "endnote";
        var id = NextId(doc, endnote);
        var text = props.GetValueOrDefault("text") ?? "";
        var run = new W.Run(new W.RunProperties(new W.RunStyle { Val = doc.Styles.ResolveStyle(endnote ? "EndnoteReference" : "FootnoteReference", "character") }),
            endnote ? new W.EndnoteReference { Id = id } : new W.FootnoteReference { Id = id });
        Place(p, run, props.TryGetValue("at", out var at) ? int.Parse(at, CultureInfo.InvariantCulture) : int.MaxValue);
        var note = endnote ? new W.Endnote { Id = id } : (OpenXmlElement)new W.Footnote { Id = id };
        Notes(doc, endnote, create: true)!.Append(note);
        SetText(doc, endnote, id.ToString(CultureInfo.InvariantCulture), text);
        return new DocxFootnote(doc, p, run.ChildElements.Last());
    }

    /// <summary>Puts a mark's run (or an equation) at a character offset of the paragraph's visible text, splitting a run when it falls inside one.</summary>
    public static void Place(W.Paragraph p, OpenXmlElement run, int offset)
    {
        var pos = 0;
        foreach (var (r, _) in DocxRuns.Walk(p).ToList())
        {
            var length = DocxRuns.RunText(r).Length;
            if (offset <= pos) { r.InsertBeforeSelf(run); return; }
            if (offset < pos + length) { DocxReplace.SplitRun(r, offset - pos); r.InsertAfterSelf(run); return; }
            pos += length;
        }
        p.Append(run);
    }

    /// <summary>The visible characters before an element of the paragraph (a mark's run, an equation).</summary>
    public static int OffsetOf(W.Paragraph p, OpenXmlElement holder)
    {
        var pos = 0;
        foreach (var (r, _) in DocxRuns.Walk(p))
        {
            if (ReferenceEquals(holder, r) || holder.IsBefore(r)) return pos;
            pos += DocxRuns.RunText(r).Length;
        }
        return pos;
    }

    static long NextId(DocxDocument doc, bool endnote) =>
        1 + (Notes(doc, endnote, create: false)?.ChildElements.Select(n => (n as W.Footnote)?.Id?.Value ?? (n as W.Endnote)?.Id?.Value ?? 0).DefaultIfEmpty(0).Max() ?? 0);

    /// <summary>The notes root (w:footnotes or w:endnotes), made with Word's separator notes when the document has none yet.</summary>
    static OpenXmlCompositeElement? Notes(DocxDocument doc, bool endnote, bool create)
    {
        if (endnote)
        {
            var part = doc.Main.EndnotesPart ?? (create ? doc.Main.AddNewPart<EndnotesPart>() : null);
            if (part is null) return null;
            if (part.Endnotes is null)
                part.Endnotes = new W.Endnotes(
                    new W.Endnote(new W.Paragraph(new W.Run(new W.SeparatorMark()))) { Type = W.FootnoteEndnoteValues.Separator, Id = -1 },
                    new W.Endnote(new W.Paragraph(new W.Run(new W.ContinuationSeparatorMark()))) { Type = W.FootnoteEndnoteValues.ContinuationSeparator, Id = 0 });
            return part.Endnotes;
        }
        var footnotes = doc.Main.FootnotesPart ?? (create ? doc.Main.AddNewPart<FootnotesPart>() : null);
        if (footnotes is null) return null;
        if (footnotes.Footnotes is null)
            footnotes.Footnotes = new W.Footnotes(
                new W.Footnote(new W.Paragraph(new W.Run(new W.SeparatorMark()))) { Type = W.FootnoteEndnoteValues.Separator, Id = -1 },
                new W.Footnote(new W.Paragraph(new W.Run(new W.ContinuationSeparatorMark()))) { Type = W.FootnoteEndnoteValues.ContinuationSeparator, Id = 0 });
        return footnotes.Footnotes;
    }

    public static OpenXmlElement? Note(DocxDocument doc, bool endnote, string id) =>
        Notes(doc, endnote, create: false)?.ChildElements.FirstOrDefault(n => ((n as W.Footnote)?.Id?.Value ?? (n as W.Endnote)?.Id?.Value)?.ToString(CultureInfo.InvariantCulture) == id);

    public static string TextOf(OpenXmlElement note) =>
        string.Join("\n", note.Elements<W.Paragraph>().Select(DocxRuns.ParagraphText)).TrimStart(' ');

    /// <summary>The note's paragraphs anew: the first opens with the note's own mark, then the text, a space between as Word writes it.</summary>
    public static void SetText(DocxDocument doc, bool endnote, string id, string text)
    {
        var note = Note(doc, endnote, id) ?? throw new WriterException(ErrorCode.Validation, $"No {(endnote ? "endnote" : "footnote")} {id}", "Add the note first.");
        note.RemoveAllChildren<W.Paragraph>();
        var lines = text.ReplaceLineEndings("\n").Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            var p = new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = doc.Styles.ResolveStyle(endnote ? "EndnoteText" : "FootnoteText", "paragraph") }));
            if (i == 0) p.Append(new W.Run(new W.RunProperties(new W.RunStyle { Val = doc.Styles.ResolveStyle(endnote ? "EndnoteReference" : "FootnoteReference", "character") }),
                endnote ? new W.EndnoteReferenceMark() : new W.FootnoteReferenceMark()));
            var line = i == 0 && lines[i].Length > 0 ? " " + lines[i] : lines[i];
            if (line.Length > 0) p.Append(new W.Run(DocxRuns.TextElements(line)));
            note.Append(p);
        }
    }

    public static void Remove(DocxDocument doc, bool endnote, string id, OpenXmlElement reference)
    {
        var run = reference.Parent as W.Run;
        reference.Remove();
        if (run is not null && run.ChildElements.All(c => c is W.RunProperties))
        {
            var (previous, next) = (run.PreviousSibling(), run.NextSibling());
            run.Remove();
            DocxRuns.Rejoin(previous, next);
        }
        Note(doc, endnote, id)?.Remove();
    }
}
