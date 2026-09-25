using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>What Word's 插入 and 引用 tabs put on a paragraph: a bookmark around it (where cross-references and #name links go), the
/// numbering of a caption (a SEQ field: 图 1, 表 2), and a drop cap (the paragraph is Word's drop-cap frame).</summary>
static partial class DocxMarks
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    // ---- bookmarks: hidden ones (_Toc, _GoBack) are Word's own and stay out of sight; _Ref ones are cross-reference targets ----

    static bool Shown(string? name) => name is { Length: > 0 } && (!name.StartsWith('_') || name.StartsWith("_Ref", StringComparison.Ordinal));

    public static string? Bookmark(W.Paragraph p) => p.Elements<W.BookmarkStart>().Select(b => b.Name?.Value).FirstOrDefault(Shown);

    /// <summary>A bookmark around the whole paragraph (none removes the paragraph's own); a name the document has elsewhere moves here.</summary>
    public static void SetBookmark(DocxDocument doc, W.Paragraph p, string name)
    {
        var body = doc.Main.Document!.Body!;
        var gone = p.Elements<W.BookmarkStart>().Where(b => Shown(b.Name?.Value)).ToList();
        if (name is not ("none" or ""))
        {
            if (!BookmarkName().IsMatch(name))
                throw new WriterException(ErrorCode.Validation, $"'{name}' is not a bookmark name", "A letter first, then letters, digits or _, up to 40 characters.");
            gone.AddRange(body.Descendants<W.BookmarkStart>().Where(b => b.Name?.Value == name));
        }
        foreach (var start in gone.Distinct().ToList())
        {
            var id = start.Id?.Value;
            foreach (var end in body.Descendants<W.BookmarkEnd>().Where(e => e.Id?.Value == id).ToList()) end.Remove();
            start.Remove();
        }
        if (name is "none" or "") return;
        var next = (1 + body.Descendants<W.BookmarkStart>().Select(b => int.TryParse(b.Id?.Value, NumberStyles.Integer, Inv, out var n) ? n : 0).DefaultIfEmpty(0).Max()).ToString(Inv);
        var bookmark = new W.BookmarkStart { Id = next, Name = name };
        if (p.ParagraphProperties is { } pp) pp.InsertAfterSelf(bookmark); else p.PrependChild(bookmark);
        p.Append(new W.BookmarkEnd { Id = next });
    }

    [GeneratedRegex(@"^\p{L}[\p{L}\p{N}_]{0,39}$|^_Ref\d{1,20}$")]
    private static partial Regex BookmarkName();

    // ---- captions: the label, then a SEQ field counting the captions with that label ----

    static readonly Regex Seq = new(@"^\s*SEQ\s+(\S+)", RegexOptions.IgnoreCase);

    /// <summary>The label of the paragraph's SEQ field (图, 表, Figure…), if it has one.</summary>
    public static string? Caption(W.Paragraph p) =>
        p.Descendants<W.SimpleField>().Select(f => f.Instruction?.Value).Concat(p.Descendants<W.FieldCode>().Select(c => c.Text))
            .Select(i => i is null ? null : Seq.Match(i)).FirstOrDefault(m => m is { Success: true })?.Groups[1].Value;

    /// <summary>Makes the paragraph a caption with this label: the Caption style and, in front, the label and a SEQ field whose result
    /// is the caption's number among those with the label before it (Word updates it on printing). An existing SEQ field just takes the
    /// new label; none takes the field out, keeping its number as text.</summary>
    public static void SetCaption(DocxDocument doc, W.Paragraph p, string label)
    {
        var simple = p.Descendants<W.SimpleField>().FirstOrDefault(f => Seq.IsMatch(f.Instruction?.Value ?? ""));
        if (label is "none" or "")
        {
            if (simple is not null)
            {
                foreach (var run in simple.Elements<W.Run>().ToList()) { run.Remove(); simple.InsertBeforeSelf(run); }
                simple.Remove();
            }
            foreach (var code in p.Descendants<W.FieldCode>().Where(c => Seq.IsMatch(c.Text)).ToList()) code.Parent?.Remove(); // a complex field keeps its result runs
            return;
        }
        if (label.Any(char.IsWhiteSpace)) throw new WriterException(ErrorCode.Validation, $"'{label}' is not a caption label", "One word, e.g. 图, 表, Figure, Table, Equation.");
        var instruction = $" SEQ {label} \\* ARABIC ";
        var pp = p.ParagraphProperties ??= new W.ParagraphProperties();
        if (pp.ParagraphStyleId is null || pp.ParagraphStyleId.Val?.Value == "Normal")
            pp.ParagraphStyleId = new W.ParagraphStyleId { Val = doc.Styles.ResolveStyle("Caption", "paragraph") };
        if (simple is not null) { simple.Instruction = instruction; return; }
        if (p.Descendants<W.FieldCode>().FirstOrDefault(c => Seq.IsMatch(c.Text)) is { } complex) { complex.Text = instruction; return; }
        var number = 1 + doc.Main.Document!.Body!.Descendants<W.Paragraph>().TakeWhile(x => x != p).Count(x => Caption(x) == label);
        var field = new W.SimpleField(new W.Run(new W.Text(number.ToString(Inv)))) { Instruction = instruction };
        var labelRun = new W.Run(new W.Text(label + " ") { Space = SpaceProcessingModeValues.Preserve });
        var first = p.ChildElements.FirstOrDefault(c => c is not (W.ParagraphProperties or W.BookmarkStart));
        if (first is null) { p.Append(labelRun); p.Append(field); }
        else { first.InsertBeforeSelf(labelRun); first.InsertBeforeSelf(field); }
    }

    // ---- drop caps: Word's frame of three lines with the text around it ----

    public static string? DropCap(W.Paragraph p) => p.ParagraphProperties?.FrameProperties?.DropCap?.Value is { } v && v != W.DropCapLocationValues.None
        ? v == W.DropCapLocationValues.Margin ? "margin" : "drop" : null;

    /// <summary>drop (in the text) or margin (beside it) makes the paragraph a drop cap three lines high, its letters sized to match; none
    /// makes it an ordinary paragraph again.</summary>
    public static void SetDropCap(W.Paragraph p, string value)
    {
        var pp = p.ParagraphProperties ??= new W.ParagraphProperties();
        if (value is "none" or "")
        {
            pp.FrameProperties = null;
            if (!pp.HasChildren) pp.Remove();
            return;
        }
        pp.FrameProperties = new W.FrameProperties
        {
            DropCap = value == "margin" ? W.DropCapLocationValues.Margin : W.DropCapLocationValues.Drop,
            Lines = 3, Wrap = W.TextWrappingValues.Around, VerticalPosition = W.VerticalAnchorValues.Text,
            HorizontalPosition = value == "margin" ? W.HorizontalAnchorValues.Page : W.HorizontalAnchorValues.Text,
        };
        pp.SpacingBetweenLines = new W.SpacingBetweenLines { Line = "780", LineRule = W.LineSpacingRuleValues.Exact, Before = "0", After = "0" };
        foreach (var (run, _) in DocxRuns.Walk(p).ToList())
        {
            var rp = run.RunProperties ??= new W.RunProperties();
            rp.Position = new W.Position { Val = "-6" };
            rp.FontSize = new W.FontSize { Val = "78" }; // 39pt: three lines of 11pt text
            rp.FontSizeComplexScript = new W.FontSizeComplexScript { Val = "78" };
        }
    }
}
