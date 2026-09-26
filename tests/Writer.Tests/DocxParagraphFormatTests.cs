using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Word's 段落 dialog on a paragraph: line spacing, spacing before and after (points or lines), indents in lengths or
/// characters, borders, keep with next, widow control, tab stops and distributed alignment, written and read back in the same words.</summary>
public class DocxParagraphFormatTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    [Fact]
    public void Paragraph_format_round_trips_and_none_removes_it()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "正文"), ("lineSpacing", "1.5"), ("spaceBefore", "6pt"), ("spaceAfter", "8pt"), ("indentLeft", "1cm"), ("indentRight", "0.5cm"),
            ("indentFirst", "2ch"), ("border", "bottom"), ("keepNext", "true"), ("keepLines", "true"), ("tabs", "left 2cm, center 8cm, right 15cm, decimal 10cm"), ("align", "distribute")), null);
        var got = p.GetProps();
        Assert.Equal(("1.5", "6pt", "8pt", "1cm", "0.5cm", "2ch", "bottom", "true", "true", "distribute"),
            (got["lineSpacing"], got["spaceBefore"], got["spaceAfter"], got["indentLeft"], got["indentRight"], got["indentFirst"], got["border"], got["keepNext"], got["keepLines"], got["align"]));
        Assert.Equal("left 2cm, center 8cm, right 15cm, decimal 10cm", got["tabs"]);
        var raw = XElement.Parse(p.GetRaw());
        Assert.Contains(raw.Descendants().Where(e => e.Name.LocalName == "ind"), e => e.Attributes().Any(a => a.Name.LocalName == "firstLineChars" && a.Value == "200"));
        Assert.Contains(raw.Descendants().Where(e => e.Name.LocalName == "spacing"), e => e.Attributes().Any(a => a.Name.LocalName == "line" && a.Value == "360"));
        Assert.Contains(raw.Descendants(), e => e.Name.LocalName == "jc" && e.Attributes().Any(a => a.Value == "distribute"));

        p = Mutations.Set(p, Props(("indentFirst", "-0.74cm"), ("lineSpacing", "18pt"), ("border", "box"), ("spaceBefore", "none")));
        got = p.GetProps();
        Assert.Equal(("-0.74cm", "18pt", "box"), (got["indentFirst"], got["lineSpacing"], got["border"]));
        Assert.False(got.ContainsKey("spaceBefore"));
        Assert.Equal("8pt", got["spaceAfter"]);
        Assert.Contains(XElement.Parse(p.GetRaw()).Descendants().Where(e => e.Name.LocalName == "ind"), e => e.Attributes().Any(a => a.Name.LocalName == "hanging"));

        p = Mutations.Set(p, Props(("lineSpacing", "min 20pt"), ("border", "none"), ("tabs", "none"), ("keepNext", "false"), ("indentFirst", "none"), ("indentLeft", "none"), ("indentRight", "none")));
        got = p.GetProps();
        Assert.Equal("min 20pt", got["lineSpacing"]);
        foreach (var name in new[] { "border", "tabs", "keepNext", "indentFirst", "indentLeft", "indentRight" }) Assert.False(got.ContainsKey(name), name);
        Assert.Equal("true", got["keepLines"]);

        var heading = Mutations.Add(body, "heading", Props(("text", "标题"), ("spaceBefore", "24pt"), ("keepNext", "true")), null);
        Assert.Equal("24pt", heading.GetProps()["spaceBefore"]);
        Assert.Throws<WriterException>(() => Mutations.Set(heading, Props(("border", "sideways"))));
        Assert.Throws<WriterException>(() => Mutations.Set(heading, Props(("tabs", "somewhere 2cm"))));
    }

    [Fact]
    public void Spacing_in_lines_is_written_as_word_counts_it_and_widow_control_turns_off_and_back()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "正文"), ("spaceBefore", "0.5lines"), ("spaceAfter", "1 行"), ("widowControl", "false")), null);
        var got = p.GetProps();
        Assert.Equal(("0.5lines", "1lines", "false"), (got["spaceBefore"], got["spaceAfter"], got["widowControl"]));
        var raw = XElement.Parse(p.GetRaw());
        var spacing = raw.Descendants().Single(e => e.Name.LocalName == "spacing");
        string? Attr(string name) => spacing.Attributes().FirstOrDefault(a => a.Name.LocalName == name)?.Value;
        Assert.Equal(("50", "120", "100", "240"), (Attr("beforeLines"), Attr("before"), Attr("afterLines"), Attr("after")));
        Assert.Contains(raw.Descendants(), e => e.Name.LocalName == "widowControl" && e.Attributes().Any(a => a.Name.LocalName == "val" && a.Value is "false" or "0"));

        p = Mutations.Set(p, Props(("spaceBefore", "6pt"), ("widowControl", "true")));
        got = p.GetProps();
        Assert.Equal(("6pt", "1lines", "true"), (got["spaceBefore"], got["spaceAfter"], got["widowControl"]));
        Assert.DoesNotContain(XElement.Parse(p.GetRaw()).Descendants(), e => e.Attributes().Any(a => a.Name.LocalName == "beforeLines"));
        Assert.False(Mutations.Set(p, Props(("widowControl", "none"))).GetProps().ContainsKey("widowControl"));
    }

    [Fact]
    public void Widow_control_a_style_turns_off_is_computed_not_the_paragraphs_own()
    {
        static void Off(MainDocumentPart main) => main.StyleDefinitionsPart!.Styles!.Append(new W.Style(
            new W.StyleName { Val = "Loose" }, new W.StyleParagraphProperties(new W.WidowControl { Val = false })) { Type = W.StyleValues.Paragraph, StyleId = "Loose" });
        using var doc = OpenDocx(Docx([P("By its style", "Loose"), P("Plain")], Off));
        var (styled, plain) = (PathResolver.Single(doc.Root, "/body/paragraph[1]"), PathResolver.Single(doc.Root, "/body/paragraph[2]"));
        Assert.False(styled.GetProps().ContainsKey("widowControl"));
        Assert.Equal("false", styled.GetComputed(styled.GetProps())?.GetValueOrDefault("widowControl"));
        Assert.Null(plain.GetComputed(plain.GetProps())?.GetValueOrDefault("widowControl"));
        var own = Mutations.Set(styled, Props(("widowControl", "true")));
        Assert.Equal("true", own.GetProps()["widowControl"]);
        Assert.Null(own.GetComputed(own.GetProps())?.GetValueOrDefault("widowControl"));
    }
}
