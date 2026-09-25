using System.Xml.Linq;
using Writer.Core;
using Writer.Formats.Docx;

namespace Writer.Tests;

/// <summary>Word's 段落 dialog on a paragraph: line spacing, spacing before and after, indents in lengths or characters, borders,
/// keep with next, tab stops and distributed alignment, written and read back in the same words.</summary>
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
}
