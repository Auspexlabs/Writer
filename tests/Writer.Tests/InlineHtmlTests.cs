using Writer.Core;
using Writer.Formats.Common;
using Writer.Formats.Docx;
using Writer.Formats.Markdown;
using Writer.Formats.Pptx;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class InlineHtmlTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    [Fact]
    public void Parses_browser_markup_into_runs()
    {
        var runs = InlineHtml.Parse("Revenue <b>grew</b> <span style=\"color: rgb(255, 0, 0); font-size: 14pt\">25%</span><br>next <a href=\"https://x.y/\">link</a>");
        Assert.Equal(
        [
            new RunSpec("Revenue "),
            new RunSpec("grew", Bold: true),
            new RunSpec(" "),
            new RunSpec("25%", Color: "FF0000", Size: "14"),
            new RunSpec("\nnext "),
            new RunSpec("link", Link: "https://x.y/"),
        ], runs);
    }

    [Fact]
    public void Understands_css_and_font_tags_and_ignores_scripts_and_unknown_tags()
    {
        var runs = InlineHtml.Parse("<p>a<script>alert(1)</script></p><p><strike>b</strike><span style=\"background-color:#ff0;font-family:Georgia, serif;font-weight:700\">c</span></p><font color=\"#00f\" face=\"Arial\">d</font><img src=x><u>e</u>&amp;&lt;");
        Assert.Equal(
        [
            new RunSpec("a\n"),
            new RunSpec("b", Strike: true),
            new RunSpec("c", Bold: true, Font: "Georgia", Highlight: "FFFF00"),
            new RunSpec("d", Color: "0000FF", Font: "Arial"),
            new RunSpec("e", Underline: true),
            new RunSpec("&<"),
        ], runs);
    }

    [Fact]
    public void Renders_runs_as_safe_html_and_round_trips()
    {
        var runs = new List<RunSpec>
        {
            new("a <b>", Bold: true, Italic: true),
            new("\n"),
            new("c", Underline: true, Strike: true, Color: "112233", Size: "10.5", Font: "Arial", Highlight: "FFFF00"),
            new("js", Link: "javascript:alert(1)"),
            new("ok", Link: "https://example.com/?a=1&b=2"),
        };
        var html = InlineHtml.Render(runs);
        Assert.Equal("<b><i>a &lt;b&gt;</i></b><br><span style=\"color:#112233;font-size:10.5pt;font-family:'Arial';background-color:#FFFF00\"><u><s>c</s></u></span><a href=\"#\">js</a><a href=\"https://example.com/?a=1&amp;b=2\">ok</a>", html);
        var back = InlineHtml.Parse(html);
        Assert.Equal(runs.Select(r => r.Link == "javascript:alert(1)" ? r with { Link = "#" } : r), back);
    }

    [Fact]
    public void Caps_and_underline_styles_round_trip()
    {
        List<RunSpec> runs = [new("A", Caps: "all"), new("b", Caps: "small", Underline: true, UnderlineStyle: "double"), new("c", Underline: true)];
        var html = InlineHtml.Render(runs);
        Assert.Equal("<span style=\"text-transform:uppercase\">A</span><span style=\"font-variant:small-caps\"><u style=\"text-decoration-style:double\">b</u></span><u>c</u>", html);
        Assert.Equal(runs, InlineHtml.Parse(html));
        Assert.Equal([new RunSpec("x", Underline: true, UnderlineStyle: "wavy"), new RunSpec("y")],
            InlineHtml.Parse("<span style=\"text-decoration:underline wavy\">x</span><span style=\"text-transform:none;text-decoration-style:dotted\">y</span>"));
    }

    [Fact]
    public void Docx_paragraphs_and_cells_read_and_write_html()
    {
        using var doc = OpenDocx(Docx(P("Old"), new DocumentFormat.OpenXml.Wordprocessing.Table(new DocumentFormat.OpenXml.Wordprocessing.TableRow(Cell("x")))));
        var p = Mutations.Set(PathResolver.Single(doc.Root, "/body/paragraph[1]"),
            Props(("html", "Plain <u>under</u> <span style=\"background-color:#ffff00;color:#c00000;font-size:16px\">mark</span>")));
        Assert.Equal("Plain under mark", p.Text);
        var runs = p.Children;
        Assert.Equal("true", runs[1].GetProps()["underline"]);
        Assert.Equal(("FFFF00", "C00000", "12"), (runs[3].GetProps()["highlight"], runs[3].GetProps()["color"], runs[3].GetProps()["size"]));
        Assert.Equal("Plain <u>under</u> <span style=\"color:#C00000;font-size:12pt;background-color:#FFFF00\">mark</span>", p.GetProps()["html"]);

        var cell = Mutations.Set(PathResolver.Single(doc.Root, "/body/table[1]/row[1]/cell[1]"), Props(("html", "<b>bold</b><br>two")));
        Assert.Equal("bold\ntwo", cell.GetProps()["text"]);
        Assert.Equal("<b>bold</b><br>two", cell.GetProps()["html"]);
        Mutations.Set(runs[1], Props(("highlight", "00FF00")));
        Assert.Equal("00FF00", PathResolver.Single(doc.Root, "/body/paragraph[1]/run[2]").GetProps()["highlight"]);
    }

    [Fact]
    public void Pptx_shapes_and_md_blocks_read_and_write_html()
    {
        using var deck = new PptxAdapter().Create();
        var slide = Mutations.Add(deck.Root, "slide", Props(("layout", "Blank")), null);
        var shape = Mutations.Add(slide, "shape", Props(("html", "<p>One <i>it</i></p><p><span style=\"color:#00ff00\">Two</span></p>")), null);
        Assert.Equal("One it\nTwo", shape.GetProps()["text"]);
        Assert.Equal("One <i>it</i><br><span style=\"color:#00FF00\">Two</span>", shape.GetProps()["html"]);
        Assert.Equal("true", shape.Children[0].Children[1].GetProps()["italic"]);

        using var md = new MdAdapter().Create();
        var h = Mutations.Add(md.Root.Children[0], "heading", Props(("html", "Title <code>x</code>"), ("level", "1")), null);
        Assert.Equal("Title <code>x</code>", h.GetProps()["html"]);
        var ms = new MemoryStream();
        md.Save(ms);
        Assert.Equal("# Title `x`\n", System.Text.Encoding.UTF8.GetString(ms.ToArray()));
    }

    [Fact]
    public void Images_can_be_added_from_data_urls()
    {
        using var doc = OpenDocx(Docx(P("x")));
        var src = "data:image/png;base64," + Convert.ToBase64String(FakePng(3, 2));
        var image = Mutations.Add(PathResolver.Single(doc.Root, "/body"), "image", Props(("src", src)), null);
        Assert.Equal("image/png", image.GetBinary()!.Value.ContentType);
        Assert.Equal(FakePng(3, 2), image.GetBinary()!.Value.Data);
        Assert.Contains("image.png", image.GetProps()["src"]);
        var bad = Assert.Throws<WriterException>(() => Mutations.Add(PathResolver.Single(doc.Root, "/body"), "image", Props(("src", "data:image/png;base64,***")), null));
        Assert.Equal(ErrorCode.Validation, bad.Code);
    }
}
