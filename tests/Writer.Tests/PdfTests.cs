using Writer.Core;
using Writer.Formats;
using Writer.Formats.Html;
using Writer.Formats.Markdown;
using Writer.Formats.Pdf;

namespace Writer.Tests;

public class PdfTests
{
    static Document Open() => new PdfAdapter().Open(new MemoryStream(TestDocs.Pdf()));

    [Fact]
    public void Reads_pages_text_blocks_and_positions()
    {
        using var doc = Open();
        Assert.Equal("2", doc.Root.GetProps()["pages"]);
        var page = doc.Root.Children[0];
        Assert.Equal("/page[1]", page.Path);
        Assert.EndsWith("cm", Registry.ToDisplay("page", page.GetProps())["width"]);
        var blocks = page.Children;
        Assert.Equal("text", blocks[0].Kind);
        Assert.Equal("Quarterly Report", blocks[0].Text);
        Assert.Contains("Revenue grew 25% year over year.", page.GetProps()["text"]);
        var shown = Registry.ToDisplay("text", blocks[0].GetProps());
        Assert.All(new[] { "x", "y", "w", "h" }, k => Assert.EndsWith("cm", shown[k]));
        Assert.Single(PathResolver.Query(doc.Root, "//text[@text~=\"costs\"]"));
        var text = Views.Text(doc.Root);
        Assert.Contains("--- /page[2] ---", text);
        Assert.Contains("Second page", text);
        Assert.Contains("/page[1]/text[1]", Views.Outline(doc.Root));
    }

    [Fact]
    public void Is_read_only_with_export_as_the_way_out()
    {
        using var doc = Open();
        var block = PathResolver.Single(doc.Root, "/page[1]/text[1]");
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(block, new Dictionary<string, string> { ["text"] = "x" }));
        Assert.Equal(ErrorCode.Validation, ex.Code);
        Assert.Equal(ErrorCode.FormatReadonly, Assert.Throws<WriterException>(() => block.Remove()).Code);
        Assert.Equal(ErrorCode.FormatReadonly, Assert.Throws<WriterException>(() => doc.Save(new MemoryStream())).Code);
        Assert.Equal(ErrorCode.FormatReadonly, Assert.Throws<WriterException>(() => new PdfAdapter().Create()).Code);
        Assert.Equal(ErrorCode.FormatReadonly, Assert.Throws<WriterException>(() => Exporter.Export(doc, new PdfAdapter(), null)).Code);
    }

    [Fact]
    public void Exports_to_markdown_and_html()
    {
        using var doc = Open();
        var (md, warnings) = Exporter.Export(doc, new MdAdapter(), null);
        Assert.Empty(warnings);
        var text = MdTests.Save(md);
        Assert.StartsWith("Quarterly Report\n\nRevenue grew 25% year over year.", text);
        Assert.Contains("Second page", text);
        var html = HtmlWriter.Render(doc);
        Assert.Contains("<section class=\"page\">", html);
        Assert.Contains("Quarterly Report", html);
    }

    [Fact]
    public void Garbage_is_a_format_error()
    {
        var ex = Assert.Throws<WriterException>(() => new PdfAdapter().Open(new MemoryStream("not a pdf"u8.ToArray())));
        Assert.Equal(ErrorCode.FormatError, ex.Code);
    }
}
