using System.Text;
using Writer.Core;
using Writer.Formats;
using Writer.Formats.Docx;
using Writer.Formats.Html;
using Writer.Formats.Markdown;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class HtmlTests
{
    [Fact]
    public void Docx_renders_headings_runs_and_tables()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        Mutations.Add(body, "heading", new Dictionary<string, string> { ["text"] = "Report <1>", ["level"] = "1" }, null);
        Mutations.Add(body, "paragraph", new Dictionary<string, string> { ["md"] = "Revenue **grew** [here](https://x.y)", ["align"] = "center" }, null);
        Mutations.Add(body, "table", new Dictionary<string, string> { ["data"] = "[[\"a\",\"b\"],[\"1\",\"2\"]]" }, null);
        var html = HtmlWriter.Render(doc);
        Assert.Contains("<title>Report &lt;1&gt;</title>", html);
        Assert.Contains("<h1>Report &lt;1&gt;</h1>", html);
        Assert.Contains("<p style=\"text-align:center\">Revenue <strong>grew</strong> <a href=\"https://x.y\">here</a></p>", html);
        Assert.Matches("<tr><td[^>]*>a</td><td[^>]*>b</td></tr>", html);
        Assert.Matches("<tr><td[^>]*>1</td><td[^>]*>2</td></tr>", html);
    }

    [Fact]
    public void Markdown_renders_nested_lists_code_and_images()
    {
        using var doc = MdTests.Open("- one\n- two **strong**\n  - nested\n- three\n\n1. first\n2. second\n\n```py\nx\n```\n\n![A](missing.png)\n\n| h |\n|---|\n| c |\n");
        var html = HtmlWriter.Render(doc);
        Assert.Contains("<ul>\n<li>one</li>\n<li>two <strong>strong</strong><ul>\n<li>nested</li>\n</ul>\n</li>\n<li>three</li>\n</ul>\n<ol>\n<li>first</li>\n<li>second</li>\n</ol>", html);
        Assert.Contains("<pre><code class=\"language-py\">x</code></pre>", html);
        Assert.Contains("<img src=\"missing.png\" alt=\"A\">", html);
        Assert.Contains("<tr><th>h</th></tr>", html);
        Assert.Contains("<tr><td>c</td></tr>", html);
    }

    [Fact]
    public void Script_and_data_urls_are_neutralised()
    {
        using var doc = MdTests.Open("[x](javascript:alert(1)) [y](mailto:a@b.c) [z](sub/page.html)\n\n![i](data:text/html,evil)\n");
        var html = HtmlWriter.Render(doc);
        Assert.Contains("<a href=\"#\">x</a>", html);
        Assert.Contains("<a href=\"mailto:a@b.c\">y</a>", html);
        Assert.Contains("<a href=\"sub/page.html\">z</a>", html);
        Assert.Contains("<img src=\"#\" alt=\"i\">", html);
    }

    [Fact]
    public void Docx_images_become_data_uris()
    {
        using var doc = new DocxAdapter().Create();
        var png = Path.Combine(Path.GetTempPath(), $"writer-html-{Guid.NewGuid():N}.png");
        File.WriteAllBytes(png, FakePng(2, 2));
        try
        {
            Mutations.Add(doc.Root.Children.Single(), "image", new Dictionary<string, string> { ["src"] = png, ["width"] = "2cm" }, null);
            var html = HtmlWriter.Render(doc);
            Assert.Contains("<img src=\"data:image/png;base64,", html);
            Assert.Contains("style=\"width:76px\"", html);
        }
        finally
        {
            File.Delete(png);
        }
    }
}

public class ExportTests
{
    static Stream Stream(string text) => new MemoryStream(Encoding.UTF8.GetBytes(text));

    [Fact]
    public void Markdown_to_docx_and_back_keeps_structure_and_formatting()
    {
        const string md = "# Report\n\nRevenue **grew** 25% ([source](https://x.y)).\n\n- one\n- two\n\n```py\nprint(1)\n```\n\n| a | b |\n|---|---|\n| 1 | 2 |\n";
        using var source = new MdAdapter().Open(Stream(md));
        var (docx, warnings) = Exporter.Export(source, new DocxAdapter(), null);
        Assert.Empty(warnings);
        var body = docx.Root.Children.Single();
        Assert.Equal(new[] { "heading", "paragraph", "paragraph", "paragraph", "code", "table" }, body.Children.Select(c => c.Kind));
        Assert.Equal("true", PathResolver.Single(docx.Root, "/body/paragraph[1]/run[2]").GetProps()["bold"]);
        Assert.Equal("https://x.y", PathResolver.Single(docx.Root, "/body/paragraph[1]/run[4]").GetProps()["link"]);
        Assert.Equal("bullet", PathResolver.Single(docx.Root, "/body/paragraph[2]").GetProps()["list"]);
        Assert.Equal("[[\"a\",\"b\"],[\"1\",\"2\"]]", PathResolver.Single(docx.Root, "/body/table[1]").GetProps()["data"]);

        var (back, moreWarnings) = Exporter.Export(docx, new MdAdapter(), null);
        Assert.Empty(moreWarnings);
        Assert.Equal("# Report\n\nRevenue **grew** 25% ([source](https://x.y)).\n\n- one\n- two\n\n```\nprint(1)\n```\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n", MdTests.Save(back));
    }

    [Fact]
    public void Docx_sample_exports_to_markdown_with_its_text()
    {
        using var file = File.OpenRead(Path.Combine(FixtureDir("docx"), "tables.docx"));
        using var doc = new DocxAdapter().Open(file);
        var (md, _) = Exporter.Export(doc, new MdAdapter(), null);
        var text = MdTests.Save(md);
        Assert.Contains("| Project Name | Phase | Owner |", text);
        Assert.Equal(Views.Text(doc.Root).Replace("\t", " ").Split('\n')[0], Views.Text(md.Root).Replace("\t", " ").Split('\n')[0]);
    }

    [Fact]
    public void Markdown_becomes_a_deck_and_tables_become_sheets()
    {
        const string md = "# Q4 Review\n\nRevenue **grew** 25%.\n\n- costs flat\n- margin up\n\n## Regions\n\n| Region | Sales |\n|---|---|\n| East | 120 |\n\n### Notes\n\nDetail line.\n";
        using var source = new MdAdapter().Open(Stream(md));
        var (deck, warnings) = Exporter.Export(source, new Writer.Formats.Pptx.PptxAdapter(), null);
        Assert.Empty(warnings);
        var slides = deck.Root.Children;
        Assert.Equal(new[] { "Q4 Review", "Regions", "Regions" }, slides.Select(s => s.GetProps()["title"]));
        var body = slides[0].Children.First(c => c.GetProps().GetValueOrDefault("placeholder") == "body");
        Assert.Equal(new[] { "Revenue grew 25%.", "costs flat", "margin up" }, body.Children.Select(p => p.Text));
        Assert.Equal("true", body.Children[0].Children[1].GetProps()["bold"]);
        Assert.Equal("bullet", body.Children[1].GetProps()["list"]);
        Assert.Contains(slides[1].Children, c => c.Kind == "table");
        Assert.DoesNotContain(slides[1].Children, c => c.GetProps().GetValueOrDefault("placeholder") == "body");
        Assert.Equal(new[] { "Notes", "Detail line." }, slides[2].Children.First(c => c.GetProps().GetValueOrDefault("placeholder") == "body").Children.Select(p => p.Text));

        var (book, bookWarnings) = Exporter.Export(source, new Writer.Formats.Xlsx.XlsxAdapter(), null);
        Assert.Empty(bookWarnings);
        var sheet = book.Root.Children.Single();
        Assert.Equal("Table1", sheet.GetProps()["name"]);
        Assert.Equal("[[\"Region\",\"Sales\"],[\"East\",\"120\"]]", PathResolver.Single(book.Root, "/sheet[1]/range[A1:B2]").GetProps()["values"]);

        var (back, backWarnings) = Exporter.Export(deck, new MdAdapter(), null);
        Assert.Empty(backWarnings);
        var text = MdTests.Save(back);
        Assert.StartsWith("## Q4 Review\n\nRevenue **grew** 25%.\n\n- costs flat\n- margin up\n\n## Regions\n\n| Region | Sales |", text);
    }

    [Fact]
    public void Decks_render_slides_with_absolute_layout()
    {
        using var deck = new Writer.Formats.Pptx.PptxAdapter().Create();
        var slide = Mutations.Add(deck.Root, "slide", new Dictionary<string, string> { ["layout"] = "Blank", ["background"] = "1A1A2E" }, null);
        Mutations.Add(slide, "shape", new Dictionary<string, string> { ["text"] = "Hello", ["x"] = "1in", ["y"] = "2in", ["w"] = "4in", ["h"] = "1in", ["fill"] = "4472C4", ["color"] = "FFFFFF", ["size"] = "24" }, null);
        var html = HtmlWriter.Render(deck);
        Assert.Contains("<body class=\"deck\">", html);
        Assert.Contains("<section class=\"slide\" style=\"width:1280px;height:720px;background:#1A1A2E\">", html);
        // the theme's minor fonts are what PowerPoint shows for a text box that names none
        Assert.Contains("left:96px;top:192px;width:384px;height:96px;background:#4472C4;font-family:'Calibri','Microsoft YaHei';font-size:24pt;color:#FFFFFF;", html);
        Assert.Contains(">Hello</span></p>", html);
    }

    [Fact]
    public void Images_are_written_next_to_markdown_exports()
    {
        var dir = Directory.CreateTempSubdirectory("writer-export").FullName;
        try
        {
            var png = Path.Combine(dir, "pic.png");
            File.WriteAllBytes(png, FakePng(3, 3));
            using var doc = new DocxAdapter().Create();
            Mutations.Add(doc.Root.Children.Single(), "image", new Dictionary<string, string> { ["src"] = png, ["alt"] = "pic" }, null);
            var target = Path.Combine(dir, "out.md");
            var (md, warnings) = Exporter.Export(doc, new MdAdapter(), target);
            Assert.Empty(warnings);
            Assert.Equal("![pic](out_files/image1.png)\n", MdTests.Save(md));
            Assert.True(File.Exists(Path.Combine(dir, "out_files", "image1.png")));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }
}
