using DocumentFormat.OpenXml;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class DocxReadTests
{
    [Fact]
    public void Projects_headings_paragraphs_and_code_with_paths()
    {
        using var doc = OpenDocx(Docx(P("Intro", "Heading1"), P("Body text"), P("Sub", "Heading2"), P("print(1)", "Code")));
        var body = doc.Root.Children.Single();
        Assert.Equal("/body", body.Path);
        Assert.Equal(new[] { "/body/heading[1]", "/body/paragraph[1]", "/body/heading[2]", "/body/code[1]" }, body.Children.Select(c => c.Path));
        var h1 = body.Children[0];
        Assert.Equal("Intro", h1.Text);
        Assert.Equal("1", h1.GetProps()["level"]);
        Assert.Equal("2", body.Children[2].GetProps()["level"]);
        Assert.Equal("docx", doc.Root.GetProps()["format"]);
        Assert.Equal("docx", h1.Format);
        Assert.False(body.Children[1].GetProps().ContainsKey("style"));
    }

    [Fact]
    public void Recognises_localised_heading_styles_and_outline_levels()
    {
        var outlineOnly = P("Outline only");
        outlineOnly.ParagraphProperties = new W.ParagraphProperties(new W.OutlineLevel { Val = 1 });
        var bytes = Docx([P("Chinese heading", "3"), outlineOnly, P("plain")], main =>
            main.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.StyleName { Val = "heading 3" }) { Type = W.StyleValues.Paragraph, StyleId = "3" }));
        using var doc = OpenDocx(bytes);
        var blocks = doc.Root.Children.Single().Children;
        Assert.Equal("heading", blocks[0].Kind);
        Assert.Equal("3", blocks[0].GetProps()["level"]);
        Assert.Equal("heading", blocks[1].Kind);
        Assert.Equal("2", blocks[1].GetProps()["level"]);
        Assert.Equal("paragraph", blocks[2].Kind);
    }

    [Fact]
    public void Code_blocks_keep_line_breaks_and_have_no_runs()
    {
        var code = new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Code" }),
            new W.Run(new W.Text("a = 1"), new W.Break(), new W.Text("b = 2")));
        using var doc = OpenDocx(Docx(code));
        var node = PathResolver.Single(doc.Root, "/body/code[1]");
        Assert.Equal("a = 1\nb = 2", node.Text);
        Assert.Empty(node.Children);
        Assert.False(node.GetProps().ContainsKey("style"));
    }

    [Fact]
    public void Lists_are_paragraphs_with_list_and_level()
    {
        using var doc = OpenDocx(Docx([ListP("one", 1, 0), ListP("two", 1, 1), ListP("first", 2, 0), P("plain")], AddNumbering));
        var ps = PathResolver.Query(doc.Root, "/body/paragraph");
        Assert.Equal(("bullet", "0"), (ps[0].GetProps()["list"], ps[0].GetProps()["level"]));
        Assert.Equal(("bullet", "1"), (ps[1].GetProps()["list"], ps[1].GetProps()["level"]));
        Assert.Equal("number", ps[2].GetProps()["list"]);
        Assert.False(ps[3].GetProps().ContainsKey("list"));
        Assert.False(ps[3].GetProps().ContainsKey("level"));
    }

    [Fact]
    public void Runs_expose_direct_formatting_and_links()
    {
        var p = new W.Paragraph(
            new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Arial" }, new W.Bold(), new W.Italic(), new W.Strike(),
                    new W.Color { Val = "ff0000" }, new W.FontSize { Val = "28" }, new W.Underline { Val = W.UnderlineValues.Single }),
                new W.Text("Loud")),
            new W.Run(new W.RunProperties(new W.Bold { Val = false }), new W.Text(" quiet") { Space = SpaceProcessingModeValues.Preserve }),
            new W.Hyperlink(new W.Run(new W.Text("site"))) { Id = "rIdLink" },
            new W.Hyperlink(new W.Run(new W.Text("top"))) { Anchor = "Top" });
        var bytes = Docx([p], main => main.AddHyperlinkRelationship(new Uri("https://example.com/x"), true, "rIdLink"));
        using var doc = OpenDocx(bytes);
        var runs = PathResolver.Query(doc.Root, "/body/paragraph[1]/run");
        Assert.Equal(4, runs.Count);
        var loud = Registry.ToDisplay("run", runs[0].GetProps());
        Assert.Equal("Loud", loud["text"]);
        Assert.Equal("true", loud["bold"]);
        Assert.Equal("true", loud["italic"]);
        Assert.Equal("true", loud["strike"]);
        Assert.Equal("true", loud["underline"]);
        Assert.Equal("FF0000", loud["color"]);
        Assert.Equal("14pt", loud["size"]);
        Assert.Equal("Arial", loud["font"]);
        Assert.False(runs[1].GetProps().ContainsKey("bold"));
        Assert.Equal("https://example.com/x", runs[2].GetProps()["link"]);
        Assert.Equal("#Top", runs[3].GetProps()["link"]);
        Assert.Equal("Loud quietsitetop", PathResolver.Single(doc.Root, "/body/paragraph[1]").Text);
    }

    [Fact]
    public void Revisions_and_content_controls_read_as_final_text()
    {
        var p = new W.Paragraph(
            new W.Run(new W.Text("keep ") { Space = SpaceProcessingModeValues.Preserve }),
            new W.DeletedRun(new W.Run(new W.DeletedText("gone "))),
            new W.InsertedRun(new W.Run(new W.Text("new"))));
        var sdt = new W.SdtBlock(new W.SdtProperties(), new W.SdtContentBlock(P("inside control")));
        using var doc = OpenDocx(Docx(p, sdt, P("after")));
        var blocks = doc.Root.Children.Single().Children;
        Assert.Equal(new[] { "keep new", "inside control", "after" }, blocks.Select(b => b.Text));
        var runs = PathResolver.Query(doc.Root, "/body/paragraph[1]/run");
        Assert.Equal(3, runs.Count);
        Assert.Equal(("gone ", "deleted"), (runs[1].Text, runs[1].GetProps()["change"]));
        Assert.Equal("inserted", runs[2].GetProps()["change"]);
    }

    [Fact]
    public void Tables_expose_rows_cols_and_data()
    {
        var table = new W.Table(
            new W.TableRow(Cell("Name"), Cell("Score")),
            new W.TableRow(Cell("Ann"), Cell("90\n91")));
        using var doc = OpenDocx(Docx(table));
        var t = PathResolver.Single(doc.Root, "/body/table[1]");
        var props = t.GetProps();
        Assert.Equal("2", props["rows"]);
        Assert.Equal("2", props["cols"]);
        Assert.Equal("[[\"Name\",\"Score\"],[\"Ann\",\"90\\n91\"]]", props["data"]);
        Assert.Equal("90\n91", PathResolver.Single(doc.Root, "/body/table[1]/row[2]/cell[2]").Text);
        Assert.Equal("[\"Ann\",\"90\\n91\"]", PathResolver.Single(doc.Root, "/body/table[1]/row[-1]").GetProps()["data"]);
        Assert.Equal("paragraph", PathResolver.Single(doc.Root, "/body/table[1]/row[1]/cell[1]/paragraph[1]").Kind);
        Assert.Contains("Name\tScore\nAnn\t90 91\n", Views.Text(doc.Root));
    }

    [Fact]
    public void Cells_report_fill_and_alignment()
    {
        var cell = new W.TableCell(
            new W.TableCellProperties(new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = "d9e2f3" }),
            new W.Paragraph(new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center }), new W.Run(new W.Text("x"))));
        using var doc = OpenDocx(Docx(new W.Table(new W.TableRow(cell))));
        var props = PathResolver.Single(doc.Root, "//cell").GetProps();
        Assert.Equal("D9E2F3", props["fill"]);
        Assert.Equal("center", props["align"]);
    }

    [Fact]
    public void Pictures_from_the_sample_file_are_images_with_size_and_source()
    {
        using var file = File.OpenRead(Path.Combine(FixtureDir("docx"), "pictures.docx"));
        using var doc = new DocxAdapter().Open(file);
        var images = PathResolver.Query(doc.Root, "//image");
        Assert.NotEmpty(images);
        var props = Registry.ToDisplay("image", images[0].GetProps());
        Assert.Contains("/media/", props["src"]);
        Assert.EndsWith("cm", props["width"]);
        Assert.EndsWith("cm", props["height"]);
    }

    [Fact]
    public void Raw_xml_is_available_for_every_node()
    {
        using var doc = OpenDocx(Docx(P("hi")));
        Assert.StartsWith("<w:p ", PathResolver.Single(doc.Root, "/body/paragraph[1]").GetRaw());
        Assert.StartsWith("<w:r", PathResolver.Single(doc.Root, "/body/paragraph[1]/run[1]").GetRaw());
        Assert.StartsWith("<w:body", PathResolver.Single(doc.Root, "/body").GetRaw());
        Assert.Throws<WriterException>(() => doc.Root.GetRaw());
    }

    [Fact]
    public void Garbage_is_a_format_error()
    {
        var ex = Assert.Throws<WriterException>(() => OpenDocx("not a zip"u8.ToArray()));
        Assert.Equal(ErrorCode.FormatError, ex.Code);
    }

    [Fact]
    public void Blank_template_has_an_empty_body_and_a_section()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        Assert.Empty(body.Children);
        Assert.Contains("w:sectPr", body.GetRaw());
    }
}
