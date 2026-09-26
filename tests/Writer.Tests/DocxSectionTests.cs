using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using Writer.Formats.Html;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Page setup, header and footer on the document node, and page breaks in the body.</summary>
public class DocxSectionTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    static Document Reopen(Document doc) => OpenDocx(Save(doc));

    static W.SectionProperties Section(Document doc) => ((DocxDocument)doc).Main.Document!.Body!.GetFirstChild<W.SectionProperties>()!;

    /// <summary>The SDK schema validator finds nothing wrong with the section properties or the header and footer parts.</summary>
    static void AssertValidSection(Document doc)
    {
        var errors = new OpenXmlValidator().Validate(((DocxDocument)doc).Package)
            .Where(e => e.Part is HeaderPart or FooterPart || e.Path?.XPath?.Contains("sectPr", StringComparison.Ordinal) == true)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void Blank_document_reads_a4_portrait_normal_margins_one_column()
    {
        using var doc = new DocxAdapter().Create();
        var props = doc.Root.GetProps();
        Assert.Equal(("A4", "portrait", "normal", "1"), (props["page"], props["orientation"], props["margin"], props["columns"]));
        Assert.False(props.ContainsKey("header"));
        Assert.False(props.ContainsKey("footer"));
    }

    [Fact]
    public void Page_setup_survives_reopen_in_schema_order()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, Props(("page", "Letter"), ("orientation", "landscape"), ("margin", "narrow"), ("columns", "2"), ("header", "H"), ("footer", "F")));
        using var reopened = Reopen(doc);
        var props = reopened.Root.GetProps();
        Assert.Equal(("Letter", "landscape", "narrow", "2", "H", "F"),
            (props["page"], props["orientation"], props["margin"], props["columns"], props["header"], props["footer"]));
        var section = Section(reopened);
        Assert.Equal(new[] { "headerReference", "footerReference", "pgSz", "pgMar", "cols" }, section.ChildElements.Select(e => e.LocalName));
        AssertValidSection(reopened);
        var size = section.GetFirstChild<W.PageSize>()!;
        Assert.Equal((15840U, 12240U), (size.Width!.Value, size.Height!.Value));

        var root = Mutations.Set(reopened.Root, Props(("orientation", "portrait"), ("columns", "1")));
        Assert.Equal(("Letter", "portrait", "1"), (root.GetProps()["page"], root.GetProps()["orientation"], root.GetProps()["columns"]));
        Assert.Null(Section(reopened).GetFirstChild<W.Columns>());
    }

    [Fact]
    public void Custom_size_and_margins_read_back_in_cm_and_bad_values_are_rejected()
    {
        using var doc = new DocxAdapter().Create();
        var root = Mutations.Set(doc.Root, Props(("page", "20cm x 25cm"), ("margin", "2cm 1cm")));
        Assert.Equal("20cm x 25cm", root.GetProps()["page"]);
        Assert.Equal("2cm 1cm 2cm 1cm", root.GetProps()["margin"]);
        using var reopened = Reopen(doc);
        Assert.Equal("20cm x 25cm", reopened.Root.GetProps()["page"]);
        Assert.Equal("2cm 1cm 2cm 1cm", reopened.Root.GetProps()["margin"]);
        Assert.Equal(ErrorCode.Validation, Assert.Throws<WriterException>(() => Mutations.Set(doc.Root, Props(("page", "Tabloid")))).Code);
        Assert.Equal(ErrorCode.Validation, Assert.Throws<WriterException>(() => Mutations.Set(doc.Root, Props(("margin", "huge")))).Code);
        Assert.Throws<WriterException>(() => Mutations.Set(doc.Root, Props(("margin", "1cm 2cm 3cm"))));
        Assert.Throws<WriterException>(() => Mutations.Set(doc.Root, Props(("columns", "5"))));
    }

    [Fact]
    public void Header_and_footer_round_trip_and_empty_removes_the_part()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, Props(("header", "Q3 <b>report</b>"), ("footer", "Page {page} of {pages}")));
        using var reopened = Reopen(doc);
        var props = reopened.Root.GetProps();
        Assert.Equal("Q3 <b>report</b>", props["header"]);
        Assert.Equal("Page {page} of {pages}", props["footer"]);
        var footer = ((DocxDocument)reopened).Main.FooterParts.Single().Footer!.OuterXml;
        Assert.Contains(" PAGE ", footer);
        Assert.Contains(" NUMPAGES ", footer);
        AssertValidSection(reopened);

        Mutations.Set(reopened.Root, Props(("footer", ""), ("header", "New")));
        using var again = Reopen(reopened);
        Assert.False(again.Root.GetProps().ContainsKey("footer"));
        Assert.Empty(((DocxDocument)again).Main.FooterParts);
        Assert.DoesNotContain("footerReference", Section(again).OuterXml);
        Assert.Equal("New", again.Root.GetProps()["header"]);
        Assert.Single(((DocxDocument)again).Main.HeaderParts);
    }

    [Fact]
    public void Changing_the_page_keeps_an_existing_footer_untouched()
    {
        var original = File.ReadAllBytes(Path.Combine(FixtureDir("docx"), "fields.docx"));
        using var doc = new DocxAdapter().Open(new MemoryStream(original));
        const string footer = "<p style=\"text-align:center\">Page {page} of {pages}</p>";
        Assert.Equal(footer, doc.Root.GetProps()["footer"]);
        Mutations.Set(doc.Root, Props(("page", "Letter")));
        var saved = Save(doc);
        Assert.Null(PackageCompare.Diff(original, saved, "word/document.xml"));
        using var reopened = OpenDocx(saved);
        Assert.Equal("Letter", reopened.Root.GetProps()["page"]);
        Assert.Equal(footer, reopened.Root.GetProps()["footer"]);
    }

    [Fact]
    public void Header_formatting_alignment_and_page_fields_round_trip()
    {
        using var doc = new DocxAdapter().Create();
        const string footer = "<p style=\"text-align:center\"><b>第 {page} 页</b> / 共 <span style=\"color:#C00000;font-size:9pt\"><i>{pages}</i></span> 页</p><p style=\"text-align:right\">x</p>";
        Mutations.Set(doc.Root, Props(("footer", footer)));
        using var reopened = Reopen(doc);
        Assert.Equal(footer, reopened.Root.GetProps()["footer"]);
        var xml = ((DocxDocument)reopened).Main.FooterParts.Single().Footer!.OuterXml;
        Assert.Contains(" PAGE ", xml);
        Assert.Contains(" NUMPAGES ", xml);
        AssertValidSection(reopened);
    }

    [Fact]
    public void First_page_header_and_footer_round_trip_with_titlePg()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, Props(("footer", "Page {page}"), ("titlePg", "true"), ("firstHeader", "Cover"), ("firstFooter", "")));
        using var reopened = Reopen(doc);
        var props = reopened.Root.GetProps();
        Assert.Equal(("true", "Cover", "Page {page}"), (props["titlePg"], props["firstHeader"], props["footer"]));
        Assert.False(props.ContainsKey("firstFooter"), "the first page shows no footer");
        Assert.False(props.ContainsKey("header"));
        Assert.Contains("w:type=\"first\"", Section(reopened).OuterXml);
        AssertValidSection(reopened);

        var off = Mutations.Set(reopened.Root, Props(("titlePg", "false"))).GetProps();
        Assert.False(off.ContainsKey("titlePg"));
        Assert.Equal("Cover", off["firstHeader"]); // kept for when it is switched on again, as Word does
    }

    const string LogoHeader = """
        <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:p><w:pPr><w:pStyle w:val="Header"/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr></w:pPr><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="Logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t xml:space="preserve"> Acme </w:t></w:r><w:fldSimple w:instr=" DATE "><w:r><w:t>2026-09-23</w:t></w:r></w:fldSimple></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>T</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:hdr>
        """;

    [Fact]
    public void Editing_a_header_keeps_its_picture_field_table_and_paragraph_style()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, Props(("header", "x")));
        var part = ((DocxDocument)doc).Main.HeaderParts.Single();
        using (var png = new MemoryStream(FakePng(100, 50))) part.AddImagePart(ImagePartType.Png, "rIdLogo").FeedData(png);
        part.Header = new W.Header(LogoHeader);
        var html = doc.Root.GetProps()["header"];
        Assert.StartsWith("<img data-keep=\"0\" width=\"100\" height=\"50\" src=\"data:image/png;base64,", html);
        Assert.EndsWith("\"> Acme <span data-keep=\"1\">2026-09-23</span>", html);

        var edited = html.Replace(" Acme ", " <b>Acme Inc</b> ");
        Mutations.Set(doc.Root, Props(("header", edited)));
        using var reopened = Reopen(doc);
        Assert.Equal(edited, reopened.Root.GetProps()["header"]);
        var xml = ((DocxDocument)reopened).Main.HeaderParts.Single().Header!.OuterXml;
        foreach (var kept in new[] { "w:pStyle w:val=\"Header\"", "<w:pBdr>", "r:embed=\"rIdLogo\"", "w:instr=\" DATE \"", "<w:tbl>" }) Assert.Contains(kept, xml);
        AssertValidSection(reopened);
    }

    [Fact]
    public void Page_break_is_a_block_of_its_own()
    {
        var mixed = new W.Paragraph(new W.Run(new W.Text("x"), new W.Break { Type = W.BreakValues.Page }));
        using var doc = OpenDocx(Docx(P("A"), P("B"), mixed));
        var body = doc.Root.Children.Single();
        var brk = Mutations.Add(body, "pagebreak", Props(), 2);
        Assert.Equal("/body/pagebreak[1]", brk.Path);
        Assert.Empty(brk.GetProps());
        Assert.Contains("w:type=\"page\"", brk.GetRaw());
        Assert.Equal(new[] { "paragraph", "pagebreak", "paragraph", "paragraph" }, body.Children.Select(c => c.Kind));
        Assert.Contains("/body/pagebreak[1]\n", Views.Outline(doc.Root));
        Assert.Contains("<div class=\"pagebreak\" style=\"page-break-after:always\"></div>", HtmlWriter.Render(doc));

        using var reopened = Reopen(doc);
        PathResolver.Single(reopened.Root, "/body/pagebreak[1]").Remove();
        Assert.Equal(new[] { "paragraph", "paragraph", "paragraph" }, reopened.Root.Children.Single().Children.Select(c => c.Kind));
        Assert.DoesNotContain("pagebreak", Views.Outline(reopened.Root));
    }

    [Fact]
    public void A_paragraph_ends_a_section_with_its_own_page_setup()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var first = Mutations.Add(body, "paragraph", Props(("text", "竖排的第一节"), ("sectionBreak", "nextPage"), ("orientation", "landscape"), ("columns", "2"), ("margin", "narrow")), null);
        Mutations.Add(body, "paragraph", Props(("text", "第二节")), null);
        Assert.Throws<WriterException>(() => Mutations.Set(PathResolver.Single(doc.Root, "/body/paragraph[2]"), Props(("columns", "2"))));
        AssertValidSection(doc);

        using var reopened = Reopen(doc);
        var got = PathResolver.Single(reopened.Root, "/body/paragraph[1]").GetProps();
        Assert.Equal(("nextPage", "landscape", "2", "narrow", "A4"), (got["sectionBreak"], got["orientation"], got["columns"], got["margin"], got["page"]));
        Assert.False(PathResolver.Single(reopened.Root, "/body/paragraph[2]").GetProps().ContainsKey("sectionBreak"));
        var root = reopened.Root.GetProps();
        Assert.Equal(("portrait", "1"), (root["orientation"], root["columns"])); // the document's last section keeps its own

        var p = Mutations.Set(PathResolver.Single(reopened.Root, "/body/paragraph[1]"), Props(("sectionBreak", "continuous")));
        Assert.Equal(("continuous", "landscape"), (p.GetProps()["sectionBreak"], p.GetProps()["orientation"]));
        p = Mutations.Set(p, Props(("sectionBreak", "none")));
        Assert.False(p.GetProps().ContainsKey("sectionBreak"));
        Assert.False(p.GetProps().ContainsKey("orientation"));
        Assert.Null(((DocxDocument)reopened).Main.Document!.Body!.GetFirstChild<W.Paragraph>()!.ParagraphProperties);
    }

    [Fact]
    public void Line_numbers_and_hyphenation_round_trip()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, Props(("lineNumbers", "true"), ("hyphenation", "true")));
        AssertValidSection(doc);
        using var reopened = Reopen(doc);
        var got = reopened.Root.GetProps();
        Assert.Equal(("true", "true"), (got["lineNumbers"], got["hyphenation"]));
        Mutations.Set(reopened.Root, Props(("lineNumbers", "false"), ("hyphenation", "false")));
        got = reopened.Root.GetProps();
        Assert.False(got.ContainsKey("lineNumbers"));
        Assert.False(got.ContainsKey("hyphenation"));
    }

    [Fact]
    public void Note_numbering_is_set_for_both_kinds_in_schema_order_and_none_takes_it_out()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, Props(("noteFormat", "lowerRoman"), ("columns", "2")));
        AssertValidSection(doc);
        using var reopened = Reopen(doc);
        Assert.Equal("lowerRoman", reopened.Root.GetProps()["noteFormat"]);
        var section = Section(reopened);
        Assert.Equal(new[] { "footnotePr", "endnotePr", "pgSz", "pgMar", "cols" }, section.ChildElements.Select(e => e.LocalName));
        Mutations.Set(reopened.Root, Props(("noteFormat", "decimalEnclosedCircleChinese")));
        Assert.Equal("decimalEnclosedCircleChinese", reopened.Root.GetProps()["noteFormat"]);
        Assert.Single(Section(reopened).Elements<W.EndnoteProperties>());
        Mutations.Set(reopened.Root, Props(("noteFormat", "none")));
        Assert.False(reopened.Root.GetProps().ContainsKey("noteFormat"));
        Assert.Null(Section(reopened).GetFirstChild<W.FootnoteProperties>());
        Assert.Throws<WriterException>(() => Mutations.Set(reopened.Root, Props(("noteFormat", "hebrew1"))));
    }

    [Fact]
    public void Custom_margins_on_four_sides_read_back_as_the_editor_writes_them()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, Props(("margin", "2.54cm 3.18cm 2.54cm 3.18cm")));
        Assert.Equal("2.54cm 3.18cm 2.54cm 3.18cm", doc.Root.GetProps()["margin"]);
        Mutations.Set(doc.Root, Props(("margin", "moderate")));
        Assert.Equal("moderate", doc.Root.GetProps()["margin"]);
        Mutations.Set(doc.Root, Props(("page", "B5")));
        Assert.Equal("B5", doc.Root.GetProps()["page"]);
    }

    [Fact]
    public void Footnotes_and_endnotes_round_trip_at_their_offsets()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "Hello world")), null);
        var note = Mutations.Add(p, "footnote", Props(("text", "First note"), ("at", "5")), null);
        Mutations.Add(p, "footnote", Props(("kind", "endnote"), ("text", "Line one\nLine two")), null);
        Assert.Equal("Hello world", p.GetProps()["text"]); // the marks are not text
        var errors = new OpenXmlValidator().Validate(((DocxDocument)doc).Package).Where(e => e.Part is FootnotesPart or EndnotesPart || e.Part is MainDocumentPart).Select(e => e.Description).ToList();
        Assert.Empty(errors);

        using var reopened = Reopen(doc);
        var notes = PathResolver.Query(reopened.Root, "//footnote").ToList();
        Assert.Equal(2, notes.Count);
        var (a, b) = (notes[0].GetProps(), notes[1].GetProps());
        Assert.Equal(("footnote", "First note", "5"), (a["kind"], a["text"], a["at"]));
        Assert.Equal(("endnote", "Line one\nLine two", "11"), (b["kind"], b["text"], b["at"]));
        var n = PathResolver.Single(reopened.Root, $"//footnote[@id={a["id"]}]");
        n = Mutations.Set(n, Props(("text", "Changed"), ("at", "0")));
        Assert.Equal(("Changed", "0"), (n.GetProps()["text"], n.GetProps()["at"]));
        Assert.Equal("Hello world", PathResolver.Single(reopened.Root, "/body/paragraph[1]").GetProps()["text"]);
        PathResolver.Single(reopened.Root, "/body/paragraph[1]").Remove();
        Assert.Empty(PathResolver.Query(reopened.Root, "//footnote"));
        var main = ((DocxDocument)reopened).Main;
        Assert.DoesNotContain(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>(), f => f.Id?.Value > 0); // only the separators are left
    }
}
