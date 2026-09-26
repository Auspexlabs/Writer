using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using Writer.Formats.Html;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>The table of contents block: a Word content control around a TOC field, regenerated from the headings.</summary>
public class DocxTocTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    static Document WithHeadings()
    {
        var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        Mutations.Add(body, "heading", Props(("text", "Intro"), ("level", "1")), null);
        Mutations.Add(body, "paragraph", Props(("text", "Body")), null);
        Mutations.Add(body, "heading", Props(("text", "Scope"), ("level", "2")), null);
        Mutations.Add(body, "heading", Props(("text", "Deep"), ("level", "4")), null);
        return doc;
    }

    static void AssertValid(Document doc, bool styles = true)
    {
        var errors = new OpenXmlValidator().Validate(((DocxDocument)doc).Package)
            .Where(e => e.Part is MainDocumentPart or DocumentSettingsPart || (styles && e.Part is StyleDefinitionsPart))
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void Style_says_how_entries_end_and_is_read_back_from_the_field_and_its_tab()
    {
        using var doc = WithHeadings();
        var body = doc.Root.Children.Single();
        var toc = Mutations.Add(body, "toc", Props(("title", "目录")), 1);
        Assert.Equal("classic", toc.GetProps()["style"]);
        Assert.Contains("w:leader=\"dot\"", toc.GetRaw());

        toc = Mutations.Set(toc, Props(("style", "simple")));
        Assert.Equal(("simple", "Intro\nScope"), (toc.GetProps()["style"], toc.GetProps()["text"]));
        Assert.Contains("w:leader=\"none\"", toc.GetRaw());
        Assert.Contains("PAGEREF", toc.GetRaw());

        toc = Mutations.Set(toc, Props(("style", "plain")));
        var raw = toc.GetRaw();
        Assert.Equal("plain", toc.GetProps()["style"]);
        Assert.Contains(" TOC \\o \"1-3\" \\h \\z \\u \\n ", raw);
        Assert.DoesNotContain("PAGEREF", raw);
        AssertValid(doc);

        using var reopened = OpenDocx(Save(doc));
        var again = PathResolver.Single(reopened.Root, "/body/toc[1]");
        Assert.Equal("plain", again.GetProps()["style"]);
        again = Mutations.Set(again, Props(("levels", "2")));
        Assert.Equal("plain", again.GetProps()["style"]); // a new level count keeps the style
    }

    [Fact]
    public void Add_builds_a_word_table_of_contents_with_linked_entries()
    {
        using var doc = WithHeadings();
        var body = doc.Root.Children.Single();
        var toc = Mutations.Add(body, "toc", Props(("title", "目录")), 1);
        Assert.Equal("/body/toc[1]", toc.Path);
        Assert.Equal(("3", "目录", "Intro\nScope"), (toc.GetProps()["levels"], toc.GetProps()["title"], toc.GetProps()["text"]));
        Assert.Equal(new[] { "toc", "heading", "paragraph", "heading", "heading" }, body.Children.Select(c => c.Kind));

        var raw = toc.GetRaw();
        Assert.StartsWith("<w:sdt", raw);
        Assert.Contains("w:docPartGallery w:val=\"Table of Contents\"", raw);
        Assert.Contains(" TOC \\o \"1-3\" \\h \\z \\u ", raw);
        Assert.Contains("w:pStyle w:val=\"TOCHeading\"", raw);
        Assert.Contains("w:pStyle w:val=\"TOC1\"", raw);
        Assert.Contains("w:pStyle w:val=\"TOC2\"", raw);
        Assert.DoesNotContain("TOC4", raw);
        Assert.Contains("w:leader=\"dot\"", raw);
        Assert.Equal(2, raw.Split("<w:hyperlink ").Length - 1);
        Assert.Contains("w:anchor=\"_Toc", raw);
        Assert.Contains(" PAGEREF _Toc", raw);
        Assert.Contains("_Toc", PathResolver.Single(doc.Root, "/body/heading[1]").GetRaw());

        var main = ((DocxDocument)doc).Main;
        Assert.True(main.DocumentSettingsPart!.Settings!.GetFirstChild<W.UpdateFieldsOnOpen>()!.Val!.Value);
        Assert.Contains(main.StyleDefinitionsPart!.Styles!.Elements<W.Style>(), s => s.StyleId == "TOC2");

        var html = HtmlWriter.Render(doc);
        Assert.Contains("<nav class=\"toc\">\n<p class=\"toc-title\">目录</p>\n<ul>\n<li><span>Intro</span></li>\n<li><span>Scope</span></li>\n</ul>\n</nav>", html);
        Assert.StartsWith("Intro\nScope\n\nIntro\n\nBody", Views.Text(doc.Root));
        AssertValid(doc);
    }

    [Fact]
    public void Set_regenerates_the_entries_and_remove_deletes_the_control()
    {
        using var doc = WithHeadings();
        var body = doc.Root.Children.Single();
        var toc = Mutations.Add(body, "toc", Props(), null);
        Assert.False(toc.GetProps().ContainsKey("title"));
        toc = Mutations.Set(toc, Props(("levels", "1")));
        Assert.Equal("Intro", toc.GetProps()["text"]);
        Assert.Contains("\\o \"1-1\"", toc.GetRaw());
        Mutations.Add(body, "heading", Props(("text", "Later"), ("level", "1")), null);
        toc = Mutations.Set(toc, Props(("levels", "4"), ("title", "Contents")));
        Assert.Equal("Intro\nScope\nDeep\nLater", toc.GetProps()["text"]);
        Assert.Equal("Contents", toc.GetProps()["title"]);
        toc = Mutations.Set(toc, Props(("title", "")));
        Assert.False(toc.GetProps().ContainsKey("title"));
        Assert.DoesNotContain("TOCHeading", toc.GetRaw());
        Assert.Single(body.Children, c => c.Kind == "toc");
        toc.Remove();
        Assert.DoesNotContain(body.Children, c => c.Kind == "toc");
        Assert.DoesNotContain("<w:sdt", body.GetRaw());
    }

    [Fact]
    public void Without_headings_the_placeholder_is_the_only_entry()
    {
        using var doc = new DocxAdapter().Create();
        var toc = Mutations.Add(doc.Root.Children.Single(), "toc", Props(), null);
        Assert.Equal("No table of contents entries found.", toc.GetProps()["text"]);
        AssertValid(doc);
    }

    [Fact]
    public void A_bare_field_over_several_paragraphs_is_one_block()
    {
        static W.Run Fld(W.FieldCharValues type) => new(new W.FieldChar { FieldCharType = type });
        var first = new W.Paragraph(Fld(W.FieldCharValues.Begin),
            new W.Run(new W.FieldCode(" TOC \\o \"1-2\" \\h ") { Space = SpaceProcessingModeValues.Preserve }),
            Fld(W.FieldCharValues.Separate), new W.Run(new W.Text("One"), new W.TabChar(), new W.Text("3")));
        var middle = P("Two\t5");
        var last = new W.Paragraph(Fld(W.FieldCharValues.End));
        using var doc = OpenDocx(Docx(P("Before"), first, middle, last, P("After", "Heading1")));
        var body = doc.Root.Children.Single();
        Assert.Equal(new[] { "paragraph", "toc", "heading" }, body.Children.Select(c => c.Kind));
        var toc = body.Children[1];
        Assert.Equal(("2", "One\t3\nTwo\t5"), (toc.GetProps()["levels"], toc.GetProps()["text"]));
        Assert.Contains("<li><span>One</span><span class=\"toc-page\">3</span></li>", HtmlWriter.Render(doc));

        toc = Mutations.Set(toc, Props(("levels", "3")));
        Assert.StartsWith("<w:sdt", toc.GetRaw());
        Assert.Equal("After", toc.GetProps()["text"]);
        Assert.Equal(new[] { "paragraph", "toc", "heading" }, body.Children.Select(c => c.Kind));
        AssertValid(doc);
    }

    [Fact]
    public void Word_toc_in_the_fixture_projects_as_one_block_and_regenerates_in_place()
    {
        var original = File.ReadAllBytes(Path.Combine(FixtureDir("docx"), "fields.docx"));
        using var doc = new DocxAdapter().Open(new MemoryStream(original));
        var body = doc.Root.Children.Single();
        var toc = Assert.Single(body.Children, c => c.Kind == "toc");
        Assert.Equal("/body/toc[1]", toc.Path);
        Assert.Contains("/body/toc[1]  levels=3", Views.Outline(doc.Root));
        Assert.Contains("<nav class=\"toc\">", HtmlWriter.Render(doc));
        Assert.Contains("Update field to see table of contents", Views.Text(doc.Root));

        Assert.Equal("Contents", toc.GetProps()["title"]); // the TOC Heading paragraph above the field belongs to the block
        Assert.DoesNotContain(body.Children, c => c.Text == "Contents");
        Assert.Contains("<p class=\"toc-title\">Contents</p>", HtmlWriter.Render(doc));

        var index = body.Children.ToList().FindIndex(c => c.Kind == "toc");
        toc = Mutations.Set(toc, Props(("levels", "2")));
        Assert.Equal(index, body.Children.ToList().FindIndex(c => c.Kind == "toc"));
        Assert.Equal("Contents", toc.GetProps()["title"]);
        Assert.StartsWith("1. Introduction\n1.1 Scope\n2. Date & Time Fields", toc.GetProps()["text"]);
        Assert.DoesNotContain("Update field", Views.Text(doc.Root));
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(body.GetRaw(), ">Contents<")); // moved into the control, not copied
        Assert.Null(PackageCompare.Diff(original, Save(doc), "word/document.xml", "word/settings.xml", "word/styles.xml"));
        AssertValid(doc, styles: false); // the fixture's own styles part predates this edit
    }

    [Fact]
    public void A_content_control_that_merely_contains_a_toc_stays_transparent()
    {
        static W.Run Fld(W.FieldCharValues type) => new(new W.FieldChar { FieldCharType = type });
        W.Paragraph Field(string cached) => new(Fld(W.FieldCharValues.Begin),
            new W.Run(new W.FieldCode(" TOC \\o \"1-3\" ") { Space = SpaceProcessingModeValues.Preserve }), Fld(W.FieldCharValues.Separate),
            new W.Run(new W.Text(cached)), Fld(W.FieldCharValues.End));
        var wrapper = new W.SdtBlock(new W.SdtProperties(), new W.SdtContentBlock(P("Intro"), Field("entries"), P("Chapter", "Heading1"), P("Body")));
        var onlyToc = new W.SdtBlock(new W.SdtProperties(), new W.SdtContentBlock(P("Contents"), Field("one")));
        var unterminated = new W.Paragraph(Fld(W.FieldCharValues.Begin), new W.Run(new W.FieldCode(" TOC ")), Fld(W.FieldCharValues.Separate));
        using var doc = OpenDocx(Docx(wrapper, onlyToc, unterminated, P("Tail")));
        Assert.Equal(new[] { "paragraph", "toc", "heading", "paragraph", "toc", "toc", "paragraph" }, doc.Root.Children.Single().Children.Select(c => c.Kind));
        Assert.Equal("Tail", doc.Root.Children.Single().Children[^1].Text);
    }

    [Fact]
    public void Entries_are_one_line_each_and_export_rebuilds_them_from_the_copied_headings()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        Mutations.Add(body, "toc", Props(("title", "目录")), null);
        Mutations.Add(body, "heading", Props(("text", "Part\tone\nand two"), ("level", "1")), null);
        var toc = Mutations.Set(PathResolver.Single(doc.Root, "/body/toc[1]"), Props(("levels", "3")));
        Assert.Equal("Part one and two", toc.GetProps()["text"]);

        var (copy, warnings) = Writer.Formats.Exporter.Export(doc, new DocxAdapter(), null);
        using (copy)
        {
            Assert.Empty(warnings);
            var copied = PathResolver.Single(copy.Root, "/body/toc[1]");
            Assert.Equal(("目录", "Part one and two"), (copied.GetProps()["title"], copied.GetProps()["text"]));
        }
    }

    [Fact]
    public void Toc_is_registered_after_pagebreak_for_docx_bodies()
    {
        var names = Registry.Kinds.Select(k => k.Name).ToList();
        Assert.Equal(names.IndexOf("pagebreak") + 1, names.IndexOf("toc"));
        var toc = Registry.Get("docx", "toc");
        Assert.Equal(["body"], toc.Parents);
        Assert.Equal(["levels", "title", "style", "text"], toc.Props.Select(p => p.Name));
    }
}
