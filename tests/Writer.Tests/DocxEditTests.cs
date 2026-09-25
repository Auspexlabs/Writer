using System.Xml.Linq;
using DocumentFormat.OpenXml;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class DocxEditTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    [Fact]
    public void Set_text_keeps_paragraph_style_and_first_run_formatting()
    {
        var p = new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Quote" }),
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("Old ")), new W.Run(new W.Text("text")));
        using var doc = OpenDocx(Docx(p));
        var node = Mutations.Set(PathResolver.Single(doc.Root, "/body/paragraph[1]"), Props(("text", "New")));
        Assert.Equal("New", node.Text);
        Assert.Equal("Quote", node.GetProps()["style"]);
        var runs = node.Children;
        Assert.Single(runs);
        Assert.Equal("true", runs[0].GetProps()["bold"]);
    }

    [Fact]
    public void Set_text_keeps_bookmarks_and_comment_anchors()
    {
        var p = new W.Paragraph(
            new W.BookmarkStart { Id = "1", Name = "mark" },
            new W.Run(new W.Text("old")),
            new W.BookmarkEnd { Id = "1" },
            new W.Run(new W.CommentReference { Id = "7" }));
        using var doc = OpenDocx(Docx(p));
        var node = Mutations.Set(PathResolver.Single(doc.Root, "/body/paragraph[1]"), Props(("text", "new")));
        var raw = node.GetRaw();
        Assert.Contains("w:bookmarkStart", raw);
        Assert.Contains("w:commentReference", raw);
        Assert.Equal("new", node.Text);
        Assert.DoesNotContain(">old<", raw);
    }

    [Fact]
    public void Md_makes_formatted_runs_and_links()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("md", "Plain **bold** *it* `code` [site](https://example.com) ~~gone~~")), null);
        var runs = p.Children.Select(r => Registry.ToDisplay("run", r.GetProps())).ToList();
        Assert.Equal("Plain bold it code site gone", p.Text);
        Assert.Equal(10, runs.Count);
        Assert.Equal("true", runs[1]["bold"]);
        Assert.Equal("true", runs[3]["italic"]);
        Assert.Equal("Consolas", runs[5]["font"]);
        Assert.Equal("https://example.com", runs[7]["link"]);
        Assert.Equal("true", runs[9]["strike"]);
        Assert.Equal("/body/paragraph[1]", p.Path);
    }

    [Fact]
    public void Styles_resolve_by_id_name_or_builtin_and_reject_unknown()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "x"), ("style", "Title")), null);
        Assert.Equal("Title", p.GetProps()["style"]);
        p = Mutations.Set(p, Props(("style", "heading 2")));
        Assert.Equal("heading", p.Kind);
        Assert.Equal("/body/heading[1]", p.Path);
        Assert.Equal("2", p.GetProps()["level"]);
        var q = Mutations.Add(body, "paragraph", Props(("text", "y")), null);
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(q, Props(("style", "Fancy"))));
        Assert.Equal(ErrorCode.Validation, ex.Code);
        Assert.Contains("Quote", ex.Hint);
    }

    [Fact]
    public void Heading_level_changes_the_style()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var h = Mutations.Add(body, "heading", Props(("text", "Intro"), ("level", "3")), null);
        Assert.Equal("3", h.GetProps()["level"]);
        h = Mutations.Set(h, Props(("level", "1")));
        Assert.Equal("1", h.GetProps()["level"]);
        Assert.Contains("Heading1", h.GetRaw());
    }

    [Fact]
    public void Lists_restart_when_separated_and_continue_when_adjacent()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var a = Mutations.Add(body, "paragraph", Props(("text", "one"), ("list", "number")), null);
        var b = Mutations.Add(body, "paragraph", Props(("text", "two"), ("level", "1"), ("list", "number")), null);
        var gap = Mutations.Add(body, "paragraph", Props(("text", "plain")), null);
        var c = Mutations.Add(body, "paragraph", Props(("text", "first again"), ("list", "number")), null);
        var bullet = Mutations.Add(body, "paragraph", Props(("text", "dot"), ("list", "bullet")), null);
        Assert.Equal(("number", "0"), (a.GetProps()["list"], a.GetProps()["level"]));
        Assert.Equal(("number", "1"), (b.GetProps()["list"], b.GetProps()["level"]));
        Assert.False(gap.GetProps().ContainsKey("list"));
        Assert.Equal("bullet", bullet.GetProps()["list"]);
        Assert.Equal("ListParagraph", bullet.GetProps()["style"]);
        Assert.Equal(NumId(a), NumId(b));
        Assert.NotEqual(NumId(a), NumId(c));
        var off = Mutations.Set(bullet, Props(("list", "none")));
        Assert.False(off.GetProps().ContainsKey("list"));
        Assert.False(off.GetProps().ContainsKey("style"));
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(gap, Props(("level", "2"))));
        Assert.Contains("list", ex.Message);
    }

    [Fact]
    public void Outline_and_chinese_lists_count_in_numbering_xml_and_numbering_restarts_where_asked()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var a = Mutations.Add(body, "paragraph", Props(("text", "第一章"), ("list", "chinese")), null);
        var b = Mutations.Add(body, "paragraph", Props(("text", "第一节"), ("list", "chinese"), ("level", "1")), null);
        var c = Mutations.Add(body, "paragraph", Props(("text", "again from one"), ("list", "chinese"), ("restart", "true")), null);
        var d = Mutations.Add(body, "paragraph", Props(("text", "two"), ("list", "chinese")), null);
        var o = Mutations.Add(body, "paragraph", Props(("text", "1.1"), ("list", "outline"), ("level", "1")), null);
        Assert.Equal(("chinese", "1"), (b.GetProps()["list"], b.GetProps()["level"]));
        Assert.Equal(NumId(a), NumId(b));
        Assert.Equal("true", c.GetProps()["restart"]);
        Assert.NotEqual(NumId(a), NumId(c));
        Assert.Equal(NumId(c), NumId(d)); // the item after the restart follows it
        Assert.False(d.GetProps().ContainsKey("restart"));
        Assert.Equal("outline", o.GetProps()["list"]);
        var numbering = ((DocxDocument)doc).Main.NumberingDefinitionsPart!.Numbering!.OuterXml;
        Assert.Contains("chineseCountingThousand", numbering);
        Assert.Contains("（%2）", numbering);
        Assert.Contains("%1.%2", numbering);
        var joined = Mutations.Set(c, Props(("restart", "false")));
        Assert.Equal(NumId(a), NumId(joined));
        Assert.Equal(NumId(a), NumId(Mutations.Refresh(d)));
        Assert.False(joined.GetProps().ContainsKey("restart"));
        Assert.Equal("chinese", Mutations.Set(joined, Props(("list", "chinese"))).GetProps()["list"]); // the same kind again keeps counting
        Assert.Equal(NumId(a), NumId(Mutations.Refresh(joined)));
    }

    static string NumId(Node paragraph) =>
        XElement.Parse(paragraph.GetRaw()).Descendants().First(e => e.Name.LocalName == "numId").Attributes().First().Value;

    [Fact]
    public void Run_formatting_and_links_round_trip()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "hello world")), null);
        var run = p.Children.Single();
        run = Mutations.Set(run, Props(("bold", "true"), ("italic", "yes"), ("color", "red"), ("size", "14"), ("font", "Arial"), ("underline", "true"), ("strike", "true")));
        var shown = Registry.ToDisplay("run", run.GetProps());
        Assert.Equal(("true", "true", "FF0000", "14pt", "Arial", "true", "true"),
            (shown["bold"], shown["italic"], shown["color"], shown["size"], shown["font"], shown["underline"], shown["strike"]));
        run = Mutations.Set(run, Props(("bold", "false"), ("link", "https://example.com/a")));
        Assert.False(run.GetProps().ContainsKey("bold"));
        Assert.Equal("https://example.com/a", run.GetProps()["link"]);
        Assert.Equal("/body/paragraph[1]/run[1]", run.Path);
        Assert.Contains("w:hyperlink", p.GetRaw());
        run = Mutations.Set(run, Props(("link", "#top")));
        Assert.Equal("#top", run.GetProps()["link"]);
        run = Mutations.Set(run, Props(("link", "")));
        Assert.False(run.GetProps().ContainsKey("link"));
        Assert.DoesNotContain("w:hyperlink", p.GetRaw());
        Assert.Equal("hello world", Mutations.Refresh(p).Text);
    }

    [Fact]
    public void Add_positions_by_index_and_keeps_section_properties_last()
    {
        using var doc = OpenDocx(Docx(P("A"), P("C")));
        var body = doc.Root.Children.Single();
        Mutations.Add(body, "paragraph", Props(("text", "B")), 2);
        Mutations.Add(body, "paragraph", Props(("text", "D")), null);
        Mutations.Add(body, "heading", Props(("text", "Top")), 1);
        Assert.Equal(new[] { "Top", "A", "B", "C", "D" }, body.Children.Select(c => c.Text));
        var raw = body.GetRaw();
        Assert.True(raw.IndexOf("w:sectPr", StringComparison.Ordinal) > raw.IndexOf(">D<", StringComparison.Ordinal));
        Assert.Equal("/body/paragraph[4]", body.Children[4].Path);
    }

    [Fact]
    public void Add_run_code_table_and_image()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "one")), null);
        var r = Mutations.Add(p, "run", Props(("text", " two"), ("bold", "true")), null);
        Assert.Equal("/body/paragraph[1]/run[2]", r.Path);
        Assert.Equal("one two", Mutations.Refresh(p).Text);
        Assert.Throws<WriterException>(() => Mutations.Add(p, "run", Props(("bold", "true")), null));

        var code = Mutations.Add(body, "code", Props(("text", "a\nb")), null);
        Assert.Equal("code", code.Kind);
        Assert.Equal("a\nb", code.Text);

        var table = Mutations.Add(body, "table", Props(("data", "[[\"h1\",\"h2\"],[\"1\",\"2\"]]")), null);
        Assert.Equal(("2", "2"), (table.GetProps()["rows"], table.GetProps()["cols"]));
        Assert.Equal("[[\"h1\",\"h2\"],[\"1\",\"2\"]]", table.GetProps()["data"]);
        var t2 = Mutations.Add(body, "table", Props(("rows", "1"), ("cols", "3")), null);
        Assert.Equal("[[\"\",\"\",\"\"]]", t2.GetProps()["data"]);
        Assert.Throws<WriterException>(() => Mutations.Add(body, "table", Props(), null));

        var png = Path.Combine(Path.GetTempPath(), $"writer-{Guid.NewGuid():N}.png");
        File.WriteAllBytes(png, FakePng(2, 1));
        try
        {
            var img = Mutations.Add(body, "image", Props(("src", png), ("width", "4cm"), ("alt", "dot")), null);
            var shown = Registry.ToDisplay("image", img.GetProps());
            Assert.Equal("image", img.Kind);
            Assert.Equal("4cm", shown["width"]);
            Assert.Equal("2cm", shown["height"]);
            Assert.Equal("dot", shown["alt"]);
            Assert.Contains("/media/", shown["src"]);
            img = Mutations.Set(img, Props(("height", "3cm"), ("alt", "")));
            Assert.Equal("3cm", Registry.ToDisplay("image", img.GetProps())["height"]);
            Assert.False(img.GetProps().ContainsKey("alt"));
            var first = img.GetProps()["src"];
            img = Mutations.Set(img, Props(("src", png)));
            Assert.NotEqual(first, img.GetProps()["src"]);
            Assert.Throws<WriterException>(() => Mutations.Add(body, "image", Props(("src", png + ".missing")), null));
        }
        finally
        {
            File.Delete(png);
        }
    }

    [Fact]
    public void Remove_and_move_keep_the_document_valid()
    {
        using var doc = OpenDocx(Docx(P("A"), P("B"), new W.Table(new W.TableRow(Cell("x"), Cell("y")))));
        var body = doc.Root.Children.Single();
        PathResolver.Single(doc.Root, "/body/paragraph[1]").Remove();
        Assert.Equal(new[] { "B" }, body.Children.Where(c => c.Kind == "paragraph").Select(c => c.Text));

        var cell = PathResolver.Single(doc.Root, "//cell[1]");
        PathResolver.Single(doc.Root, "//cell[1]/paragraph[1]").Remove();
        Assert.Equal("", Mutations.Refresh(cell).Text);
        Assert.Contains("<w:p", Mutations.Refresh(cell).GetRaw());
        var ex = Assert.Throws<WriterException>(() => PathResolver.Single(doc.Root, "//row[1]").Remove());
        Assert.Contains("last row", ex.Message);
        PathResolver.Single(doc.Root, "//cell[2]").Remove();
        Assert.Throws<WriterException>(() => PathResolver.Single(doc.Root, "//cell[1]").Remove());

        var moved = Mutations.Move(PathResolver.Single(doc.Root, "/body/paragraph[1]"), PathResolver.Single(doc.Root, "//cell[1]"), null);
        Assert.Equal("/body/table[1]/row[1]/cell[1]/paragraph[2]", moved.Path);
        Assert.DoesNotContain(body.Children, c => c.Kind == "paragraph");
        var self = Assert.Throws<WriterException>(() =>
            Mutations.Move(PathResolver.Single(doc.Root, "/body/table[1]"), PathResolver.Single(doc.Root, "//cell[1]"), null));
        Assert.Contains("itself", self.Message);
    }

    [Fact]
    public void Table_data_rows_cols_and_cells()
    {
        using var doc = OpenDocx(Docx(new W.Table(new W.TableRow(Cell("a"), Cell("b")))));
        var t = PathResolver.Single(doc.Root, "/body/table[1]");
        t = Mutations.Set(t, Props(("data", "[[\"h1\",\"h2\",\"h3\"],[\"1\",\"2\",\"3\"]]")));
        Assert.Equal(("2", "3"), (t.GetProps()["rows"], t.GetProps()["cols"]));
        t = Mutations.Set(t, Props(("rows", "3"), ("cols", "2")));
        Assert.Equal("[[\"h1\",\"h2\"],[\"1\",\"2\"],[\"\",\"\"]]", t.GetProps()["data"]);
        var row = Mutations.Set(PathResolver.Single(doc.Root, "//row[3]"), Props(("data", "[\"x\",\"y\"]")));
        Assert.Equal("[\"x\",\"y\"]", row.GetProps()["data"]);
        var added = Mutations.Add(t, "row", Props(("data", "[\"p\",\"q\"]")), 1);
        Assert.Equal("/body/table[1]/row[1]", added.Path);
        Assert.Equal("4", Mutations.Refresh(t).GetProps()["rows"]);
        var cell = Mutations.Set(PathResolver.Single(doc.Root, "//row[1]/cell[1]"), Props(("md", "**P**"), ("fill", "D9E2F3"), ("align", "center")));
        Assert.Equal("P", cell.Text);
        Assert.Equal("true", cell.Children[0].Children[0].GetProps()["bold"]);
        Assert.Equal("D9E2F3", cell.GetProps()["fill"]);
        Assert.Equal("center", cell.GetProps()["align"]);
        var newCell = Mutations.Add(PathResolver.Single(doc.Root, "//row[1]"), "cell", Props(("text", "r")), null);
        Assert.Equal("/body/table[1]/row[1]/cell[3]", newCell.Path);
        Assert.Equal("3", Mutations.Refresh(t).GetProps()["cols"]);
    }

    [Fact]
    public void Raw_xml_replaces_elements_and_rejects_wrong_kinds()
    {
        using var doc = OpenDocx(Docx(P("old")));
        var p = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        p.SetRaw("<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>raw</w:t></w:r></w:p>");
        p = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        Assert.Equal("raw", p.Text);
        Assert.Equal("center", p.GetProps()["align"]);
        var wrong = Assert.Throws<WriterException>(() => p.SetRaw("<w:tbl/>"));
        Assert.Contains("w:p", wrong.Message);
        Assert.Throws<WriterException>(() => p.SetRaw("not xml"));
        PathResolver.Single(doc.Root, "/body/paragraph[1]/run[1]").SetRaw("<w:r><w:rPr><w:b/></w:rPr><w:t>bold raw</w:t></w:r>");
        Assert.Equal("true", PathResolver.Single(doc.Root, "/body/paragraph[1]/run[1]").GetProps()["bold"]);
    }

    [Fact]
    public void Document_title_is_writable()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, Props(("title", "Q4")));
        Assert.Equal("Q4", doc.Root.GetProps()["title"]);
    }

    [Fact]
    public void Editing_one_paragraph_changes_only_that_paragraph()
    {
        var original = File.ReadAllBytes(Path.Combine(FixtureDir("docx"), "paragraph-formatting.docx"));
        using var doc = new DocxAdapter().Open(new MemoryStream(original));
        var target = PathResolver.Query(doc.Root, "/body/paragraph").First(p => p.Text!.Length > 0);
        var blocks = ((DocxDocument)doc).Main.Document!.Body!.Elements().Where(e => e is W.Paragraph or W.Table).ToList();
        var index = blocks.IndexOf((OpenXmlElement)target.Anchor);
        Mutations.Set(target, Props(("text", "EDITED")));
        var saved = Save(doc);
        Assert.Null(PackageCompare.Diff(original, saved, "word/document.xml"));
        Assert.Null(PackageCompare.DiffBodyExcept(original, saved, index));
    }

    [Fact]
    public void Adding_a_builtin_style_touches_only_styles_and_document()
    {
        var original = File.ReadAllBytes(Path.Combine(FixtureDir("docx"), "tables.docx"));
        using var doc = new DocxAdapter().Open(new MemoryStream(original));
        var body = doc.Root.Children.Single();
        Mutations.Add(body, "paragraph", Props(("text", "quoted"), ("style", "Quote")), null);
        var saved = Save(doc);
        Assert.Null(PackageCompare.Diff(original, saved, "word/document.xml", "word/styles.xml"));
        Assert.NotEqual(PackageCompare.Canonical(PackageCompare.Parts(original)["word/styles.xml"]), PackageCompare.Canonical(PackageCompare.Parts(saved)["word/styles.xml"]));
    }
}
