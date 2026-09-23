using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using Writer.Formats.Html;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Merged cells and table formatting: gridSpan and vMerge behind colspan/rowspan, styles, borders, widths.</summary>
public class DocxTableTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static Document Table(string data)
    {
        var doc = new DocxAdapter().Create();
        Mutations.Add(doc.Root.Children.Single(), "table", Props(("data", data)), null);
        return doc;
    }

    static Node At(Document doc, string path) => PathResolver.Single(doc.Root, path);
    static string Data(Document doc) => At(doc, "/body/table[1]").GetProps()["data"];
    static int GridColumns(Document doc) => ((W.Table)At(doc, "/body/table[1]").Anchor).GetFirstChild<W.TableGrid>()!.Elements<W.GridColumn>().Count();

    static void AssertValid(Document doc)
    {
        var errors = new OpenXmlValidator().Validate(((DocxDocument)doc).Package)
            .Where(e => e.Part is MainDocumentPart or StyleDefinitionsPart)
            .Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void Colspan_absorbs_the_right_neighbour_and_splits_back_into_empty_cells()
    {
        using var doc = Table("[[\"a\",\"b\",\"c\"],[\"d\",\"e\",\"f\"]]");
        var cell = Mutations.Set(At(doc, "/body/table[1]/row[1]/cell[1]"), Props(("colspan", "2")));
        Assert.Equal(("2", "a\nb"), (cell.GetProps()["colspan"], cell.Text));
        Assert.Equal(2, cell.Children.Count);
        Assert.Contains("<w:gridSpan w:val=\"2\"", cell.GetRaw());
        Assert.Equal("[[\"a\\nb\",\"c\"],[\"d\",\"e\",\"f\"]]", Data(doc));
        Assert.Equal(("2", "3"), (At(doc, "/body/table[1]").GetProps()["rows"], At(doc, "/body/table[1]").GetProps()["cols"]));
        Assert.Equal(3, GridColumns(doc));
        Assert.Contains("colspan=\"2\"", HtmlWriter.Render(doc));

        cell = Mutations.Set(cell, Props(("colspan", "1")));
        Assert.False(cell.GetProps().ContainsKey("colspan"));
        Assert.Equal("[[\"a\\nb\",\"\",\"c\"],[\"d\",\"e\",\"f\"]]", Data(doc));
        Assert.Equal(3, GridColumns(doc));
        AssertValid(doc);
    }

    [Fact]
    public void Rowspan_absorbs_the_cell_below_hides_it_and_splits_back()
    {
        using var doc = Table("[[\"a\",\"b\"],[\"c\",\"d\"],[\"e\",\"f\"]]");
        var cell = Mutations.Set(At(doc, "/body/table[1]/row[2]/cell[1]"), Props(("rowspan", "2")));
        Assert.Equal(("2", "c\ne"), (cell.GetProps()["rowspan"], cell.Text));
        Assert.Contains("<w:vMerge w:val=\"restart\"", cell.GetRaw());
        Assert.Equal("[\"f\"]", At(doc, "/body/table[1]/row[3]").GetProps()["data"]);
        Assert.Single(At(doc, "/body/table[1]/row[3]").Children);
        Assert.Contains("<w:vMerge />", At(doc, "/body/table[1]/row[3]").GetRaw());
        Assert.Equal("[[\"a\",\"b\"],[\"c\\ne\",\"d\"],[\"f\"]]", Data(doc));
        Assert.Contains("rowspan=\"2\"", HtmlWriter.Render(doc));
        Assert.Contains("a\tb\nc e\td\nf\n", Views.Text(doc.Root));

        var ex = Assert.Throws<WriterException>(() => Mutations.Set(cell, Props(("rowspan", "3"))));
        Assert.Equal(ErrorCode.Validation, ex.Code);
        cell = Mutations.Set(cell, Props(("rowspan", "1")));
        Assert.False(cell.GetProps().ContainsKey("rowspan"));
        Assert.Equal("[[\"a\",\"b\"],[\"c\\ne\",\"d\"],[\"\",\"f\"]]", Data(doc));
        AssertValid(doc);
    }

    [Fact]
    public void Widening_a_vertical_merge_moves_every_neighbour_into_the_visible_cell()
    {
        using var doc = Table("[[\"a\",\"b\",\"c\"],[\"d\",\"e\",\"f\"],[\"g\",\"h\",\"i\"]]");
        Mutations.Set(At(doc, "/body/table[1]/row[2]/cell[1]"), Props(("rowspan", "2")));
        var cell = Mutations.Set(At(doc, "/body/table[1]/row[2]/cell[1]"), Props(("colspan", "2")));
        Assert.Equal("d\ng\ne\nh", cell.Text);
        Assert.Equal("[[\"a\",\"b\",\"c\"],[\"d\\ng\\ne\\nh\",\"f\"],[\"i\"]]", Data(doc));
        Assert.Equal(("2", "2"), (cell.GetProps()["colspan"], cell.GetProps()["rowspan"]));
        Assert.Contains("d g e h\tf", Views.Text(doc.Root));
        AssertValid(doc);
    }

    [Fact]
    public void Merging_down_over_a_narrower_cell_widens_it_and_over_a_wider_one_is_refused()
    {
        using var doc = Table("[[\"a\",\"b\",\"c\"],[\"d\",\"e\",\"f\"]]");
        Mutations.Set(At(doc, "/body/table[1]/row[1]/cell[1]"), Props(("colspan", "2")));
        var cell = Mutations.Set(At(doc, "/body/table[1]/row[1]/cell[1]"), Props(("rowspan", "2")));
        Assert.Equal("a\nb\nd\ne", cell.Text);
        Assert.Equal("[[\"a\\nb\\nd\\ne\",\"c\"],[\"f\"]]", Data(doc));
        var hidden = ((W.TableRow)At(doc, "/body/table[1]/row[2]").Anchor).GetFirstChild<W.TableCell>()!;
        Assert.Equal(2, hidden.TableCellProperties!.GridSpan!.Val!.Value);
        Assert.NotNull(hidden.TableCellProperties.VerticalMerge);
        AssertValid(doc);

        using var wider = Table("[[\"a\",\"b\",\"c\"],[\"d\",\"e\",\"f\"]]");
        Mutations.Set(At(wider, "/body/table[1]/row[2]/cell[1]"), Props(("colspan", "2")));
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(At(wider, "/body/table[1]/row[1]/cell[1]"), Props(("rowspan", "2"))));
        Assert.Equal(ErrorCode.Validation, ex.Code);
        Assert.Contains("colspan=2", ex.Hint);
        Assert.Equal("[[\"a\",\"b\",\"c\"],[\"d\\ne\",\"f\"]]", Data(wider));
    }

    [Fact]
    public void Removing_the_top_row_of_a_merge_hands_the_cell_to_the_row_below()
    {
        using var doc = Table("[[\"a\",\"b\"],[\"c\",\"d\"],[\"e\",\"f\"]]");
        Mutations.Set(At(doc, "/body/table[1]/row[1]/cell[1]"), Props(("rowspan", "3"), ("fill", "D9E2F3")));
        At(doc, "/body/table[1]/row[1]").Remove();
        var cell = At(doc, "/body/table[1]/row[1]/cell[1]");
        Assert.Equal(("a\nc\ne", "2", "D9E2F3"), (cell.Text, cell.GetProps()["rowspan"], cell.GetProps()["fill"]));
        Assert.Equal("[[\"a\\nc\\ne\",\"d\"],[\"f\"]]", Data(doc));
        At(doc, "/body/table[1]/row[1]").Remove();
        cell = At(doc, "/body/table[1]/row[1]/cell[1]");
        Assert.Equal("a\nc\ne", cell.Text);
        Assert.False(cell.GetProps().ContainsKey("rowspan"));
        Assert.Equal("[[\"a\\nc\\ne\",\"f\"]]", Data(doc));
        AssertValid(doc);
    }

    [Fact]
    public void Adding_a_row_inside_a_vertical_merge_extends_it_and_outside_leaves_it_alone()
    {
        using var doc = Table("[[\"a\",\"b\"],[\"c\",\"d\"],[\"e\",\"f\"]]");
        Mutations.Set(At(doc, "/body/table[1]/row[1]/cell[1]"), Props(("rowspan", "3")));
        var inside = Mutations.Add(At(doc, "/body/table[1]"), "row", Props(("data", "[\"x\"]")), 2);
        Assert.Equal("/body/table[1]/row[2]", inside.Path);
        Assert.Equal("4", At(doc, "/body/table[1]/row[1]/cell[1]").GetProps()["rowspan"]);
        Assert.Equal("[[\"a\\nc\\ne\",\"b\"],[\"x\"],[\"d\"],[\"f\"]]", Data(doc));

        var top = Mutations.Add(At(doc, "/body/table[1]"), "row", Props(("data", "[\"p\",\"q\"]")), 1);
        Assert.Equal("/body/table[1]/row[1]", top.Path);
        Assert.Equal("4", At(doc, "/body/table[1]/row[2]/cell[1]").GetProps()["rowspan"]);
        var bottom = Mutations.Add(At(doc, "/body/table[1]"), "row", Props(("data", "[\"y\",\"z\"]")), null);
        Assert.Equal("/body/table[1]/row[6]", bottom.Path);
        Assert.Equal("[[\"p\",\"q\"],[\"a\\nc\\ne\",\"b\"],[\"x\"],[\"d\"],[\"f\"],[\"y\",\"z\"]]", Data(doc));
        Assert.Equal(2, GridColumns(doc));
        AssertValid(doc);
    }

    [Fact]
    public void Data_keeps_merges_that_fit_and_rows_cols_count_the_grid()
    {
        using var doc = Table("[[\"a\",\"b\",\"c\"],[\"d\",\"e\",\"f\"]]");
        Mutations.Set(At(doc, "/body/table[1]/row[1]/cell[1]"), Props(("colspan", "2")));
        var t = Mutations.Set(At(doc, "/body/table[1]"), Props(("data", "[[\"X\",\"Y\"],[\"1\",\"2\",\"3\"]]")));
        Assert.Equal("[[\"X\",\"Y\"],[\"1\",\"2\",\"3\"]]", t.GetProps()["data"]);
        Assert.Equal(("2", "3"), (t.GetProps()["rows"], t.GetProps()["cols"]));
        Assert.Equal("2", At(doc, "/body/table[1]/row[1]/cell[1]").GetProps()["colspan"]);
        t = Mutations.Set(t, Props(("rows", "3")));
        Assert.Equal("[[\"X\",\"Y\"],[\"1\",\"2\",\"3\"],[\"\",\"\",\"\"]]", t.GetProps()["data"]);
        Assert.Equal(3, GridColumns(doc));
    }

    [Fact]
    public void Split_cells_keep_the_formatting_and_widening_into_a_merge_from_above_is_refused()
    {
        using var doc = Table("[[\"a\",\"b\",\"c\"],[\"d\",\"e\",\"f\"],[\"g\",\"h\",\"i\"]]");
        Mutations.Set(At(doc, "/body/table[1]/row[3]/cell[1]"), Props(("fill", "D9E2F3"), ("colspan", "2")));
        Mutations.Set(At(doc, "/body/table[1]/row[3]/cell[1]"), Props(("colspan", "1")));
        Assert.Equal("D9E2F3", At(doc, "/body/table[1]/row[3]/cell[2]").GetProps()["fill"]);

        Mutations.Set(At(doc, "/body/table[1]/row[1]/cell[2]"), Props(("rowspan", "3")));
        var before = Data(doc);
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(At(doc, "/body/table[1]/row[2]/cell[1]"), Props(("colspan", "2"))));
        Assert.Contains("row above", ex.Message);
        Assert.Equal(before, Data(doc));
        Assert.Equal("3", At(doc, "/body/table[1]/row[1]/cell[2]").GetProps()["rowspan"]);
        AssertValid(doc);
    }

    [Fact]
    public void Writing_a_cell_edits_its_paragraphs_in_place_so_fields_and_other_paragraphs_survive()
    {
        var field = new W.Paragraph(new W.Run(new W.Text("Page ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
            new W.SimpleField(new W.Run(new W.Text("3"))) { Instruction = " PAGE " });
        var grid = new W.TableGrid(new W.GridColumn { Width = "3000" }, new W.GridColumn { Width = "3000" });
        using var doc = OpenDocx(Docx(new W.Table(new W.TableProperties(), grid, new W.TableRow(new W.TableCell(field, P("second")), Cell("x")))));
        var cell = At(doc, "/body/table[1]/row[1]/cell[1]");
        cell = Mutations.Set(cell, Props(("html", cell.GetProps()["html"].Replace("second", "changed"))));
        Assert.Contains("w:fldSimple", cell.GetRaw());
        Assert.Equal(2, cell.Children.Count(c => c.Kind == "paragraph"));
        Assert.EndsWith("\nchanged", cell.Text);

        cell = Mutations.Set(cell, Props(("text", "one line"))); // fewer lines than paragraphs: the first one takes them
        Assert.Single(cell.Children, c => c.Kind == "paragraph");
        Assert.StartsWith("one line", cell.Text);
        AssertValid(doc);
    }

    [Fact]
    public void Column_widths_follow_the_cells_and_new_rows_are_not_header_rows()
    {
        using var doc = Table("[[\"a\",\"b\",\"c\"],[\"d\",\"e\",\"f\"]]");
        Mutations.Set(At(doc, "/body/table[1]"), Props(("widths", "[\"2cm\",\"5cm\",\"3cm\"]")));
        At(doc, "/body/table[1]/row[2]/cell[1]").Remove();
        At(doc, "/body/table[1]/row[1]/cell[1]").Remove();
        Assert.Equal("[\"5cm\",\"3cm\"]", At(doc, "/body/table[1]").GetProps()["widths"]);

        Mutations.Set(At(doc, "/body/table[1]/row[1]"), Props(("header", "true"), ("height", "1cm")));
        var row = Mutations.Add(At(doc, "/body/table[1]"), "row", Props(), 2);
        Assert.False(row.GetProps().ContainsKey("header"));
        Assert.Equal("1cm", Registry.ToDisplay("docx", "row", row.GetProps())["height"]);
        Assert.Equal("true", At(doc, "/body/table[1]/row[1]").GetProps()["header"]);
    }

    [Fact]
    public void Removing_a_column_cell_by_cell_keeps_merges_and_a_merged_cell_goes_whole()
    {
        using var doc = Table("[[\"a\",\"b\",\"c\"],[\"d\",\"e\",\"f\"],[\"g\",\"h\",\"i\"]]");
        Mutations.Set(At(doc, "/body/table[1]/row[1]/cell[2]"), Props(("rowspan", "2")));
        foreach (var path in new[] { "/body/table[1]/row[3]/cell[1]", "/body/table[1]/row[2]/cell[1]", "/body/table[1]/row[1]/cell[1]" })
            At(doc, path).Remove();
        Assert.Equal("[[\"b\\ne\",\"c\"],[\"f\"],[\"h\",\"i\"]]", Data(doc));
        Assert.Equal("2", At(doc, "/body/table[1]/row[1]/cell[1]").GetProps()["rowspan"]);
        Assert.Equal(2, GridColumns(doc));

        At(doc, "/body/table[1]/row[1]/cell[1]").Remove();
        Assert.Equal("[[\"c\"],[\"f\"],[\"h\",\"i\"]]", Data(doc));
        AssertValid(doc);

        using var ragged = Table("[[\"a\",\"b\"],[\"c\"]]");
        Mutations.Set(At(ragged, "/body/table[1]/row[1]/cell[1]"), Props(("rowspan", "2")));
        var ex = Assert.Throws<WriterException>(() => At(ragged, "/body/table[1]/row[1]/cell[1]").Remove());
        Assert.Contains("row below", ex.Message);
    }

    [Fact]
    public void Border_presets_write_the_lines_they_name_and_style_removes_the_override()
    {
        using var doc = Table("[[\"a\",\"b\"],[\"c\",\"d\"]]");
        var t = At(doc, "/body/table[1]");
        var pr = ((W.Table)t.Anchor).GetFirstChild<W.TableProperties>()!;
        t = Mutations.Set(t, Props(("borders", "horizontal")));
        var b = pr.TableBorders!;
        Assert.Equal((W.BorderValues.Single, W.BorderValues.Single, W.BorderValues.Single), (b.TopBorder!.Val!.Value, b.BottomBorder!.Val!.Value, b.InsideHorizontalBorder!.Val!.Value));
        Assert.Equal((W.BorderValues.None, W.BorderValues.None, W.BorderValues.None), (b.LeftBorder!.Val!.Value, b.RightBorder!.Val!.Value, b.InsideVerticalBorder!.Val!.Value));
        foreach (var kind in new[] { "none", "all", "outside", "inside", "horizontal" })
            Assert.Equal(kind, Mutations.Set(t, Props(("borders", kind))).GetProps()["borders"]);
        Mutations.Set(t, Props(("borders", "inside")));
        Assert.Matches("<table style=\"[^\"]*border-style:hidden;", HtmlWriter.Render(doc));

        t = Mutations.Set(t, Props(("style", "ThreeLineTable"), ("borders", "style")));
        Assert.False(t.GetProps().ContainsKey("borders"));
        Assert.Null(pr.TableBorders);
        var style = ((DocxDocument)doc).Main.StyleDefinitionsPart!.Styles!.Elements<W.Style>().Single(s => s.StyleId == "ThreeLineTable");
        Assert.Contains("w:tblStylePr w:type=\"firstRow\"", style.OuterXml);

        t = Mutations.Set(t, Props(("borderColor", "C00000")));
        Assert.Equal("C00000", t.GetProps()["borderColor"]); // the style's top and bottom rules, recolored, and nothing added
        Assert.Null(pr.TableBorders!.InsideHorizontalBorder);
        Assert.Null(pr.TableBorders.LeftBorder);
        AssertValid(doc);
    }

    [Fact]
    public void Built_in_styles_reuse_a_localized_documents_own_definitions()
    {
        static W.Style Style(W.StyleValues type, string id, string name, bool isDefault = false) =>
            new(new W.StyleName { Val = name }) { Type = type, StyleId = id, Default = isDefault ? true : null };
        var bytes = Docx([P("x", "1")], main => main.StyleDefinitionsPart!.Styles = new W.Styles(
            Style(W.StyleValues.Paragraph, "a", "Normal", true), Style(W.StyleValues.Paragraph, "1", "heading 1"),
            Style(W.StyleValues.Table, "a3", "Table Grid"), Style(W.StyleValues.Paragraph, "10", "toc 1")));
        using var doc = OpenDocx(bytes);
        var body = doc.Root.Children.Single();
        var table = Mutations.Add(body, "table", Props(("rows", "1"), ("cols", "1")), null);
        Assert.Equal("a3", table.GetProps()["style"]);
        Mutations.Add(body, "toc", Props(("title", "目录")), 1);
        var styles = ((DocxDocument)doc).Main.StyleDefinitionsPart!.Styles!.Elements<W.Style>().ToList();
        Assert.Single(styles, s => s.StyleName?.Val?.Value == "Normal");
        Assert.Single(styles, s => s.StyleName?.Val?.Value == "toc 1");
        Assert.Single(styles, s => s.StyleName?.Val?.Value == "Table Grid");
        Assert.Equal("1", styles.Single(s => s.StyleId == "TOCHeading").BasedOn!.Val!.Value);
        Assert.Contains("w:pStyle w:val=\"10\"", PathResolver.Single(doc.Root, "/body/toc[1]").GetRaw());
        var wrong = Assert.Throws<WriterException>(() => Mutations.Set(table, Props(("style", "Quote"))));
        Assert.Contains("paragraph style", wrong.Message);
    }

    [Fact]
    public void Table_row_and_cell_formatting_round_trips_and_reaches_the_html()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var t = Mutations.Add(body, "table", Props(("rows", "2"), ("cols", "2"), ("style", "Plain Table 1"), ("borders", "horizontal"), ("borderColor", "FF0000"), ("width", "80%"), ("align", "center")), null);
        var props = t.GetProps();
        Assert.Equal(("PlainTable1", "horizontal", "FF0000", "80%", "center"), (props["style"], props["borders"], props["borderColor"], props["width"], props["align"]));
        Assert.Contains(((DocxDocument)doc).Main.StyleDefinitionsPart!.Styles!.Elements<W.Style>(), s => s.StyleId == "PlainTable1");
        t = Mutations.Set(t, Props(("widths", "[\"3cm\",\"5cm\"]"), ("borders", "none")));
        Assert.Equal("[\"3cm\",\"5cm\"]", t.GetProps()["widths"]);
        Assert.Equal("8cm", t.GetProps()["width"]); // 1701 + 2835 twips, within a twip of 8cm
        Assert.Equal("none", t.GetProps()["borders"]);
        Assert.Contains("w:tblLayout w:type=\"fixed\"", t.GetRaw());
        Assert.Equal("5cm", Registry.ToDisplay("docx", "cell", At(doc, "/body/table[1]/row[1]/cell[2]").GetProps())["width"]);
        Assert.Throws<WriterException>(() => Mutations.Set(t, Props(("style", "NoSuchStyle"))));
        t = Mutations.Set(t, Props(("style", "")));
        Assert.False(t.GetProps().ContainsKey("style"));

        var row = Mutations.Set(At(doc, "/body/table[1]/row[1]"), Props(("header", "true"), ("height", "1cm")));
        Assert.Equal(("true", "1cm"), (row.GetProps()["header"], Registry.ToDisplay("docx", "row", row.GetProps())["height"]));
        Assert.Contains("<w:tblHeader />", row.GetRaw());
        row = Mutations.Set(row, Props(("header", "false"), ("height", "0")));
        Assert.False(row.GetProps().ContainsKey("header"));
        Assert.False(row.GetProps().ContainsKey("height"));
        Assert.DoesNotContain("<w:trPr", row.GetRaw());

        var cell = Mutations.Set(At(doc, "/body/table[1]/row[1]/cell[1]"), Props(("fill", "D9E2F3"), ("borders", "outside"), ("valign", "middle"), ("width", "4cm"), ("colspan", "2")));
        var shown = Registry.ToDisplay("docx", "cell", cell.GetProps());
        Assert.Equal(("D9E2F3", "outside", "middle", "2"), (shown["fill"], shown["borders"], shown["valign"], shown["colspan"]));
        Assert.Contains("w:tcBorders", cell.GetRaw());
        cell = Mutations.Set(cell, Props(("fill", "none")));
        Assert.False(cell.GetProps().ContainsKey("fill"));

        var html = HtmlWriter.Render(doc);
        Assert.Contains("<table style=\"width:302.4px;margin-left:auto;margin-right:auto;\">", html);
        Assert.Contains("<td colspan=\"2\" style=\"border:none;border:1px solid #bbb;vertical-align:middle;width:", html);
        AssertValid(doc);
    }

    [Fact]
    public void Merged_word_table_reads_visible_cells_and_survives_an_edit()
    {
        var original = File.ReadAllBytes(Path.Combine(FixtureDir("docx"), "tables.docx"));
        using var doc = new DocxAdapter().Open(new MemoryStream(original));
        var t = At(doc, "/body/table[2]");
        Assert.Equal(("8", "5"), (t.GetProps()["rows"], t.GetProps()["cols"]));
        Assert.Equal("[\"Category\",\"Line Item\",\"Amount (10K USD)\",\"Notes\"]", At(doc, "/body/table[2]/row[1]").GetProps()["data"]);
        Assert.Equal("2", At(doc, "/body/table[2]/row[1]/cell[3]").GetProps()["colspan"]);
        Assert.Equal("2", At(doc, "/body/table[2]/row[1]/cell[1]").GetProps()["rowspan"]);
        Assert.Equal("[\"Budget\",\"Actual\"]", At(doc, "/body/table[2]/row[2]").GetProps()["data"]);
        Assert.Contains("<td colspan=\"2\"", HtmlWriter.Render(doc));

        Mutations.Set(At(doc, "/body/table[2]/row[3]/cell[1]"), Props(("text", "Sales")));
        var saved = new MemoryStream();
        doc.Save(saved);
        Assert.Null(PackageCompare.Diff(original, saved.ToArray(), "word/document.xml"));
        using var reopened = OpenDocx(saved.ToArray());
        Assert.Equal("Sales", At(reopened, "/body/table[2]/row[3]/cell[1]").Text);
        Assert.Equal("2", At(reopened, "/body/table[2]/row[1]/cell[1]").GetProps()["rowspan"]);
    }
}
