using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Cli;
using Writer.Core;
using Writer.Formats.Xlsx;

namespace Writer.Tests;

/// <summary>Cell formatting, links and notes: every property round-trips, writes keep the rest of a cell's style,
/// and untouched parts stay byte-identical.</summary>
public class XlsxFormattingTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static byte[] Fixture => File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("xlsx"), "cell-formatting.xlsx"));

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    static Document Reopen(Document doc) => new XlsxAdapter().Open(new MemoryStream(Save(doc)));

    static IReadOnlyDictionary<string, string> Get(Document doc, string path) => PathResolver.Single(doc.Root, path).GetProps();

    static Node Set(Document doc, string path, params (string Name, string Value)[] pairs) => Mutations.Set(PathResolver.Single(doc.Root, path), Props(pairs));

    static Stylesheet Styles(Document doc) => ((XlsxDocument)doc).Workbook.WorkbookStylesPart!.Stylesheet!;

    static (int Code, string Out, string Err) Run(params string[] argv)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var code = Runner.Run(argv, stdout, stderr);
        return (code, stdout.ToString(), stderr.ToString());
    }

    [Fact]
    public void Every_formatting_prop_round_trips_on_a_blank_workbook()
    {
        using var doc = new XlsxAdapter().Create();
        Set(doc, "/sheet[1]/cell[B2]", ("value", "Total"), ("bold", "true"), ("italic", "true"), ("underline", "true"), ("strike", "true"), ("size", "14pt"),
            ("font", "Arial"), ("color", "C00000"), ("fill", "FFF2CC"), ("format", "0.00"), ("align", "center"), ("valign", "middle"), ("wrap", "true"), ("indent", "2"),
            ("border", "medium"), ("borderColor", "1F4E79"), ("link", "https://example.com/report"), ("note", "Reviewed by Ann"));
        using var reopened = Reopen(doc);
        var props = Get(reopened, "/sheet[1]/cell[B2]");
        Assert.Equal("Total", props["value"]);
        foreach (var flag in new[] { "bold", "italic", "underline", "strike", "wrap" }) Assert.Equal("true", props[flag]);
        Assert.Equal(("14", "Arial", "C00000", "FFF2CC", "0.00"), (props["size"], props["font"], props["color"], props["fill"], props["format"]));
        Assert.Equal(("center", "middle", "2"), (props["align"], props["valign"], props["indent"]));
        Assert.Equal(("medium", "1F4E79"), (props["border"], props["borderColor"]));
        Assert.False(props.ContainsKey("borders"));
        Assert.Equal(("https://example.com/report", "Reviewed by Ann"), (props["link"], props["note"]));
        Assert.Equal("14pt", Registry.ToDisplay("cell", props)["size"]);

        Set(reopened, "/sheet[1]/cell[B2]", ("bold", "false"), ("italic", "false"), ("underline", "false"), ("strike", "false"), ("wrap", "false"), ("indent", "0"),
            ("border", "none"), ("fill", "none"), ("link", ""), ("note", ""));
        using var cleared = Reopen(reopened);
        var after = Get(cleared, "/sheet[1]/cell[B2]");
        foreach (var name in new[] { "bold", "italic", "underline", "strike", "wrap", "indent", "border", "borderColor", "fill", "link", "note" })
            Assert.False(after.ContainsKey(name), name);
        Assert.Equal(("14", "Arial", "C00000", "center", "middle", "0.00"), (after["size"], after["font"], after["color"], after["align"], after["valign"], after["format"]));
    }

    [Fact]
    public void Fixture_cells_expose_fonts_fills_borders_alignment_and_links()
    {
        using var doc = new XlsxAdapter().Open(new MemoryStream(Fixture));
        Assert.Equal("Georgia", Get(doc, "/sheet[1]/cell[B3]")["font"]);
        Assert.Equal("18", Get(doc, "/sheet[1]/cell[B4]")["size"]);
        Assert.Equal("true", Get(doc, "/sheet[1]/cell[B6]")["italic"]);
        Assert.Equal("true", Get(doc, "/sheet[1]/cell[B8]")["underline"]);
        Assert.Equal("true", Get(doc, "/sheet[1]/cell[B9]")["underline"]);
        Assert.Equal("true", Get(doc, "/sheet[1]/cell[B10]")["strike"]);
        var combined = Get(doc, "/sheet[1]/cell[B13]");
        Assert.Equal(("true", "true", "14", "2E75B6"), (combined["bold"], combined["italic"], combined["size"], combined["color"]));
        var plainBold = Get(doc, "/sheet[1]/cell[B5]");
        Assert.Equal("true", plainBold["bold"]);
        Assert.False(plainBold.ContainsKey("size"));
        Assert.False(plainBold.ContainsKey("font"));

        Assert.Equal("E63946", Get(doc, "/sheet[Fills]/cell[A2]")["fill"]);
        Assert.Equal("left", Get(doc, "/sheet[Fills]/cell[A6]")["align"]);
        Assert.Equal("center", Get(doc, "/sheet[Fills]/cell[A7]")["align"]);
        Assert.Equal("right", Get(doc, "/sheet[Fills]/cell[A8]")["align"]);
        Assert.Equal("top", Get(doc, "/sheet[Fills]/cell[C6]")["valign"]);
        Assert.Equal("middle", Get(doc, "/sheet[Fills]/cell[C7]")["valign"]);
        Assert.Equal("bottom", Get(doc, "/sheet[Fills]/cell[C8]")["valign"]);
        Assert.Equal("true", Get(doc, "/sheet[Fills]/cell[A10]")["wrap"]);
        Assert.Equal("3", Get(doc, "/sheet[Fills]/cell[A16]")["indent"]);

        var thin = Get(doc, "/sheet[Borders]/cell[B3]");
        Assert.Equal("thin", thin["border"]);
        Assert.False(thin.ContainsKey("borders"));
        Assert.False(thin.ContainsKey("borderColor"));
        var colored = Get(doc, "/sheet[Borders]/cell[B7]");
        Assert.Equal(("thick", "C00000"), (colored["border"], colored["borderColor"]));
        var bottomOnly = Get(doc, "/sheet[Borders]/cell[B9]");
        Assert.Equal(("double", "{\"bottom\":\"double\"}"), (bottomOnly["border"], bottomOnly["borders"]));
        var mixed = Get(doc, "/sheet[Borders]/cell[B13]");
        Assert.Equal("{\"top\":\"thin\",\"right\":\"medium\",\"bottom\":\"double\",\"left\":\"thick\"}", mixed["borders"]);
        Assert.False(Get(doc, "/sheet[Borders]/cell[B15]").ContainsKey("border"));

        Assert.Equal("https://github.com/iOfficeAI/OfficeCLI", Get(doc, "/sheet[Data]/cell[A9]")["link"]);
        Assert.Equal("$#,##0.00", Get(doc, "/sheet[Numbers]/cell[B6]")["format"]);
    }

    [Fact]
    public void Setting_one_property_keeps_the_others_and_touches_only_styles_and_the_sheet()
    {
        var original = Fixture;
        using var doc = new XlsxAdapter().Open(new MemoryStream(original));
        var xfs = Styles(doc).CellFormats!.Elements<CellFormat>().Count();
        Set(doc, "/sheet[1]/cell[B13]", ("size", "20pt"));
        Set(doc, "/sheet[Borders]/cell[B7]", ("fill", "FFF2CC"));
        Set(doc, "/sheet[Numbers]/cell[B6]", ("bold", "true"));
        var b13 = Get(doc, "/sheet[1]/cell[B13]");
        Assert.Equal(("true", "true", "20", "2E75B6"), (b13["bold"], b13["italic"], b13["size"], b13["color"]));
        var b7 = Get(doc, "/sheet[Borders]/cell[B7]");
        Assert.Equal(("thick", "C00000", "FFF2CC"), (b7["border"], b7["borderColor"], b7["fill"]));
        var b6 = Get(doc, "/sheet[Numbers]/cell[B6]");
        Assert.Equal(("true", "$#,##0.00", "number"), (b6["bold"], b6["format"], b6["type"]));
        Assert.Equal(xfs + 2, Styles(doc).CellFormats!.Elements<CellFormat>().Count());
        Assert.Contains("s=\"48\"", PathResolver.Single(doc.Root, "/sheet[Numbers]/cell[B6]").GetRaw());
        Assert.Null(PackageCompare.Diff(original, Save(doc), "xl/styles.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet3.xml", "xl/worksheets/sheet4.xml"));
    }

    [Fact]
    public void Border_writes_cover_all_sides_single_sides_and_color_in_any_order()
    {
        using var doc = new XlsxAdapter().Create();
        var a1 = Set(doc, "/sheet[1]/cell[A1]", ("border", "thin"), ("borderColor", "FF0000")).GetProps();
        Assert.Equal(("thin", "FF0000"), (a1["border"], a1["borderColor"]));
        var a2 = Set(doc, "/sheet[1]/cell[A2]", ("borderColor", "00FF00"), ("border", "thick")).GetProps();
        Assert.Equal(("thick", "00FF00"), (a2["border"], a2["borderColor"]));

        var a3 = Set(doc, "/sheet[1]/cell[A3]", ("borders", "{\"bottom\":\"medium\",\"top\":\"thin\"}")).GetProps();
        Assert.Equal(("thin", "{\"top\":\"thin\",\"bottom\":\"medium\"}"), (a3["border"], a3["borders"]));
        a3 = Set(doc, "/sheet[1]/cell[A3]", ("borders", "{\"left\":\"dashed\"}")).GetProps();
        Assert.Equal("{\"top\":\"thin\",\"bottom\":\"medium\",\"left\":\"dashed\"}", a3["borders"]);
        Assert.False(a3.ContainsKey("borderColor"));
        a3 = Set(doc, "/sheet[1]/cell[A3]", ("borderColor", "0000FF")).GetProps();
        Assert.Equal("0000FF", a3["borderColor"]);
        a3 = Set(doc, "/sheet[1]/cell[A3]", ("borders", "{\"top\":\"none\",\"bottom\":\"none\",\"left\":\"none\"}")).GetProps();
        Assert.False(a3.ContainsKey("border"));

        var cleared = Set(doc, "/sheet[1]/cell[A1]", ("border", "none")).GetProps();
        Assert.False(cleared.ContainsKey("border"));
        Assert.False(cleared.ContainsKey("borderColor"));
        using var reopened = Reopen(doc);
        Assert.Equal(("thick", "00FF00"), (Get(reopened, "/sheet[1]/cell[A2]")["border"], Get(reopened, "/sheet[1]/cell[A2]")["borderColor"]));

        var side = Assert.Throws<WriterException>(() => Set(doc, "/sheet[1]/cell[A4]", ("borders", "{\"middle\":\"thin\"}")));
        Assert.Equal(ErrorCode.Validation, side.Code);
        var style = Assert.Throws<WriterException>(() => Set(doc, "/sheet[1]/cell[A4]", ("borders", "{\"top\":\"wavy\"}")));
        Assert.Contains("thin", style.Hint);
    }

    [Fact]
    public void Range_formatting_applies_to_every_cell_and_reuses_identical_styles()
    {
        using var doc = new XlsxAdapter().Create();
        Set(doc, "/sheet[1]/range[A1:D1]", ("values", "[[\"Region\",\"Q1\",\"Q2\",\"Q3\"]]"));
        var range = Set(doc, "/sheet[1]/range[A1:D1]", ("bold", "true"), ("fill", "D9E2F3"), ("align", "center"), ("border", "thin"), ("borderColor", "999999"));
        foreach (var reference in new[] { "A1", "B1", "C1", "D1" })
        {
            var cell = Get(doc, $"/sheet[1]/cell[{reference}]");
            Assert.Equal(("true", "D9E2F3", "center", "thin", "999999"), (cell["bold"], cell["fill"], cell["align"], cell["border"], cell["borderColor"]));
        }
        var shared = range.GetProps();
        Assert.Equal(("true", "D9E2F3", "center", "thin"), (shared["bold"], shared["fill"], shared["align"], shared["border"]));
        Assert.Equal("[[\"Region\",\"Q1\",\"Q2\",\"Q3\"]]", shared["values"]);

        var styles = Styles(doc);
        Assert.Equal(2, styles.CellFormats!.Elements<CellFormat>().Count()); // the five props make one format, not one each
        Assert.Equal(2u, styles.CellFormats.Count!.Value);
        Assert.Equal(2, styles.Fonts!.Elements<Font>().Count());
        Assert.Equal(3, styles.Fills!.Elements<Fill>().Count());
        Assert.Equal(2, styles.Borders!.Elements<Border>().Count());
        Set(doc, "/sheet[1]/range[A1:D1]", ("bold", "true"), ("fill", "D9E2F3"));
        Assert.Equal(2, styles.CellFormats.Elements<CellFormat>().Count());

        Set(doc, "/sheet[1]/cell[B1]", ("italic", "true"));
        shared = Get(doc, "/sheet[1]/range[A1:D1]");
        Assert.Equal("true", shared["bold"]);
        Assert.False(shared.ContainsKey("italic"));
        Assert.Equal(3, styles.CellFormats.Elements<CellFormat>().Count());
        Assert.Equal(3, styles.Fonts.Elements<Font>().Count());
    }

    [Fact]
    public void Links_point_to_urls_or_places_in_the_workbook()
    {
        var original = Fixture;
        using var doc = new XlsxAdapter().Open(new MemoryStream(original));
        Set(doc, "/sheet[1]/cell[A3]", ("link", "https://example.com/docs"));
        Set(doc, "/sheet[1]/cell[A4]", ("link", "#Numbers!B3"));
        using var reopened = Reopen(doc);
        Assert.Equal("https://example.com/docs", Get(reopened, "/sheet[1]/cell[A3]")["link"]);
        Assert.Equal("#Numbers!B3", Get(reopened, "/sheet[1]/cell[A4]")["link"]);
        Assert.Null(PackageCompare.Diff(original, Save(reopened), "xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels"));
        var sheet1 = ((XlsxDocument)reopened).Sheets[0].Part;
        Assert.Single(sheet1.HyperlinkRelationships);

        Set(reopened, "/sheet[1]/cell[A3]", ("link", ""));
        Set(reopened, "/sheet[1]/cell[A4]", ("link", ""));
        Assert.False(Get(reopened, "/sheet[1]/cell[A3]").ContainsKey("link"));
        Assert.Empty(sheet1.HyperlinkRelationships);
        Assert.Null(PackageCompare.Diff(original, Save(reopened), "xl/worksheets/_rels/sheet1.xml.rels"));

        var data = ((XlsxDocument)reopened).Sheets[4].Part;
        Set(reopened, "/sheet[Data]/cell[A9]", ("link", "https://example.com/new"));
        Assert.Equal("https://example.com/new", Get(reopened, "/sheet[Data]/cell[A9]")["link"]);
        Assert.Single(data.HyperlinkRelationships);
        Assert.Throws<WriterException>(() => Set(reopened, "/sheet[Data]/cell[A9]", ("link", "http://bad host/")));
    }

    [Fact]
    public void Notes_create_update_and_remove_with_the_parts_excel_needs()
    {
        var original = Fixture;
        using var doc = new XlsxAdapter().Open(new MemoryStream(original));
        Set(doc, "/sheet[1]/cell[B3]", ("note", "Check this font"));
        Set(doc, "/sheet[1]/cell[D8]", ("note", "Note on an empty cell"));
        var saved = Save(doc);
        var parts = PackageCompare.Parts(saved);
        Assert.Contains("xl/comments1.xml", parts.Keys);
        var vmlName = Assert.Single(parts.Keys, k => k.EndsWith(".vml"));
        var vml = Encoding.UTF8.GetString(parts[vmlName]);
        Assert.Equal(2, vml.Split("ObjectType=\"Note\"").Length - 1);
        Assert.Contains("<x:Row>2</x:Row><x:Column>1</x:Column>", vml);
        Assert.Contains("legacyDrawing", Encoding.UTF8.GetString(parts["xl/worksheets/sheet1.xml"]));
        Assert.Contains("vmlDrawing", Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
        Assert.Null(PackageCompare.Diff(original, saved, "xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels", "xl/comments1.xml", vmlName, "[Content_Types].xml"));

        using var reopened = new XlsxAdapter().Open(new MemoryStream(saved));
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        var sheet1 = ((XlsxDocument)reopened).Sheets[0].Part;
        Assert.Empty(validator.Validate(sheet1.WorksheetCommentsPart!));
        Assert.DoesNotContain(validator.Validate(sheet1), e => e.Part == sheet1);
        Assert.Equal("Check this font", Get(reopened, "/sheet[1]/cell[B3]")["note"]);
        Assert.Equal("Note on an empty cell", Get(reopened, "/sheet[1]/cell[D8]")["note"]);
        Set(reopened, "/sheet[1]/cell[B3]", ("note", "Font checked"));
        Assert.Equal("Font checked", Get(reopened, "/sheet[1]/cell[B3]")["note"]);
        var updated = PackageCompare.Parts(Save(reopened));
        Assert.Equal(2, Encoding.UTF8.GetString(updated[vmlName]).Split("ObjectType=\"Note\"").Length - 1);
        Assert.Contains("Font checked", Encoding.UTF8.GetString(updated["xl/comments1.xml"]));

        Set(reopened, "/sheet[1]/cell[B3]", ("note", ""));
        PathResolver.Single(reopened.Root, "/sheet[1]/cell[D8]").Remove();
        Assert.False(Get(reopened, "/sheet[1]/cell[B3]").ContainsKey("note"));
        var cleared = Save(reopened);
        Assert.DoesNotContain(PackageCompare.Parts(cleared).Keys, k => k.Contains("comments") || k.EndsWith(".vml"));
        Assert.Null(PackageCompare.Diff(original, cleared, "xl/worksheets/_rels/sheet1.xml.rels", "[Content_Types].xml"));
    }

    [Fact]
    public void Reading_a_workbook_without_styles_creates_none()
    {
        var ms = new MemoryStream();
        using (var package = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbook = package.AddWorkbookPart();
            workbook.Workbook = new Workbook(new Sheets());
            var sheet = workbook.AddNewPart<WorksheetPart>();
            sheet.Worksheet = new Worksheet(new SheetData(new Row(new Cell(new CellValue("1")) { CellReference = "A1" }) { RowIndex = 1U }));
            workbook.Workbook.Sheets!.Append(new Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1U, Name = "Bare" });
        }
        using var doc = new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));
        Assert.Equal("1", Get(doc, "/sheet[1]/cell[A1]")["value"]);
        Assert.Equal("[[\"1\"]]", Get(doc, "/sheet[1]/range[A1]")["values"]);
        _ = NodeJson.Serialize(doc.Root, int.MaxValue);
        Assert.Null(((XlsxDocument)doc).Workbook.WorkbookStylesPart);
        Set(doc, "/sheet[1]/cell[A1]", ("italic", "true"), ("border", "hair"));
        var props = Get(doc, "/sheet[1]/cell[A1]");
        Assert.Equal(("true", "hair"), (props["italic"], props["border"]));
    }

    [Fact]
    public void Help_lists_formatting_and_cli_examples_work()
    {
        var (code, help, _) = Run("help", "xlsx", "cell");
        Assert.Equal(0, code);
        foreach (var name in new[] { "italic", "underline", "strike", "size", "font", "align", "valign", "wrap", "indent", "border", "borders", "borderColor", "link", "note" })
            Assert.Contains($"  {name} ", help);
        Assert.Contains("--prop size=\"12pt\"", help);
        Assert.Contains("none | thin | medium | thick | dashed | dotted | double | hair", help);
        Assert.Contains("  bold ", Run("help", "xlsx", "range").Out);

        var dir = Directory.CreateTempSubdirectory("writer-xlsx-format").FullName;
        try
        {
            var file = Path.Combine(dir, "report.xlsx");
            Assert.Equal(0, Run("create", file).Code);
            Assert.Equal(0, Run("set", file, "/sheet[1]/range[A1:D1]", "--prop", "values=[[\"Region\",\"Q1\",\"Q2\",\"Q3\"]]").Code);
            Assert.Equal(0, Run("set", file, "/sheet[1]/range[A1:D1]", "--prop", "bold=true", "--prop", "fill=D9E2F3", "--prop", "border=thin").Code);
            Assert.Equal(0, Run("set", file, "/sheet[1]/cell[B3]", "--prop", "value=1234.5", "--prop", "size=12pt", "--prop", "align=right", "--prop", "format=#,##0.00").Code);
            Assert.Equal(0, Run("set", file, "/sheet[1]/cell[A3]", "--prop", "link=https://example.com", "--prop", "note=Source: finance").Code);
            Assert.Equal(0, Run("set", file, "/sheet[1]/cell[A4]", "--prop", "borders={\"bottom\":\"double\"}", "--prop", "borderColor=FF0000").Code);

            JsonElement PropsOf(string path) => JsonDocument.Parse(Run("get", file, path).Out).RootElement.GetProperty("props");
            var b3 = PropsOf("/sheet[1]/cell[B3]");
            Assert.Equal(("12pt", "right", "#,##0.00"), (b3.GetProperty("size").GetString(), b3.GetProperty("align").GetString(), b3.GetProperty("format").GetString()));
            var a1 = PropsOf("/sheet[1]/cell[A1]");
            Assert.Equal(("true", "D9E2F3", "thin"), (a1.GetProperty("bold").GetString(), a1.GetProperty("fill").GetString(), a1.GetProperty("border").GetString()));
            var a3 = PropsOf("/sheet[1]/cell[A3]");
            Assert.Equal(("https://example.com", "Source: finance"), (a3.GetProperty("link").GetString(), a3.GetProperty("note").GetString()));
            var a4 = PropsOf("/sheet[1]/cell[A4]");
            Assert.Equal("double", a4.GetProperty("borders").GetProperty("bottom").GetString());
            Assert.Equal("FF0000", a4.GetProperty("borderColor").GetString());
            var bad = Run("set", file, "/sheet[1]/cell[A4]", "--prop", "valign=centre");
            Assert.NotEqual(0, bad.Code);
            Assert.Contains("VALIDATION", bad.Err);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact]
    public void Styles_xml_does_not_grow_with_edits_that_repeat_or_come_back()
    {
        static (int Xfs, int Fonts, int Fills, int Borders, int Dxfs) Counts(Document d)
        {
            var st = Styles(d);
            return (st.CellFormats!.Count(), st.Fonts!.Count(), st.Fills!.Count(), st.Borders!.Count(), st.DifferentialFormats?.Count() ?? 0);
        }
        static uint Index(Document d) => ((Cell)PathResolver.Single(d.Root, "/sheet[1]/cell[A9]").Anchor).StyleIndex?.Value ?? 0;
        Document Cycle(Document d, string path, params (string Name, string Value)[] props)
        {
            Set(d, path, props);
            var next = Reopen(d);
            d.Dispose();
            return next;
        }
        var doc = new XlsxAdapter().Open(new MemoryStream(File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("xlsx"), "budget-tracker.xlsx"))));
        var before = Counts(doc);
        var index = Index(doc);
        var cf = Get(doc, "/sheet[1]")["cf"];
        for (var round = 0; round < 2; round++)
        {
            doc = Cycle(doc, "/sheet[1]/cell[A9]", ("bold", "true"), ("italic", "true"), ("fill", "FFF2CC"), ("border", "medium"), ("format", "0.0"));
            var styled = Counts(doc);
            Assert.True(styled.Xfs == before.Xfs + 1 && styled.Fonts == before.Fonts + 1 && styled.Fills <= before.Fills + 1 && styled.Borders <= before.Borders + 1, $"{before} → {styled}"); // one each at most, no step in between
            doc = Cycle(doc, "/sheet[1]/cell[A9]", ("bold", "false"), ("italic", "false"), ("fill", "none"), ("border", "thin"), ("format", "General"));
            Assert.Equal(index, Index(doc)); // back to the cell's own format
            Assert.Equal(before, Counts(doc)); // and what the edit added is gone
        }
        doc = Cycle(doc, "/sheet[1]", ("cf", cf)); // the editor writes the rules back as it read them
        Assert.Equal(before, Counts(doc));
        Assert.Equal(cf, Get(doc, "/sheet[1]")["cf"]);
        doc.Dispose();
    }
}
