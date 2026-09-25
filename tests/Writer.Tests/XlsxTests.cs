using Writer.Core;
using Writer.Formats.Xlsx;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace Writer.Tests;

public class XlsxTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    [Fact]
    public void Cells_are_typed_from_their_values()
    {
        using var doc = new XlsxAdapter().Create();
        Assert.Equal("1", doc.Root.GetProps()["sheets"]);
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[A1]"), Props(("value", "Name")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[Sheet1]/cell[b1]"), Props(("value", "42.5")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[C1]"), Props(("value", "true")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[D1]"), Props(("value", "2025-01-05")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[E1]"), Props(("value", "007"), ("type", "string")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[F1]"), Props(("formula", "=B1*2")));

        using var reopened = new XlsxAdapter().Open(new MemoryStream(Save(doc)));
        string Prop(string cell, string prop) => PathResolver.Single(reopened.Root, $"/sheet[1]/cell[{cell}]").GetProps().GetValueOrDefault(prop) ?? "";
        Assert.Equal(("Name", "string"), (Prop("A1", "value"), Prop("A1", "type")));
        Assert.Equal(("42.5", "number"), (Prop("B1", "value"), Prop("B1", "type")));
        Assert.Equal(("true", "bool"), (Prop("C1", "value"), Prop("C1", "type")));
        Assert.Equal(("2025-01-05", "date", "yyyy-mm-dd"), (Prop("D1", "value"), Prop("D1", "type"), Prop("D1", "format")));
        Assert.Equal(("007", "string"), (Prop("E1", "value"), Prop("E1", "type")));
        Assert.Equal("B1*2", Prop("F1", "formula"));
        Assert.Equal("A1:F1", reopened.Root.Children[0].GetProps()["range"]);
        Assert.Equal("", Prop("Z9", "value"));
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(PathResolver.Single(reopened.Root, "/sheet[1]/cell[A1]"), Props(("type", "number"))));
        Assert.Equal(ErrorCode.Validation, ex.Code);
    }

    [Fact]
    public void Rows_ranges_and_paths()
    {
        using var doc = new XlsxAdapter().Create();
        var sheet = doc.Root.Children.Single();
        Assert.Equal("/sheet[1]", sheet.Path);
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/range[A1:C2]"), Props(("values", "[[\"Name\",\"Score\",\"Pass\"],[\"Ann\",90,true]]")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/range[A3]"), Props(("values", "Bob,85,false\nCy,\"70,5\",true\n")));
        Assert.Equal("[[\"Name\",\"Score\",\"Pass\"],[\"Ann\",\"90\",\"true\"],[\"Bob\",\"85\",\"false\"],[\"Cy\",\"70,5\",\"true\"]]",
            PathResolver.Single(doc.Root, "/sheet[1]/range[A1:C4]").GetProps()["values"]);
        Assert.Equal("[\"Ann\",\"90\",\"true\"]", PathResolver.Single(doc.Root, "/sheet[1]/row[2]").GetProps()["data"]);
        Assert.Equal("/sheet[1]/row[2]/cell[B2]", PathResolver.Single(doc.Root, "/sheet[1]/row[2]/cell[B2]").Path);
        Assert.Equal(4, sheet.Children.Count);
        Assert.Equal("90", PathResolver.Single(doc.Root, "//cell[@value=\"90\"]").Text);
        Assert.Equal("Name\tScore\tPass\nAnn\t90\ttrue\nBob\t85\tfalse\nCy\t70,5\ttrue\n", Views.Text(doc.Root));

        var added = Mutations.Add(sheet, "row", Props(("data", "[\"Dee\",60,false]")), null);
        Assert.Equal("/sheet[1]/row[5]", added.Path);
        var cell = Mutations.Add(added, "cell", Props(("value", "extra")), null);
        Assert.Equal("/sheet[1]/row[5]/cell[D5]", cell.Path);
        Assert.Equal("A1:D5", Mutations.Refresh(sheet).GetProps()["range"]);
        PathResolver.Single(doc.Root, "/sheet[1]/range[A5:D5]").Remove();
        Assert.Equal("A1:C4", Mutations.Refresh(sheet).GetProps()["range"]);
        var miss = Assert.Throws<WriterException>(() => Mutations.Add(sheet, "cell", Props(("value", "x")), null));
        Assert.Contains("cell[B3]", miss.Hint);
        var outline = Views.Outline(doc.Root);
        Assert.Contains("/sheet[1]  name=Sheet1  range=A1:C4", outline);
        Assert.DoesNotContain("/sheet[1]/row[1]", outline);
    }

    [Fact]
    public void Styles_survive_and_stack()
    {
        using var doc = new XlsxAdapter().Create();
        var cell = PathResolver.Single(doc.Root, "/sheet[1]/cell[B2]");
        cell = Mutations.Set(cell, Props(("value", "1234.5"), ("bold", "true"), ("color", "C00000"), ("fill", "FFF2CC"), ("format", "#,##0.00")));
        var props = cell.GetProps();
        Assert.Equal(("true", "C00000", "FFF2CC", "#,##0.00"), (props["bold"], props["color"], props["fill"], props["format"]));
        Assert.Equal("number", props["type"]);
        cell = Mutations.Set(cell, Props(("bold", "false"), ("fill", "none")));
        Assert.False(cell.GetProps().ContainsKey("bold"));
        Assert.False(cell.GetProps().ContainsKey("fill"));
        Assert.Equal("C00000", cell.GetProps()["color"]);
        using var reopened = new XlsxAdapter().Open(new MemoryStream(Save(doc)));
        Assert.Equal("#,##0.00", PathResolver.Single(reopened.Root, "/sheet[1]/cell[B2]").GetProps()["format"]);
    }

    [Fact]
    public void Sheets_add_rename_move_remove()
    {
        using var doc = new XlsxAdapter().Create();
        var sales = Mutations.Add(doc.Root, "sheet", Props(("name", "Sales")), null);
        Assert.Equal("/sheet[2]", sales.Path);
        Assert.Equal("Sales", sales.GetProps()["name"]);
        var auto = Mutations.Add(doc.Root, "sheet", Props(), 1);
        Assert.Equal(("Sheet3", "/sheet[1]"), (auto.GetProps()["name"], auto.Path));
        Mutations.Set(auto, Props(("name", "Intro")));
        Assert.Equal(new[] { "Intro", "Sheet1", "Sales" }, doc.Root.Children.Select(s => s.GetProps()["name"]));
        Assert.Throws<WriterException>(() => Mutations.Add(doc.Root, "sheet", Props(("name", "Sales")), null));
        Assert.Throws<WriterException>(() => Mutations.Set(auto, Props(("name", "bad:name"))));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[Sales]/cell[A1]"), Props(("value", "x")));
        Mutations.Move(PathResolver.Single(doc.Root, "/sheet[Sales]"), doc.Root, 1);
        Assert.Equal("x", PathResolver.Single(doc.Root, "/sheet[1]/cell[A1]").Text);
        PathResolver.Single(doc.Root, "/sheet[Intro]").Remove();
        PathResolver.Single(doc.Root, "/sheet[Sheet1]").Remove();
        Assert.Throws<WriterException>(() => PathResolver.Single(doc.Root, "/sheet[1]").Remove());
        using var reopened = new XlsxAdapter().Open(new MemoryStream(Save(doc)));
        Assert.Equal("Sales", reopened.Root.Children.Single().GetProps()["name"]);
    }

    [Fact]
    public void Sheets_inserted_moved_and_removed_keep_what_refers_to_their_position()
    {
        // As Excel writes a workbook: sheets A, B, C; C is the active and selected tab (bookViews activeTab 2); B has an autofilter,
        // a defined name whose localSheetId is B's position.
        var ms = new MemoryStream();
        using (var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(ms, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var wb = package.AddWorkbookPart();
            wb.Workbook = new S.Workbook(new S.BookViews(new S.WorkbookView { ActiveTab = 2U }), new S.Sheets());
            foreach (var (name, i) in new[] { "A", "B", "C" }.Select((n, i) => (n, i)))
            {
                var part = wb.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
                part.Worksheet = new S.Worksheet(new S.SheetViews(new S.SheetView { WorkbookViewId = 0U, TabSelected = name == "C" ? true : null }), new S.SheetData());
                wb.Workbook.Sheets!.Append(new S.Sheet { Id = wb.GetIdOfPart(part), SheetId = (uint)i + 1, Name = name });
            }
            wb.Workbook.Append(new S.DefinedNames(new S.DefinedName { Name = "_xlnm._FilterDatabase", LocalSheetId = 1U, Hidden = true, Text = "B!$A$1:$B$3" }));
        }
        using var doc = new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));
        Mutations.Add(doc.Root, "sheet", Props(("name", "New")), 2);                  // a new sheet placed after A, as a save places a copy
        Mutations.Move(PathResolver.Single(doc.Root, "/sheet[@id=3]"), doc.Root, 1);   // C to the front, addressed by its stable id
        PathResolver.Single(doc.Root, "/sheet[A]").Remove();

        using var saved = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(new MemoryStream(Save(doc)), false);
        var workbook = saved.WorkbookPart!.Workbook!;
        var sheets = workbook.Sheets!.Elements<S.Sheet>().ToList();
        Assert.Equal(new[] { "C", "New", "B" }, sheets.Select(s => s.Name!.Value));
        var filter = workbook.DefinedNames!.Elements<S.DefinedName>().Single();
        Assert.Equal(("B", "B!$A$1:$B$3"), (sheets[(int)filter.LocalSheetId!.Value].Name!.Value, filter.Text));
        Assert.Equal("C", sheets[(int)workbook.BookViews!.GetFirstChild<S.WorkbookView>()!.ActiveTab!.Value].Name!.Value);
        var selected = sheets.Where(s => ((DocumentFormat.OpenXml.Packaging.WorksheetPart)saved.WorkbookPart.GetPartById(s.Id!.Value!)).Worksheet!
            .SheetViews?.GetFirstChild<S.SheetView>()?.TabSelected?.Value == true).Select(s => s.Name!.Value);
        Assert.Equal(new[] { "C" }, selected); // the active tab is the one selected tab: Excel would group two
    }

    [Fact]
    public void Sample_workbooks_read_values_and_formats()
    {
        using var file = File.OpenRead(Path.Combine(TestDocs.FixtureDir("xlsx"), "cell-formatting.xlsx"));
        using var doc = new XlsxAdapter().Open(file);
        var sheet = doc.Root.Children[0];
        Assert.Contains(":", sheet.GetProps()["range"]);
        var cells = PathResolver.Query(doc.Root, "//cell");
        Assert.NotEmpty(cells);
        Assert.Contains(cells, c => c.GetProps().ContainsKey("bold") || c.GetProps().ContainsKey("fill") || c.GetProps().ContainsKey("format"));
        Assert.Contains("\t", Views.Text(doc.Root));
    }

    [Fact]
    public void Sheet_layout_round_trips()
    {
        using var doc = new XlsxAdapter().Create();
        var sheet = doc.Root.Children[0];
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/range[A1:D3]"), Props(("values", "[[\"Title\"],[\"Name\",\"Q1\",\"Q2\",\"Q3\"],[\"Ann\",1,2,3]]")));
        sheet = Mutations.Set(sheet, Props(("merges", "[\"A1:D1\",\"a3:b3\"]"), ("widths", "{\"A\":12.5,\"C\":30}"), ("heights", "{\"3\":24}"), ("freeze", "A2"), ("filter", "A2:D3")));
        var expected = ("[\"A1:D1\",\"A3:B3\"]", "{\"A\":12.5,\"C\":30}", "{\"3\":24}", "A2", "A2:D3");
        static (string, string, string, string, string) Layout(Node s) => (s.GetProps()["merges"], s.GetProps()["widths"], s.GetProps()["heights"], s.GetProps()["freeze"], s.GetProps()["filter"]);
        Assert.Equal(expected, Layout(sheet));
        var saved = Save(doc);
        using (var reopened = new XlsxAdapter().Open(new MemoryStream(saved))) Assert.Equal(expected, Layout(reopened.Root.Children[0]));
        using (var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(new MemoryStream(saved), false))
        {
            var ws = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Assert.Equal(new[] { "sheetViews", "cols", "sheetData", "autoFilter", "mergeCells" }, ws.ChildElements.Select(e => e.LocalName));
            var pane = ws.SheetViews!.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetView>()!.Pane!;
            Assert.Equal((1.0, "A2", "frozen", "bottomLeft"), (pane.VerticalSplit!.Value, pane.TopLeftCell!.Value, pane.State!.InnerText, pane.ActivePane!.InnerText));
            Assert.Null(pane.HorizontalSplit);
            var name = package.WorkbookPart.Workbook!.DefinedNames!.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Single();
            Assert.Equal(("_xlnm._FilterDatabase", 0u, "Sheet1!$A$2:$D$3", true), (name.Name!.Value, name.LocalSheetId!.Value, name.Text, name.Hidden!.Value));
        }
        Assert.Contains("/sheet[1]  name=Sheet1  range=A1:D3  freeze=A2  filter=A2:D3", Views.Outline(doc.Root));

        sheet = Mutations.Set(sheet, Props(("widths", "{\"B\":8,\"C\":null}"), ("heights", "{\"3\":\"default\"}"), ("freeze", "none"), ("filter", "none"), ("merges", "[]")));
        Assert.Equal("{\"A\":12.5,\"B\":8}", sheet.GetProps()["widths"]);
        Assert.DoesNotContain(sheet.GetProps().Keys, k => k is "heights" or "freeze" or "filter" or "merges");
        sheet = Mutations.Set(sheet, Props(("freeze", "B2")));
        Assert.Equal("B2", sheet.GetProps()["freeze"]);

        Check(() => Mutations.Set(sheet, Props(("merges", "[\"A1:C1\",\"B1:B2\"]"))), "overlaps");
        Check(() => Mutations.Set(sheet, Props(("merges", "[\"A1:\"]"))), "not a cell range");
        Check(() => Mutations.Set(sheet, Props(("merges", "[\"A1\"]"))), "single cell");
        Check(() => Mutations.Set(sheet, Props(("widths", "{\"A\":-1}"))), "above 0");
        Check(() => Mutations.Set(sheet, Props(("widths", "{\"1\":10}"))), "column letter");
        Check(() => Mutations.Set(sheet, Props(("heights", "{\"x\":10}"))), "row number");
        Check(() => Mutations.Set(sheet, Props(("freeze", "nope"))), "cell reference");
        Check(() => Mutations.Set(sheet, Props(("filter", "1:2"))), "cell range");
        Assert.Equal("B2", Mutations.Refresh(sheet).GetProps()["freeze"]);

        var second = Mutations.Add(doc.Root, "sheet", Props(("name", "Two")), null);
        Mutations.Set(second, Props(("filter", "A1:B2")));
        PathResolver.Single(doc.Root, "/sheet[1]").Remove();
        using var package2 = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(new MemoryStream(Save(doc)), false);
        var moved = package2.WorkbookPart!.Workbook!.DefinedNames!.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Single();
        Assert.Equal((0u, "Two!$A$1:$B$2"), (moved.LocalSheetId!.Value, moved.Text));

        static void Check(Action act, string message)
        {
            var ex = Assert.Throws<WriterException>(act);
            Assert.Equal(ErrorCode.Validation, ex.Code);
            Assert.Contains(message, ex.Message);
        }
    }

    [Fact]
    public void Sample_workbook_layout_reads_and_an_edit_touches_only_its_sheet()
    {
        var original = File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("xlsx"), "cell-formatting.xlsx"));
        using var doc = new XlsxAdapter().Open(new MemoryStream(original));
        var data = PathResolver.Single(doc.Root, "/sheet[Data]");
        Assert.Equal(("[\"A13:C13\"]", "{\"A\":40,\"B\":16}"), (data.GetProps()["merges"], data.GetProps()["widths"]));
        Assert.Equal("{\"6\":34,\"7\":34,\"8\":34,\"14\":40}", PathResolver.Single(doc.Root, "/sheet[Fills]").GetProps()["heights"]);
        Mutations.Set(data, Props(("freeze", "B2"), ("widths", "{\"B\":18}"), ("merges", "[\"A13:C13\",\"A1:B1\"]")));
        var saved = Save(doc);
        Assert.Null(PackageCompare.Diff(original, saved, ((XlsxDocument)doc).Sheets[4].Part.Uri.ToString().TrimStart('/')));
        using var reopened = new XlsxAdapter().Open(new MemoryStream(saved));
        var again = PathResolver.Single(reopened.Root, "/sheet[Data]").GetProps();
        Assert.Equal(("B2", "{\"A\":40,\"B\":18}", "[\"A13:C13\",\"A1:B1\"]"), (again["freeze"], again["widths"], again["merges"]));

        using var settings = new XlsxAdapter().Open(File.OpenRead(Path.Combine(TestDocs.FixtureDir("xlsx"), "sheet-settings.xlsx")));
        Assert.Equal("B2", settings.Root.Children[0].GetProps()["freeze"]);
        Assert.Equal("A1:B5", settings.Root.Children[3].GetProps()["filter"]);
    }

    [Fact]
    public void Shared_formulas_are_read_shifted_and_survive_edits()
    {
        var file = Path.Combine(TestDocs.FixtureDir("xlsx"), "budget-tracker.xlsx");
        using var doc = new XlsxAdapter().Open(new MemoryStream(File.ReadAllBytes(file)));
        static string Prop(Document d, string cell, string prop) => PathResolver.Single(d.Root, $"/sheet[1]/cell[{cell}]").GetProps().GetValueOrDefault(prop) ?? "";
        Assert.Equal("SUM(C9:F9)", Prop(doc, "G9", "formula")); // <f t="shared" si="0"/>: the master's SUM(C8:F8) moved down one row
        Assert.Equal("G14/B14", Prop(doc, "H14", "formula"));

        // the same formula written back (as the app does for a retyped cell) leaves the group shared and the cached results in place
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[G8]"), Props(("formula", "SUM(C8:F8)")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[G9]"), Props(("formula", "=SUM(C9:F9)")));
        var data = ((XlsxDocument)doc).Sheets[0].Part.Worksheet!.GetFirstChild<S.SheetData>()!;
        Assert.Equal(7, data.Descendants<S.CellFormula>().Count(f => f.FormulaType?.InnerText == "shared" && f.SharedIndex?.Value == 0));
        Assert.Equal("375000", Prop(doc, "G9", "value"));

        // a real change to one member: every other member keeps its own formula and value, the other group is untouched
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[G10]"), Props(("formula", "SUM(C10:E10)")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[H9]"), Props(("value", "1")));
        using var reopened = new XlsxAdapter().Open(new MemoryStream(Save(doc)));
        Assert.Equal("SUM(C10:E10)", Prop(reopened, "G10", "formula"));
        Assert.Equal(("SUM(C9:F9)", "375000"), (Prop(reopened, "G9", "formula"), Prop(reopened, "G9", "value")));
        Assert.Equal(("SUM(C14:F14)", "95000"), (Prop(reopened, "G14", "formula"), Prop(reopened, "G14", "value")));
        Assert.Equal(("", "1"), (Prop(reopened, "H9", "formula"), Prop(reopened, "H9", "value")));
        Assert.Equal(("G10/B10", "0.95"), (Prop(reopened, "H10", "formula"), Prop(reopened, "H10", "value")));
        Assert.Equal("SUM(B8:B14)", Prop(reopened, "B15", "formula"));
    }

    [Fact]
    public void Strings_keep_their_storage_and_the_shared_table_is_reused()
    {
        var file = Path.Combine(TestDocs.FixtureDir("xlsx"), "cell-formatting.xlsx");
        using var doc = new XlsxAdapter().Open(new MemoryStream(File.ReadAllBytes(file)));
        var rich = PathResolver.Single(doc.Root, "/sheet[6]/cell[A3]"); // a shared string of three formatted runs
        var text = rich.GetProps()["value"];
        Assert.Equal("Bold + Red  Italic + Blue  Normal", text);
        Mutations.Set(rich, Props(("value", text), ("type", "string")));
        var sst = ((XlsxDocument)doc).Workbook.SharedStringTablePart!.SharedStringTable!;
        Assert.Equal(3, sst.Elements<S.SharedStringItem>().First().Elements<S.Run>().Count()); // its runs are still there
        Assert.Equal(3, sst.Elements<S.SharedStringItem>().Count());

        var plain = PathResolver.Single(doc.Root, "/sheet[1]/cell[A1]"); // stored as t="str" by the file's writer
        Mutations.Set(plain, Props(("value", "Retyped")));
        Assert.Equal("str", ((XlsxDocument)doc).Sheets[0].Part.Worksheet!.Descendants<S.Cell>().First(c => c.CellReference?.Value == "A1").DataType?.InnerText);
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[6]/cell[B1]"), Props(("value", "Twice")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[6]/cell[B2]"), Props(("value", "Twice")));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[6]/cell[B3]"), Props(("value", "Bold + Red  Italic + Blue  Normal")));
        Assert.Equal(5, sst.Elements<S.SharedStringItem>().Count()); // one entry for the two "Twice" cells; the rich entry is never handed out, B3 gets a plain one
        using var reopened = new XlsxAdapter().Open(new MemoryStream(Save(doc)));
        Assert.Equal("Twice", PathResolver.Single(reopened.Root, "/sheet[6]/cell[B2]").GetProps()["value"]);
        Assert.Equal("Retyped", PathResolver.Single(reopened.Root, "/sheet[1]/cell[A1]").GetProps()["value"]);
    }
}

public class XlsxFidelityTests
{
    public static TheoryData<string> Files => new(TestDocs.Fixtures("xlsx", "*.xlsx").Select(f => Path.GetFileName(f)));

    [Theory]
    [MemberData(nameof(Files))]
    public void Open_read_everything_save_changes_nothing(string file)
    {
        var original = File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("xlsx"), file));
        using var doc = new XlsxAdapter().Open(new MemoryStream(original));
        _ = Views.Outline(doc.Root);
        _ = Views.Text(doc.Root);
        _ = NodeJson.Serialize(doc.Root, int.MaxValue);
        using var saved = new MemoryStream();
        doc.Save(saved);
        Assert.Null(PackageCompare.Diff(original, saved.ToArray()));
    }

    [Theory]
    [MemberData(nameof(Files))]
    public void Writing_one_cell_touches_only_its_sheet_and_shared_strings(string file)
    {
        var original = File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("xlsx"), file));
        using var doc = new XlsxAdapter().Open(new MemoryStream(original));
        var sheet = (WorksheetPartOf(doc));
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[ZZ999]"), new Dictionary<string, string> { ["value"] = "EDITED" });
        using var saved = new MemoryStream();
        doc.Save(saved);
        Assert.Null(PackageCompare.Diff(original, saved.ToArray(), sheet, "xl/sharedStrings.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/workbook.xml"));
    }

    static string WorksheetPartOf(Document doc) =>
        ((XlsxDocument)doc).Sheets[0].Part.Uri.ToString().TrimStart('/');
}
