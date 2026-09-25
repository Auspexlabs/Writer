using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
using Writer.Formats.Xlsx;

namespace Writer.Tests;

/// <summary>Sheet rules: conditional formatting, data validation, filter criteria, hidden rows and columns, tab colour and text rotation
/// read from Excel's own files and round-trip through the engine.</summary>
public class XlsxRulesTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static Document Fixture(string name) => new XlsxAdapter().Open(new MemoryStream(File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("xlsx"), name))));

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    static Document Reopen(Document doc) => new XlsxAdapter().Open(new MemoryStream(Save(doc)));

    static IReadOnlyDictionary<string, string> Get(Document doc, string path) => PathResolver.Single(doc.Root, path).GetProps();

    static Node Set(Document doc, string path, params (string Name, string Value)[] pairs) => Mutations.Set(PathResolver.Single(doc.Root, path), Props(pairs));

    static JsonElement Json(string s) => JsonDocument.Parse(s).RootElement;

    [Fact]
    public void Conditional_formatting_reads_excels_rules_and_round_trips_the_editors()
    {
        using var fixture = Fixture("conditional-formatting.xlsx");
        var cells = Json(Get(fixture, "/sheet[1]")["cf"]);
        Assert.Equal(4, cells.GetArrayLength());
        Assert.Equal(("A2:A11", "cellIs", "greaterThan", "80", "C6EFCE"), (cells[0].GetProperty("range").GetString(), cells[0].GetProperty("type").GetString(), cells[0].GetProperty("operator").GetString(), cells[0].GetProperty("value").GetString(), cells[0].GetProperty("fill").GetString()));
        Assert.Equal(("between", "50", "70"), (cells[2].GetProperty("operator").GetString(), cells[2].GetProperty("value").GetString(), cells[2].GetProperty("value2").GetString()));
        var text = Json(Get(fixture, "/sheet[2]")["cf"]);
        Assert.Equal(("containsText", "error"), (text[0].GetProperty("type").GetString(), text[0].GetProperty("text").GetString()));
        var top = Json(Get(fixture, "/sheet[3]")["cf"]);
        Assert.Equal((3, true), (top[1].GetProperty("rank").GetInt32(), top[1].GetProperty("bottom").GetBoolean()));
        Assert.True(top[2].GetProperty("percent").GetBoolean());
        Assert.False(top[4].GetProperty("above").GetBoolean());
        Assert.Equal("638EC6", Json(Get(fixture, "/sheet[4]")["cf"])[0].GetProperty("color").GetString());
        var scale = Json(Get(fixture, "/sheet[5]")["cf"]);
        Assert.Equal(new[] { "F8696B", "FFEB84", "63BE7B" }, scale[1].GetProperty("colors").EnumerateArray().Select(c => c.GetString()));
        Assert.Equal(("3Arrows", "duplicateValues", "thisMonth"), (Json(Get(fixture, "/sheet[6]")["cf"])[1].GetProperty("iconSet").GetString(), Json(Get(fixture, "/sheet[7]")["cf"])[1].GetProperty("type").GetString(), Json(Get(fixture, "/sheet[7]")["cf"])[3].GetProperty("period").GetString()));

        using var doc = new XlsxAdapter().Create();
        Set(doc, "/sheet[1]/range[A1:A5]", ("values", "[[5],[15],[15],[40],[\"note error\"]]"));
        const string rules = "[{\"range\":\"A1:A5\",\"type\":\"cellIs\",\"operator\":\"greaterThan\",\"value\":\"10\",\"fill\":\"FFC7CE\",\"color\":\"9C0006\"},"
            + "{\"range\":\"A1:A5\",\"type\":\"between\",\"operator\":\"between\",\"value\":\"1\",\"value2\":\"9\",\"fill\":\"FFEB9C\"},"
            + "{\"range\":\"A1:A5\",\"type\":\"containsText\",\"text\":\"error\",\"fill\":\"C6EFCE\",\"color\":\"006100\",\"bold\":true},"
            + "{\"range\":\"A1:A5\",\"type\":\"duplicateValues\",\"fill\":\"FFC7CE\"},{\"range\":\"A1:A5\",\"type\":\"top10\",\"rank\":3,\"percent\":true,\"fill\":\"C6EFCE\"},"
            + "{\"range\":\"A1:A5\",\"type\":\"dataBar\",\"color\":\"638EC6\"},{\"range\":\"A1:A5\",\"type\":\"colorScale\",\"colors\":[\"F8696B\",\"FFEB84\",\"63BE7B\"]}]";
        Assert.Contains("not a rule type", Assert.Throws<WriterException>(() => Set(doc, "/sheet[1]", ("cf", rules))).Message);
        var fixed_ = rules.Replace("\"type\":\"between\"", "\"type\":\"cellIs\"");
        Set(doc, "/sheet[1]", ("cf", fixed_));
        using var reopened = Reopen(doc);
        var back = Json(Get(reopened, "/sheet[1]")["cf"]);
        Assert.Equal(7, back.GetArrayLength());
        Assert.Equal(("greaterThan", "10", "FFC7CE", "9C0006"), (back[0].GetProperty("operator").GetString(), back[0].GetProperty("value").GetString(), back[0].GetProperty("fill").GetString(), back[0].GetProperty("color").GetString()));
        Assert.Equal(("between", "1", "9"), (back[1].GetProperty("operator").GetString(), back[1].GetProperty("value").GetString(), back[1].GetProperty("value2").GetString()));
        Assert.Equal(("error", true), (back[2].GetProperty("text").GetString(), back[2].GetProperty("bold").GetBoolean()));
        Assert.Equal((3, true), (back[4].GetProperty("rank").GetInt32(), back[4].GetProperty("percent").GetBoolean()));
        Assert.Equal("638EC6", back[5].GetProperty("color").GetString());
        Assert.Equal(3, back[6].GetProperty("colors").GetArrayLength());
        using (var package = SpreadsheetDocument.Open(new MemoryStream(Save(reopened)), false))
        {
            var ws = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Assert.Equal(7, ws.Elements<ConditionalFormatting>().Count());
            Assert.Equal("NOT(ISERROR(SEARCH(\"error\",A1)))", ws.Elements<ConditionalFormatting>().ElementAt(2).GetFirstChild<ConditionalFormattingRule>()!.GetFirstChild<Formula>()!.Text);
            Assert.Equal(new[] { 1, 2, 3, 4, 5, 6, 7 }, ws.Elements<ConditionalFormatting>().Select(c => c.GetFirstChild<ConditionalFormattingRule>()!.Priority!.Value));
            Assert.Equal(5u, package.WorkbookPart.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Count!.Value); // seven rules, five distinct looks: identical ones share a dxf
        }
        Set(reopened, "/sheet[1]", ("cf", "[]"));
        Assert.False(Get(reopened, "/sheet[1]").ContainsKey("cf"));
    }

    [Fact]
    public void Data_validation_reads_excels_rules_and_round_trips_the_editors()
    {
        using var fixture = Fixture("data-validation.xlsx");
        var lists = Json(Get(fixture, "/sheet[1]")["validations"]);
        Assert.Equal(("A2:A20", "list"), (lists[0].GetProperty("range").GetString(), lists[0].GetProperty("type").GetString()));
        Assert.Equal(new[] { "Draft", "Review", "Approved", "Rejected" }, lists[0].GetProperty("values").EnumerateArray().Select(v => v.GetString()));
        Assert.Equal("H2:H5", lists[1].GetProperty("source").GetString());
        var numbers = Json(Get(fixture, "/sheet[2]")["validations"]);
        Assert.Equal(("whole", "between", "1", "100"), (numbers[0].GetProperty("type").GetString(), numbers[0].GetProperty("operator").GetString(), numbers[0].GetProperty("value").GetString(), numbers[0].GetProperty("value2").GetString()));
        Assert.Equal(("decimal", "lessThanOrEqual", "0.5"), (numbers[1].GetProperty("type").GetString(), numbers[1].GetProperty("operator").GetString(), numbers[1].GetProperty("value").GetString()));
        Assert.Equal(("date", "45292"), (Json(Get(fixture, "/sheet[3]")["validations"])[0].GetProperty("type").GetString(), Json(Get(fixture, "/sheet[3]")["validations"])[0].GetProperty("value").GetString()));

        using var doc = new XlsxAdapter().Create();
        Set(doc, "/sheet[1]", ("validations", "[{\"range\":\"A2:A20\",\"type\":\"list\",\"values\":[\"Yes\",\"No\"],\"error\":\"Pick Yes or No\",\"errorTitle\":\"Answer\"},"
            + "{\"range\":\"B2:B20\",\"type\":\"whole\",\"operator\":\"between\",\"value\":\"1\",\"value2\":\"100\"},{\"range\":\"C2:C20\",\"type\":\"list\",\"source\":\"Sheet1!H2:H5\"},"
            + "{\"range\":\"D2:D9\",\"type\":\"decimal\",\"operator\":\"greaterThan\",\"value\":\"0\",\"allowBlank\":false},{\"range\":\"E1\",\"type\":\"any\"}]"));
        using var reopened = Reopen(doc);
        var back = Json(Get(reopened, "/sheet[1]")["validations"]);
        Assert.Equal(4, back.GetArrayLength());
        Assert.Equal(new[] { "Yes", "No" }, back[0].GetProperty("values").EnumerateArray().Select(v => v.GetString()));
        Assert.Equal(("Pick Yes or No", "Answer"), (back[0].GetProperty("error").GetString(), back[0].GetProperty("errorTitle").GetString()));
        Assert.Equal(("whole", "between", "1", "100"), (back[1].GetProperty("type").GetString(), back[1].GetProperty("operator").GetString(), back[1].GetProperty("value").GetString(), back[1].GetProperty("value2").GetString()));
        Assert.Equal("Sheet1!H2:H5", back[2].GetProperty("source").GetString());
        Assert.False(back[3].GetProperty("allowBlank").GetBoolean());
        using (var package = SpreadsheetDocument.Open(new MemoryStream(Save(reopened)), false))
        {
            var dvs = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<DataValidations>()!;
            Assert.Equal(4u, dvs.Count!.Value);
            Assert.Equal(("\"Yes,No\"", "Sheet1!$H$2:$H$5"), (dvs.Elements<DataValidation>().First().Formula1!.Text, dvs.Elements<DataValidation>().ElementAt(2).Formula1!.Text));
            Assert.True(dvs.Elements<DataValidation>().First().ShowErrorMessage!.Value);
        }
        Assert.Contains("not a validation type", Assert.Throws<WriterException>(() => Set(reopened, "/sheet[1]", ("validations", "[{\"range\":\"A1\",\"type\":\"nope\"}]"))).Message);
    }

    [Fact]
    public void Filter_criteria_hidden_lines_tab_colour_and_rotation_round_trip()
    {
        using var doc = new XlsxAdapter().Create();
        Set(doc, "/sheet[1]/range[A1:C4]", ("values", "[[\"Region\",\"Item\",\"Sales\"],[\"East\",\"Pen\",120],[\"West\",\"Pencil\",80],[\"East\",\"\",200]]"));
        Assert.Contains("autofilter range", Assert.Throws<WriterException>(() => Set(doc, "/sheet[1]", ("filters", "{\"A\":{\"values\":[\"East\"]}}"))).Message);
        Set(doc, "/sheet[1]", ("filter", "A1:C4"), ("filters", "{\"A\":{\"values\":[\"East\",\"\"]},\"B\":{\"operator\":\"contains\",\"value\":\"Pen\"},\"C\":{\"operator\":\"between\",\"value\":\"100\",\"value2\":\"300\"}}"),
            ("hidden", "{\"rows\":[3],\"cols\":[\"B\",\"D\"]}"), ("color", "C0392B"));
        Set(doc, "/sheet[1]/cell[A1]", ("rotate", "45"));
        Set(doc, "/sheet[1]/cell[B1]", ("rotate", "-30"));
        Assert.Contains("outside the autofilter range", Assert.Throws<WriterException>(() => Set(doc, "/sheet[1]", ("filters", "{\"F\":{\"values\":[\"x\"]}}"))).Message);
        using var reopened = Reopen(doc);
        var sheet = Get(reopened, "/sheet[1]");
        var filters = Json(sheet["filters"]);
        Assert.Equal(new[] { "East", "" }, filters.GetProperty("A").GetProperty("values").EnumerateArray().Select(v => v.GetString()));
        Assert.Equal(("contains", "Pen"), (filters.GetProperty("B").GetProperty("operator").GetString(), filters.GetProperty("B").GetProperty("value").GetString()));
        Assert.Equal(("between", "100", "300"), (filters.GetProperty("C").GetProperty("operator").GetString(), filters.GetProperty("C").GetProperty("value").GetString(), filters.GetProperty("C").GetProperty("value2").GetString()));
        Assert.Equal("{\"rows\":[3],\"cols\":[\"B\",\"D\"]}", sheet["hidden"]);
        Assert.Equal("C0392B", sheet["color"]);
        Assert.Equal(("45", "-30"), (Get(reopened, "/sheet[1]/cell[A1]")["rotate"], Get(reopened, "/sheet[1]/cell[B1]")["rotate"]));
        using (var package = SpreadsheetDocument.Open(new MemoryStream(Save(reopened)), false))
        {
            var ws = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Assert.Equal("sheetPr", ws.ChildElements.First().LocalName);
            var af = ws.GetFirstChild<AutoFilter>()!;
            Assert.Equal(new uint[] { 0, 1, 2 }, af.Elements<FilterColumn>().Select(c => c.ColumnId!.Value));
            Assert.Equal("*Pen*", af.Elements<FilterColumn>().ElementAt(1).GetFirstChild<CustomFilters>()!.GetFirstChild<CustomFilter>()!.Val!.Value);
            Assert.True(af.Elements<FilterColumn>().First().GetFirstChild<Filters>()!.Blank!.Value);
            uint Rotation(string reference) => package.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>()
                .ElementAt((int)ws.Descendants<Cell>().First(c => c.CellReference == reference).StyleIndex!.Value).Alignment!.TextRotation!.Value;
            Assert.Equal((45u, 120u), (Rotation("A1"), Rotation("B1"))); // Excel: 90 + degrees for clockwise
        }
        Set(reopened, "/sheet[1]", ("filters", "{}"), ("hidden", "{\"rows\":[],\"cols\":[\"D\"]}"), ("color", "none"));
        var cleared = Get(reopened, "/sheet[1]");
        Assert.False(cleared.ContainsKey("filters"));
        Assert.False(cleared.ContainsKey("color"));
        Assert.Equal("{\"rows\":[],\"cols\":[\"D\"]}", cleared["hidden"]);
        Assert.Equal("A1:C4", cleared["filter"]);

        using var settings = Fixture("sheet-settings.xlsx");
        Assert.Equal("C0392B", settings.Root.Children[3].GetProps()["color"]);
    }
}
