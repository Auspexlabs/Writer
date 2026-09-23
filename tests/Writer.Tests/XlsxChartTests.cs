using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Cli;
using Writer.Core;
using Writer.Formats.Xlsx;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Writer.Tests;

public class XlsxChartTests
{
    const string Series = "[{\"name\":\"B1\",\"values\":\"B2:B6\"},{\"name\":\"C1\",\"values\":\"C2:C6\"}]";
    const long Col = 609600; // a default column
    const long Row = 190500; // a default row

    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    /// <summary>A workbook with Month, Sales and Cost columns in A1:C6.</summary>
    static Document Data()
    {
        var doc = new XlsxAdapter().Create();
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/range[A1:C6]"), Props(("values",
            "[[\"Month\",\"Sales\",\"Cost\"],[\"Jan\",120,80],[\"Feb\",135,90],[\"Mar\",148,95],[\"Apr\",162,100],[\"May\",155,98]]")));
        return doc;
    }

    static string PartNamed(byte[] package, string contains) => PackageCompare.Parts(package).Keys.Single(k => k.Contains(contains));

    static string Emu(string length) => Units.ParseLength(length).ToString();

    [Theory]
    [InlineData("column", "barChart")]
    [InlineData("bar", "barChart")]
    [InlineData("line", "lineChart")]
    [InlineData("pie", "pieChart")]
    [InlineData("area", "areaChart")]
    [InlineData("scatter", "scatterChart")]
    [InlineData("doughnut", "doughnutChart")]
    public void Adds_each_chart_type_the_way_excel_writes_it(string type, string element)
    {
        using var doc = Data();
        var chart = Mutations.Add(doc.Root.Children[0], "chart",
            Props(("type", type), ("title", "Sales by month"), ("categories", "A2:A6"), ("series", Series), ("legend", "right"), ("stacked", "true")), null);
        Assert.Equal("/sheet[1]/chart[1]", chart.Path);
        var bytes = Save(doc);

        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package).Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}"));
        var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!;
        var anchor = drawing.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().Single();
        var from = anchor.FromMarker!;
        Assert.Equal(("4", "0", "0", "0"), (from.ColumnId!.Text, from.ColumnOffset!.Text, from.RowId!.Text, from.RowOffset!.Text));
        Assert.Equal("Chart 2", anchor.GetFirstChild<Xdr.GraphicFrame>()!.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!.Name!.Value);
        var c = drawing.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!;
        var plot = c.PlotArea!;
        var el = plot.ChildElements.Single(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        Assert.Equal(element, el.LocalName);
        var sers = el.ChildElements.Where(e => e.LocalName == "ser").ToList();
        Assert.Equal(2, sers.Count);
        var formulas = sers[0].Descendants<C.Formula>().Select(f => f.Text).ToList();
        Assert.Contains("Sheet1!$B$1", formulas);
        Assert.Contains("Sheet1!$B$2:$B$6", formulas);
        Assert.Contains("Sheet1!$A$2:$A$6", formulas);
        var cache = sers[0].Descendants<C.NumberingCache>().Last();
        Assert.Equal(5, cache.Elements<C.NumericPoint>().Count());
        Assert.Equal("120", cache.Elements<C.NumericPoint>().First().NumericValue!.Text);
        Assert.Equal("Sales by month", string.Concat(c.Title!.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text)));
        Assert.False(c.AutoTitleDeleted!.Val!.Value);
        Assert.Equal("r", c.Legend!.LegendPosition!.Val!.InnerText);
        Assert.True(c.PlotVisibleOnly!.Val!.Value);
        Assert.Equal(type is "pie" or "doughnut" ? 0 : 2, plot.ChildElements.Count(e => e.LocalName is "catAx" or "valAx"));
        if (type == "bar") Assert.Equal("bar", el.GetFirstChild<C.BarDirection>()!.Val!.InnerText);
        if (type is "column" or "bar" or "line" or "area") Assert.Equal("stacked", el.ChildElements.First(e => e.LocalName == "grouping").GetAttributes().First().Value);

        using var reopened = new XlsxAdapter().Open(new MemoryStream(bytes));
        var props = PathResolver.Single(reopened.Root, "/sheet[1]/chart[1]").GetProps();
        Assert.Equal((type, "Sales by month", "A2:A6", Series, "right"), (props["type"], props["title"], props["categories"], props["series"], props["legend"]));
        Assert.Equal(type is "column" or "bar" or "line" or "area" ? "true" : null, props.GetValueOrDefault("stacked"));
        Assert.Equal(((4 * Col).ToString(), "0", Emu("15cm"), Emu("7.5cm"), "2"), (props["x"], props["y"], props["w"], props["h"], props["id"]));
        Assert.Contains($"/sheet[1]/chart[1]  type={type}  title=Sales by month", Views.Outline(reopened.Root));
        Assert.Equal("/sheet[1]/chart[1]", PathResolver.Single(reopened.Root, "/sheet[1]/chart[Sales by month]").Path);
    }

    [Fact]
    public void Position_round_trips_over_custom_widths_and_heights()
    {
        using var doc = Data();
        var sheet = doc.Root.Children[0];
        Mutations.Set(sheet, Props(("widths", "{\"B\":20,\"D\":5}"), ("heights", "{\"2\":30}")));
        var box = Props(("x", "3.3cm"), ("y", "1.7cm"), ("w", "12.34cm"), ("h", "6.78cm"));
        var chart = Mutations.Add(sheet, "chart", new Dictionary<string, string>(box) { ["series"] = "[{\"values\":\"B2:B6\"}]" }, null);
        var expected = (Emu("3.3cm"), Emu("1.7cm"), Emu("12.34cm"), Emu("6.78cm"));
        var p = chart.GetProps();
        Assert.Equal(expected, (p["x"], p["y"], p["w"], p["h"]));

        var bytes = Save(doc);
        using var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var anchor = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>().Single();
        Assert.Equal(("1", "578400", "2", "40500"), (anchor.FromMarker!.ColumnId!.Text, anchor.FromMarker.ColumnOffset!.Text, anchor.FromMarker.RowId!.Text, anchor.FromMarker.RowOffset!.Text));
        Assert.Equal("8", anchor.ToMarker!.ColumnId!.Text);
        using var reopened = new XlsxAdapter().Open(new MemoryStream(bytes));
        p = PathResolver.Single(reopened.Root, "//chart[1]").GetProps();
        Assert.Equal(expected, (p["x"], p["y"], p["w"], p["h"]));

        chart = Mutations.Set(chart, Props(("x", "0cm")));
        Assert.Equal(("0", Emu("12.34cm")), (chart.GetProps()["x"], chart.GetProps()["w"]));
    }

    [Fact]
    public void Set_title_type_and_position_change_only_their_part()
    {
        using var doc = Data();
        var sheet = doc.Root.Children[0];
        var chart = Mutations.Add(sheet, "chart", Props(("title", "Before"), ("categories", "A2:A6"), ("series", Series)), null);
        var before = Save(doc);
        var chartPart = PartNamed(before, "/charts/");
        var drawingPart = PartNamed(before, "/drawings/drawing");

        chart = Mutations.Set(chart, Props(("title", "After")));
        var afterTitle = Save(doc);
        Assert.Null(PackageCompare.Diff(before, afterTitle, chartPart));
        Assert.Contains("After", Encoding.UTF8.GetString(PackageCompare.Parts(afterTitle)[chartPart]));

        chart = Mutations.Set(chart, Props(("type", "line")));
        var afterType = Save(doc);
        Assert.Null(PackageCompare.Diff(afterTitle, afterType, chartPart));
        Assert.Equal(("line", Series, "A2:A6", "After"), (chart.GetProps()["type"], chart.GetProps()["series"], chart.GetProps()["categories"], chart.GetProps()["title"]));
        Assert.Contains("lineChart", Encoding.UTF8.GetString(PackageCompare.Parts(afterType)[chartPart]));

        chart = Mutations.Set(chart, Props(("x", "1cm"), ("y", "2cm")));
        var afterMove = Save(doc);
        Assert.Null(PackageCompare.Diff(afterType, afterMove, drawingPart));
        Assert.Equal((Emu("1cm"), Emu("2cm"), Emu("15cm")), (chart.GetProps()["x"], chart.GetProps()["y"], chart.GetProps()["w"]));

        chart = Mutations.Set(chart, Props(("legend", "none"), ("stacked", "true"), ("series", "[{\"name\":\"Total\",\"values\":\"C2:C6\"}]"), ("categories", "")));
        Assert.Equal(("none", "true", "[{\"name\":\"Total\",\"values\":\"C2:C6\"}]", null), (chart.GetProps()["legend"], chart.GetProps()["stacked"], chart.GetProps()["series"], chart.GetProps().GetValueOrDefault("categories")));
        chart = Mutations.Set(chart, Props(("title", ""), ("legend", "top")));
        Assert.False(chart.GetProps().ContainsKey("title"));
        Assert.Equal("top", chart.GetProps()["legend"]);
        using var reopened = new XlsxAdapter().Open(new MemoryStream(Save(doc)));
        var c = PathResolver.Single(reopened.Root, "/sheet[1]/chart[1]").GetProps();
        Assert.Equal(("line", "top", "true"), (c["type"], c["legend"], c["stacked"]));
        Assert.False(c.ContainsKey("title"));
    }

    [Fact]
    public void Remove_deletes_the_anchor_and_the_parts()
    {
        using var doc = Data();
        var sheet = doc.Root.Children[0];
        var plain = Save(doc);
        Mutations.Add(sheet, "chart", Props(("series", "[{\"values\":\"B2:B6\"}]")), null);
        Mutations.Add(sheet, "chart", Props(("series", "[{\"values\":\"C2:C6\"}]"), ("title", "Second")), null);
        var second = PathResolver.Single(doc.Root, "/sheet[1]/chart[2]");
        Assert.Equal("Second", second.GetProps()["title"]);
        Assert.Equal((Units.ParseLength("7.5cm") + Row).ToString(), second.GetProps()["y"]);
        Assert.Equal("/sheet[1]/chart[2]", PathResolver.Single(doc.Root, "//chart[@id=3]").Path);
        Assert.Equal(2, PackageCompare.Parts(Save(doc)).Keys.Count(k => k.Contains("/charts/")));

        PathResolver.Single(doc.Root, "/sheet[1]/chart[1]").Remove();
        var one = Save(doc);
        Assert.Single(PackageCompare.Parts(one).Keys, k => k.Contains("/charts/"));
        Assert.Equal("Second", PathResolver.Single(doc.Root, "/sheet[1]/chart[1]").GetProps()["title"]);

        PathResolver.Single(doc.Root, "/sheet[1]/chart[1]").Remove();
        Assert.Empty(PathResolver.Query(doc.Root, "//chart"));
        var none = Save(doc);
        Assert.DoesNotContain(PackageCompare.Parts(none).Keys, k => k.Contains("/charts/") || k.Contains("/drawings/"));
        Assert.Null(PackageCompare.Diff(plain, none, "xl/worksheets/_rels/sheet1.xml.rels"));
    }

    [Fact]
    public void Reads_charts_made_by_other_tools_and_edits_them_gently()
    {
        var original = File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("xlsx"), "charts.xlsx"));
        using var doc = new XlsxAdapter().Open(new MemoryStream(original));
        var charts = PathResolver.Query(doc.Root, "/sheet[1]/chart");
        Assert.Equal(4, charts.Count);
        var first = charts[0].GetProps();
        Assert.Equal(("column", "Monthly Sales and YoY Growth Trend", "bottom", "A2:A13", "2"), (first["type"], first["title"], first["legend"], first["categories"], first["id"]));
        Assert.Equal(5, JsonDocument.Parse(first["series"]).RootElement.GetArrayLength());
        Assert.Contains("{\"name\":\"East Sales\",\"values\":\"B2:B13\"}", first["series"]);
        Assert.Equal(((7 * Col).ToString(), "0", (11 * Col).ToString(), (18 * Row).ToString()), (first["x"], first["y"], first["w"], first["h"]));
        Assert.Equal(new[] { "column", "pie", "doughnut" }, charts.Skip(1).Select(c => c.GetProps()["type"]));
        Assert.Equal("/sheet[1]/chart[3]", PathResolver.Single(doc.Root, "//chart[@id=4]").Path);
        Assert.Contains("/sheet[1]/chart[1]  type=column  title=Monthly Sales and YoY Growth Trend", Views.Outline(doc.Root));
        var types = PathResolver.Query(doc.Root, "//chart").Select(c => c.GetProps()["type"]).ToList();
        Assert.Equal(8, types.Count);
        Assert.Contains("scatter", types);
        Assert.Contains("radar", types);
        Assert.Equal("A2:A16", PathResolver.Single(doc.Root, "//chart[@type=\"scatter\"]").GetProps()["categories"]);

        var edited = Mutations.Set(charts[0], Props(("title", "Renamed"), ("x", "0cm")));
        Assert.Equal(5, JsonDocument.Parse(edited.GetProps()["series"]).RootElement.GetArrayLength());
        var chartPart = "xl/drawings/charts/chart1.xml";
        Assert.Null(PackageCompare.Diff(original, Save(doc), chartPart, "xl/drawings/drawing1.xml"));
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(PathResolver.Single(doc.Root, "//chart[@type=\"radar\"]"), Props(("series", "[{\"values\":\"B2:B6\"}]"))));
        Assert.Equal(ErrorCode.Validation, ex.Code);
        Assert.Contains("type=column", ex.Hint);
    }

    [Fact]
    public void Renaming_a_sheet_keeps_its_charts_and_filter_pointing_at_it()
    {
        using var doc = Data();
        var sheet = doc.Root.Children[0];
        Mutations.Add(sheet, "chart", Props(("categories", "A2:A6"), ("series", Series)), null);
        Mutations.Set(sheet, Props(("filter", "A1:C6"), ("name", "Sales 2025")));
        Assert.Equal("A2:A6", PathResolver.Single(doc.Root, "/sheet[1]/chart[1]").GetProps()["categories"]);
        using var package = SpreadsheetDocument.Open(new MemoryStream(Save(doc)), false);
        var formulas = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.Formula>().Select(f => f.Text).ToList();
        Assert.Contains("'Sales 2025'!$A$2:$A$6", formulas);
        Assert.DoesNotContain(formulas, f => f.StartsWith("Sheet1!"));
        Assert.Equal("'Sales 2025'!$A$1:$C$6", package.WorkbookPart.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text);
    }

    [Fact]
    public void Invalid_chart_input_is_a_validation_error_with_a_hint()
    {
        using var doc = Data();
        var sheet = doc.Root.Children[0];
        Check(() => Mutations.Add(sheet, "chart", Props(("type", "treemap")), null), "column | bar");
        Check(() => Mutations.Add(sheet, "chart", Props(("series", "[{\"values\":\"B2:\"}]")), null), "B2:B6");
        Check(() => Mutations.Add(sheet, "chart", Props(("series", "{\"values\":\"B2:B6\"}")), null), "B2:B6");
        Check(() => Mutations.Add(sheet, "chart", Props(("series", "[{\"name\":\"B1\"}]")), null), "values");
        Check(() => Mutations.Add(sheet, "chart", Props(("categories", "Nope!A2:A6"), ("series", "[{\"values\":\"B2:B6\"}]")), null), "Sheets: Sheet1");
        Check(() => Mutations.Add(sheet, "chart", Props(("series", "[{\"values\":\"B2:B6\"}]"), ("legend", "middle")), null), "none | right");
        Assert.Empty(PathResolver.Query(doc.Root, "//chart"));
        Assert.DoesNotContain(PackageCompare.Parts(Save(doc)).Keys, k => k.Contains("/charts/"));

        static void Check(Action act, string hint)
        {
            var ex = Assert.Throws<WriterException>(act);
            Assert.Equal(ErrorCode.Validation, ex.Code);
            Assert.Contains(hint, ex.Hint);
        }
    }

    [Fact]
    public void Help_describes_charts_and_sheet_layout()
    {
        var chart = Run("help", "xlsx", "chart");
        Assert.Contains("chart (xlsx): A chart on the worksheet.", chart);
        Assert.Contains("Goes under: sheet", chart);
        Assert.Contains("column | bar | line | pie | area | scatter | doughnut", chart);
        Assert.Contains("writer add file.xlsx //sheet[1] --type chart --prop type=\"column\"", chart);
        var sheet = Run("help", "xlsx", "sheet");
        foreach (var prop in new[] { "merges", "widths", "heights", "freeze", "filter" }) Assert.Contains($"  {prop,-9} ", sheet);
        Assert.Contains("A2 freezes row 1", sheet);
        Assert.Contains("chart", Run("help", "xlsx"));

        static string Run(params string[] argv)
        {
            var stdout = new StringWriter();
            Assert.Equal(0, Runner.Run(argv, stdout, new StringWriter()));
            return stdout.ToString();
        }
    }
}
