using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
using Writer.Formats.Html;
using Writer.Formats.Xlsx;
using A = DocumentFormat.OpenXml.Drawing;

namespace Writer.Tests;

/// <summary>What a workbook looks like: theme and indexed colours resolved, grid lines and the default font read, and the HTML view
/// drawing cells, number formats and charts as Excel does.</summary>
public class XlsxLookTests
{
    const string ThemeXml = """
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="T">
        <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
        <a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
        <a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>
        <a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink>
        <a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>
        """;

    const string StylesXml = """
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="2"><font><sz val="10"/><color theme="1"/><name val="Arial"/></font><font><b/><sz val="10"/><color theme="4" tint="-0.249977111117893"/><name val="Arial"/></font></fonts>
        <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.5999938962981048"/><bgColor indexed="64"/></patternFill></fill></fills>
        <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="medium"><color indexed="10"/></bottom><diagonal/></border></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs>
        </styleSheet>
        """;

    static Document Themed()
    {
        var ms = new MemoryStream();
        using (var package = SpreadsheetDocument.Create(ms, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbook = package.AddWorkbookPart();
            workbook.Workbook = new Workbook(new Sheets());
            var sheet = workbook.AddNewPart<WorksheetPart>();
            sheet.Worksheet = new Worksheet(new SheetViews(new SheetView { WorkbookViewId = 0U, ShowGridLines = false }),
                new SheetData(new Row(new Cell { CellReference = "A1", StyleIndex = 1U, CellValue = new CellValue("5") }, new Cell { CellReference = "B1", CellValue = new CellValue("6") }) { RowIndex = 1U }));
            workbook.Workbook.Sheets!.Append(new Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1U, Name = "S" });
            workbook.AddNewPart<ThemePart>().Theme = new A.Theme(ThemeXml);
            workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(StylesXml);
        }
        ms.Position = 0;
        return new XlsxAdapter().Open(ms);
    }

    static IReadOnlyDictionary<string, string> Get(Document doc, string path) => PathResolver.Single(doc.Root, path).GetProps();

    [Fact]
    public void Theme_and_indexed_colours_read_as_they_show_and_go_back_as_theme_references()
    {
        using var doc = Themed();
        var a1 = Get(doc, "/sheet[1]/cell[A1]");
        Assert.Equal(("376092", "B9CDE5"), (a1["color"], a1["fill"])); // accent 1 darker 25% / lighter 60%
        Assert.Equal(("medium", "FF0000"), (a1["border"], a1["borderColor"])); // indexed 10 is red
        Assert.Equal(("Arial", "10"), (Get(doc, "/")["font"], Get(doc, "/")["size"]));

        // B1 takes A1's look as the editor sends it (hex): the stylesheet gets no rgb twin of the theme colours
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[B1]"), new Dictionary<string, string> { ["fill"] = "B9CDE5", ["color"] = "376092", ["bold"] = "true" });
        var styles = ((XlsxDocument)doc).Workbook.WorkbookStylesPart!.Stylesheet!;
        Assert.DoesNotContain(styles.Descendants<ColorType>(), c => c.Rgb?.Value is "FFB9CDE5" or "FF376092");
        CellFormat Xf(string cell) => styles.CellFormats!.Elements<CellFormat>().ElementAt((int)((Cell)PathResolver.Single(doc.Root, "/sheet[1]/cell[" + cell + "]").Anchor).StyleIndex!.Value);
        Assert.Equal((Xf("A1").FontId!.Value, Xf("A1").FillId!.Value), (Xf("B1").FontId!.Value, Xf("B1").FillId!.Value)); // A1's own font and fill entries
    }

    [Fact]
    public void Grid_lines_read_and_write_on_the_sheet_view()
    {
        using var doc = Themed();
        var sheet = PathResolver.Single(doc.Root, "/sheet[1]");
        Assert.Equal("false", sheet.GetProps()["gridlines"]);
        Mutations.Set(sheet, new Dictionary<string, string> { ["gridlines"] = "true" });
        Assert.False(PathResolver.Single(doc.Root, "/sheet[1]").GetProps().ContainsKey("gridlines"));
        using var blank = new XlsxAdapter().Create();
        Mutations.Set(PathResolver.Single(blank.Root, "/sheet[1]"), new Dictionary<string, string> { ["gridlines"] = "false" });
        Assert.Equal("false", PathResolver.Single(blank.Root, "/sheet[1]").GetProps()["gridlines"]);
    }

    [Theory]
    [InlineData(1234.5, "#,##0.00", "1,234.50")]
    [InlineData(0.125, "0.0%", "12.5%")]
    [InlineData(-5, "\"$\"#,##0", "-$5")]
    [InlineData(-1234.5, "#,##0.00_);(#,##0.00)", "(1,234.50)")]
    [InlineData(1500000, "#,##0,\"K\"", "1,500K")]
    [InlineData(12345.678, "0.00E+00", "1.23E+04")]
    [InlineData(45306.4375, "yyyy-mm-dd hh:mm", "2024-01-15 10:30")]
    [InlineData(45306.75, "m/d/yyyy h:mm AM/PM", "1/15/2024 6:00 PM")]
    [InlineData(45306, "[$-409]mmmm d, yyyy;@", "January 15, 2024")]
    [InlineData(45306, "yyyy\"年\"m\"月\"d\"日\" aaaa", "2024年1月15日 星期一")]
    [InlineData(1.5, "[h]:mm:ss", "36:00:00")]
    [InlineData(1, "m/d/yyyy", "1/1/1900")]
    [InlineData(42, "\"No. \"@", "No. 42")]
    [InlineData(0, "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_)", " $- ")]
    [InlineData(7, "[Red]General", "7")]
    [InlineData(1.5, "# ?/?", "1 1/2")]
    [InlineData(3.14159, "# ??/??", "3 14/99")]
    [InlineData(0.75, "?/4", "3/4")]
    public void Number_formats_show_as_in_Excel(double value, string code, string shown) =>
        Assert.Equal(shown, XlsxNumberFormat.Format(value, code).Text);

    [Fact]
    public void A_section_colour_comes_with_the_text() =>
        Assert.Equal(("5", "FF0000"), XlsxNumberFormat.Format(-5, "0;[Red]0"));

    [Fact]
    public void The_html_view_draws_the_sheet_as_Excel_does()
    {
        using var doc = new XlsxAdapter().Open(new MemoryStream(File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("xlsx"), "budget-tracker.xlsx"))));
        var html = HtmlWriter.Render(doc);
        Assert.Contains("<col style=\"width:112px\">", html); // width 16 → Excel's 112 px
        Assert.Contains("<td colspan=\"8\" style=\"background:#2D5016;color:#FFFFFF;font-weight:bold;font-size:18pt;text-align:center;vertical-align:middle;\">2026 Annual Budget Tracker</td>", html);
        Assert.Contains(">$2,500,000</td>", html);
        Assert.Contains(">75%</td>", html);
        Assert.Contains("border-bottom:1px solid #000000", html);
        Assert.Contains("<tr style=\"height:53px\">", html); // 40 pt
        Assert.Matches("<svg class=\"chart\"[^>]*>.*Department Quarterly Spending.*<rect x=\"[\\d.]+\" y=\"[\\d.]+\" width=\"[\\d.]+\" height=\"[\\d.]+\" fill=\"#5B9BD5\"/>", html);
        Assert.Contains(">Engineering</text>", html);

        using var themed = Themed();
        var view = HtmlWriter.Render(themed);
        Assert.Contains("<table class=\"grid\"", view); // grid lines off
        Assert.Contains("font-family:'Arial'", view);
        Assert.Contains("border-bottom:2px solid #FF0000", view);
    }

    /// <summary>A workbook with dates in A1:A3 (yyyy-mm-dd), a date-time in A4 and a percent and a currency cell, in either date system.</summary>
    static Document Dated(bool date1904, params string[] serials)
    {
        var ms = new MemoryStream();
        using (var package = SpreadsheetDocument.Create(ms, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbook = package.AddWorkbookPart();
            workbook.Workbook = new Workbook(new WorkbookProperties { Date1904 = date1904 ? true : null }, new Sheets());
            var sheet = workbook.AddNewPart<WorksheetPart>();
            var data = new SheetData();
            for (var i = 0; i < serials.Length; i++)
                data.Append(new Row(new Cell { CellReference = "A" + (i + 1), StyleIndex = 1U, CellValue = new CellValue(serials[i]) }) { RowIndex = (uint)i + 1 });
            data.Append(new Row(new Cell { CellReference = "B9", StyleIndex = 2U, CellValue = new CellValue("0.1255") }, new Cell { CellReference = "C9", StyleIndex = 3U, CellValue = new CellValue("-1234.5") }) { RowIndex = 9U });
            sheet.Worksheet = new Worksheet(data);
            workbook.Workbook.Sheets!.Append(new Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1U, Name = "S" });
            workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(XlsxStyles.MinimalXml.Replace("<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>",
                "<cellXfs count=\"4\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/><xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
                + "<xf numFmtId=\"10\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs>")
                .Replace("<fonts", "<numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/><numFmt numFmtId=\"165\" formatCode=\"&quot;$&quot;#,##0.00\"/></numFmts><fonts"));
        }
        ms.Position = 0;
        return new XlsxAdapter().Open(ms);
    }

    static string Serial(Document doc, string cell) => ((Cell)PathResolver.Single(doc.Root, "/sheet[1]/cell[" + cell + "]").Anchor).CellValue!.Text;

    static void SetValue(Document doc, string cell, string value) =>
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[" + cell + "]"), new Dictionary<string, string> { ["value"] = value });

    [Fact]
    public void A_1904_workbook_reads_and_writes_its_own_serials()
    {
        using var doc = Dated(true, "0", "1462", "43894.5");
        Assert.Equal(("1904-01-01", "1908-01-02", "2024-03-05 12:00:00"), (Get(doc, "/sheet[1]/cell[A1]")["value"], Get(doc, "/sheet[1]/cell[A2]")["value"], Get(doc, "/sheet[1]/cell[A3]")["value"]));
        SetValue(doc, "A4", "2024-03-05 12:00:00"); // what the editor sends for a typed date-time
        Assert.Equal("43894.5", Serial(doc, "A4"));
        SetValue(doc, "A1", "x");
        SetValue(doc, "A1", "1904-01-01"); // changed and changed back: the same serial
        Assert.Equal("0", Serial(doc, "A1"));
        Assert.Contains(">2024-03-05</td>", HtmlWriter.Render(doc));
    }

    [Fact]
    public void The_1900_system_counts_as_Excel_does_before_March_1900()
    {
        using var doc = Dated(false, "1", "59", "61");
        Assert.Equal(("1900-01-01", "1900-02-28", "1900-03-01"), (Get(doc, "/sheet[1]/cell[A1]")["value"], Get(doc, "/sheet[1]/cell[A2]")["value"], Get(doc, "/sheet[1]/cell[A3]")["value"]));
        SetValue(doc, "A4", "1900-01-01");
        Assert.Equal("1", Serial(doc, "A4"));
        SetValue(doc, "A4", "2024-03-05 12:00:00");
        Assert.Equal("45356.5", Serial(doc, "A4"));
    }

    [Fact]
    public void Percent_and_currency_cells_keep_value_and_code_through_an_edit()
    {
        using var doc = Dated(false, "45356");
        SetValue(doc, "B9", "0.125");
        SetValue(doc, "C9", "-1234.5");
        Assert.Equal(("0.125", "0.00%", "-1234.5", "\"$\"#,##0.00"), (Serial(doc, "B9"), Get(doc, "/sheet[1]/cell[B9]")["format"], Serial(doc, "C9"), Get(doc, "/sheet[1]/cell[C9]")["format"]));
        var html = HtmlWriter.Render(doc);
        Assert.Contains(">12.50%</td>", html);
        Assert.Contains(">-$1,234.50</td>", html);
        Assert.Contains(">2024-03-05</td>", html);
    }
}
