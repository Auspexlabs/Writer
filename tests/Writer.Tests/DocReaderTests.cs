using System.Text.RegularExpressions;
using Writer.Core;
using Writer.Formats;
using Writer.Formats.Compat;

namespace Writer.Tests;

public class DocReaderTests
{
    static List<Block> ReadSample(out List<string> warnings)
    {
        warnings = [];
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "compat", "sample.doc");
        return DocReader.Read(File.ReadAllBytes(path), warnings);
    }

    static string Strip(string html) => Regex.Replace(html, "<[^>]+>", "");

    [Fact]
    public void Reads_headings_runs_lists_and_tables_from_a_word_97_file()
    {
        var blocks = ReadSample(out var warnings);

        var headings = blocks.Where(b => b.Kind == BlockKind.Heading).ToList();
        Assert.Equal(2, headings.Count);
        Assert.Equal(1, headings[0].Level);
        Assert.Equal("Chapter One", Strip(headings[0].Html));
        Assert.Equal(2, headings[1].Level);
        Assert.Equal("Second level 标题", Strip(headings[1].Html));

        var runs = blocks.Single(b => b.Kind == BlockKind.Paragraph && b.Html.Contains("bold"));
        Assert.Contains("<b>bold</b>", runs.Html);
        Assert.Contains("<i>italic</i>", runs.Html);
        Assert.Matches("<span style=\"color:#F[0-9A-F]{5};[^\"]*\">red</span>", runs.Html); // textutil rounds the red a little
        Assert.Contains("<u>link</u>", runs.Html); // textutil writes no HYPERLINK field, only the blue underlined text

        Assert.Equal("center", blocks.Single(b => b.Html.Contains("Centered")).Align);

        var bullets = blocks.Where(b => b.List == "bullet").ToList();
        Assert.Equal(["First bullet", "Second bullet"], bullets.Select(b => Strip(b.Html)));
        var numbers = blocks.Where(b => b.List == "number").ToList();
        Assert.Equal(["Step one", "Step two"], numbers.Select(b => Strip(b.Html)));

        var table = blocks.Single(b => b.Kind == BlockKind.Table);
        Assert.Equal(2, table.Rows!.Count);
        Assert.All(table.Rows, r => Assert.Equal(2, r.Count));
        Assert.Equal("A1", Strip(table.Rows[0][0].Html));
        Assert.Equal("B1", Strip(table.Rows[0][1].Html));
        Assert.Equal("A2", Strip(table.Rows[1][0].Html));
        Assert.Equal("B2", Strip(table.Rows[1][1].Html));

        // order survives: heading, runs, heading, centered, lists, table, closing paragraphs
        Assert.True(blocks.IndexOf(table) > blocks.IndexOf(numbers[1]));
        Assert.Equal("The end.", Strip(blocks[^1].Html));
        Assert.DoesNotContain(warnings, w => w.Contains("could not be read"));
    }

    static List<Block> ReadForm(out List<string> warnings, out PageModel? page)
    {
        warnings = [];
        return DocReader.Read(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "compat", "form-table.doc")), warnings, out page);
    }

    [Fact]
    public void A_form_keeps_its_rows_its_merged_cells_and_its_column_widths()
    {
        // form-table.doc: LibreOffice's Word 97 save of form-table.doc.html, a 13-column form. A row of that many cells keeps its
        // cell definitions and the mark that ends it in the Data stream (sprmPHugePapx): read without it, the five rows were one
        // row of sixty cells, drawn as that many hair-thin columns.
        var table = ReadForm(out var warnings, out _).Single(b => b.Kind == BlockKind.Table);
        Assert.Equal(5, table.Rows!.Count);
        Assert.Equal(("员工信息登记表", 13), (Strip(table.Rows[0][0].Html), table.Rows[0][0].ColSpan)); // one cell across the form
        Assert.Equal(("center", "middle"), (table.Rows[0][0].Align, table.Rows[0][0].VAlign)); // its paragraph's and the cell's alignment
        Assert.Single(table.Rows[0]);
        Assert.Equal(13, table.Rows[1].Count);
        Assert.Equal(("照片", 3), (Strip(table.Rows[1][12].Html), table.Rows[1][12].RowSpan)); // merged down: the rows under it have 12
        Assert.Equal([12, 12], table.Rows.Skip(2).Take(2).Select(r => r.Count));
        Assert.Equal(("合计", 3), (Strip(table.Rows[4][0].Html), table.Rows[4][0].ColSpan));
        Assert.Equal(11, table.Rows[4].Count);
        Assert.All(table.Rows.SelectMany(r => r), c => Assert.Null(c.Fill)); // automatic shading (0xFF000000) is no fill, not black
        Assert.Equal(13, table.ColumnsCm!.Count);
        Assert.InRange(table.ColumnsCm.Sum(), 16, 16.5);
        Assert.DoesNotContain(warnings, w => w.Contains("could not be read"));
    }

    [Fact]
    public void The_page_is_the_first_sections_size_orientation_and_margins()
    {
        ReadForm(out _, out var page);
        Assert.Equal(new PageModel(29.7, 21, 2, 2.5, 2, 2.5), page); // A4 landscape, as the @page rule set it up

        using var converted = Adapters.ForPath("form-table.doc").Open(File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Fixtures", "compat", "form-table.doc")));
        var root = converted.Root.GetProps();
        Assert.Equal(("A4", "landscape", "2cm 2.5cm 2cm 2.5cm"), (root["page"], root["orientation"], root["margin"]));
        var table = converted.Root.Children.First(c => c.Kind == "body").Children.Single(c => c.Kind == "table").GetProps();
        Assert.Equal(("5", "13"), (table["rows"], table["cols"]));
        Assert.StartsWith("[\"1.22cm\",", table["widths"]);
    }

    [Fact]
    public void Rejects_what_is_not_a_word_file()
    {
        var e = Assert.Throws<WriterException>(() => DocReader.Read(new byte[1024], []));
        Assert.Equal(ErrorCode.FormatError, e.Code);
    }
}
