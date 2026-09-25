using System.Text.RegularExpressions;
using Writer.Core;
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

    [Fact]
    public void Rejects_what_is_not_a_word_file()
    {
        var e = Assert.Throws<WriterException>(() => DocReader.Read(new byte[1024], []));
        Assert.Equal(ErrorCode.FormatError, e.Code);
    }
}
