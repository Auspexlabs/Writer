using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Tests;

/// <summary>A cell's own borders side by side (top bottom left right), as the ribbon's 单元格边框 writes them, next to the kinds;
/// the table's 标题行 look and the gallery's orange style.</summary>
public class DocxTableFormatTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    [Fact]
    public void Cell_borders_take_sides_as_well_as_kinds()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        Mutations.Add(body, "table", Props(("rows", "2"), ("cols", "2"), ("style", "TableGrid")), null);
        Node Cell(int r, int c) => PathResolver.Single(doc.Root, $"/body/table[1]/row[{r}]/cell[{c}]");
        var cell = Mutations.Set(Cell(1, 1), Props(("borders", "top bottom")));
        Assert.Equal("top bottom", cell.GetProps()["borders"]);
        var raw = cell.GetRaw();
        Assert.Contains("<w:top w:val=\"single\"", raw);
        Assert.Contains("<w:bottom w:val=\"single\"", raw);
        Assert.Contains("<w:left w:val=\"none\"", raw);
        Assert.Contains("<w:right w:val=\"none\"", raw);
        Assert.Equal("left", Mutations.Set(Cell(1, 2), Props(("borders", "left"))).GetProps()["borders"]);
        Assert.Equal("all", Mutations.Set(Cell(2, 1), Props(("borders", "all"))).GetProps()["borders"]);
        Assert.Equal("none", Mutations.Set(Cell(2, 2), Props(("borders", "none"))).GetProps()["borders"]);
        Assert.Throws<WriterException>(() => Mutations.Set(Cell(2, 2), Props(("borders", "sideways"))));
        var row = Mutations.Set(PathResolver.Single(doc.Root, "/body/table[1]/row[1]"), Props(("header", "true"), ("height", "1cm")));
        Assert.Equal("true", row.GetProps()["header"]);
        Assert.Equal("1cm", Units.FormatLength(long.Parse(row.GetProps()["height"])));
    }

    [Fact]
    public void Header_look_turns_the_styles_first_row_off_and_on_in_both_forms_of_tblLook()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var table = Mutations.Add(body, "table", Props(("rows", "3"), ("cols", "2"), ("style", "GridTable4Accent2")), null);
        Assert.Contains(((DocxDocument)doc).Main.StyleDefinitionsPart!.Styles!.Elements<W.Style>(), s => s.StyleId == "GridTable4Accent2");
        table = Mutations.Set(table, Props(("header", "false")));
        Assert.Equal(("GridTable4Accent2", "false"), (table.GetProps()["style"], table.GetProps()["header"]));
        var look = ((W.Table)table.Anchor).GetFirstChild<W.TableProperties>()!.TableLook!;
        Assert.False(look.FirstRow!.Value);
        Assert.True(look.Val is null || (Convert.ToInt32(look.Val.Value, 16) & 0x20) == 0, "the older val keeps the same bit");
        table = Mutations.Set(table, Props(("header", "true")));
        Assert.Equal("true", table.GetProps()["header"]);
        look.FirstRow = null; look.Val = "0400"; // a file that only has the older form: bit 0x20 is the first row
        Assert.Equal("false", table.GetProps()["header"]);
    }
}
