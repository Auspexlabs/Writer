using Writer.Core;
using Writer.Formats.Docx;

namespace Writer.Tests;

/// <summary>A cell's own borders side by side (top bottom left right), as the ribbon's 单元格边框 writes them, next to the kinds.</summary>
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
}
