using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Pptx;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Tests;

/// <summary>PowerPoint's layouts and the design edits: real slide layouts with placeholders (p:ph type/idx) that survive a save,
/// layouts made on demand, a layout change that moves placeholders as PowerPoint does, palettes, fonts, align/distribute, fit.</summary>
public class PptxLayoutTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);
    static Document Reopen(Document doc) { var ms = new MemoryStream(); doc.Save(ms); ms.Position = 0; return new PptxAdapter().Open(ms); }
    static List<string> Errors(Document doc) => new OpenXmlValidator(FileFormatVersions.Office2016).Validate(((PptxDocument)doc).Package)
        .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToList();
    /// <summary>The slide's placeholders as p:ph says them: type (or obj), idx; and whether each is empty.</summary>
    static List<string> Placeholders(Node slide) => ((SlidePart)slide.Anchor).Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>()
        .Select(s => (Ph: s.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape, Text: s.TextBody?.InnerText ?? ""))
        .Where(x => x.Ph is not null)
        .Select(x => $"{x.Ph!.Type?.InnerText ?? "obj"}{(x.Ph.Index is null ? "" : ":" + x.Ph.Index.Value)}{(x.Text.Length == 0 ? "" : "=" + x.Text)}").ToList();
    static Node Shape(Node slide, int n) => slide.Children.Where(c => c.Kind == "shape").ElementAt(n - 1);

    [Fact]
    public void Every_standard_layout_is_a_real_layout_whose_slides_carry_its_empty_placeholders()
    {
        using var doc = new PptxAdapter().Create();
        foreach (var spec in PptxTemplate.Layouts) Mutations.Add(doc.Root, "slide", Props(("layout", spec.Key)), null);
        Assert.Empty(Errors(doc));
        using var back = Reopen(doc);
        var slides = back.Root.Children;
        Assert.Equal(PptxTemplate.Layouts.Select(l => l.Name), slides.Select(s => s.GetProps()["layout"]));
        Assert.Equal(["ctrTitle", "subTitle:1"], Placeholders(slides[0]));
        Assert.Equal(["title", "obj:1"], Placeholders(slides[1]));
        Assert.Equal(["title", "body:1"], Placeholders(slides[2]));
        Assert.Equal(["title", "obj:1", "obj:2"], Placeholders(slides[3]));
        Assert.Equal(["title", "body:1", "obj:2", "body:3", "obj:4"], Placeholders(slides[4]));
        Assert.Equal(["title"], Placeholders(slides[5]));
        Assert.Empty(Placeholders(slides[6]));
        Assert.Equal(["title", "obj:1", "body:2"], Placeholders(slides[7]));
        Assert.Equal(["title", "pic:1", "body:2"], Placeholders(slides[8]));
        Assert.Equal(["title", "body:1"], Placeholders(slides[9]));
        // the layout's types, as PowerPoint names them, and the two columns where the slide editor draws them (1600-unit grid)
        var two = ((SlidePart)slides[3].Anchor).SlideLayoutPart!;
        Assert.Equal("twoObj", two.SlideLayout!.Type!.InnerText);
        Assert.Equal(("2.709cm", "17.611cm", "13.547cm"), (Registry.ToDisplay("shape", Shape(slides[3], 2).GetProps())["x"], Registry.ToDisplay("shape", Shape(slides[3], 3).GetProps())["x"], Registry.ToDisplay("shape", Shape(slides[3], 3).GetProps())["w"]));
        Assert.Null(((P.Shape)Shape(slides[3], 2).Anchor).ShapeProperties!.Transform2D); // placed by the layout, not by the slide
        Assert.Equal("quote", PptxTemplate.Standard("引用")!.Key);
        Assert.Equal("comparison", PptxTemplate.Standard("Comparison")!.Key);
    }

    [Fact]
    public void A_standard_layout_the_deck_lacks_is_made_in_its_master()
    {
        using var doc = new PptxAdapter().Open(new MemoryStream(File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("pptx"), "Mars-Settlement-Guide.pptx"))));
        var before = Errors(doc).ToHashSet();
        var layouts = ((PptxDocument)doc).Presentation.SlideMasterParts.Sum(m => m.SlideLayoutParts.Count());
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "quote")), null);
        Assert.Equal("Quote", slide.GetProps()["layout"]);
        Assert.Equal(["title", "body:1"], Placeholders(slide));
        var again = Mutations.Add(doc.Root, "slide", Props(("layout", "quote")), null); // the one it made, not another
        Assert.Equal(layouts + 1, ((PptxDocument)doc).Presentation.SlideMasterParts.Sum(m => m.SlideLayoutParts.Count()));
        var ids = ((PptxDocument)doc).Presentation.SlideMasterParts.SelectMany(m => m.SlideMaster!.SlideLayoutIdList!.Elements<P.SlideLayoutId>()).Select(l => l.Id!.Value).ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.Empty(Errors(doc).Except(before));
        using var back = Reopen(doc);
        Assert.Equal("Quote", back.Root.Children[^1].GetProps()["layout"]);
        _ = again;
    }

    [Fact]
    public void Changing_the_layout_moves_placeholders_as_PowerPoint_does()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "title")), null);
        Mutations.Set(Shape(slide, 1), Props(("text", "Q3 review")));
        Mutations.Set(Shape(slide, 2), Props(("text", "Numbers first")));
        Mutations.Set(Shape(slide, 2), Props(("x", "3cm"))); // moved by hand: the new layout places it again
        slide = Mutations.Set(slide, Props(("layout", "two")));
        Assert.Equal("Two Content", slide.GetProps()["layout"]);
        Assert.Equal(["title=Q3 review", "obj:1=Numbers first", "obj:2"], Placeholders(slide)); // the subtitle's text went to the first column
        Assert.Null(((P.Shape)Shape(slide, 2).Anchor).ShapeProperties!.Transform2D);
        Assert.Equal("Q3 review", slide.GetProps()["title"]);

        slide = Mutations.Set(slide, Props(("layout", "titleOnly")));
        // the second column was empty and goes; the first has text, so it stays where it was
        Assert.Equal(["title=Q3 review", "obj:1=Numbers first"], Placeholders(slide));
        Assert.Equal("2.709cm", Registry.ToDisplay("shape", Shape(slide, 2).GetProps())["x"]);
        Assert.Empty(Errors(doc));
    }

    [Fact]
    public void A_palette_and_fonts_dress_the_whole_deck()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "content"), ("title", "Plan")), null);
        var box = Mutations.Add(slide, "shape", Props(("text", "Note"), ("font", "Arial")), null);
        Mutations.Set(doc.Root, Props(("palette", "sea"), ("fonts", "Georgia,Verdana")));
        Assert.Equal(("sea", "Georgia,Verdana"), (doc.Root.GetProps()["palette"], doc.Root.GetProps()["fonts"]));
        slide = doc.Root.Children[0];
        Assert.Equal("1F3550", slide.GetComputed(slide.GetProps())!["background"]); // a dark palette: the colour map is swapped
        var title = Shape(slide, 1);
        Assert.Equal(("FFFFFF", "Georgia"), (title.GetComputed(title.GetProps())!["color"], title.GetComputed(title.GetProps())!["font"]));
        box = slide.Children.Last();
        Assert.False(box.GetProps().ContainsKey("font")); // the font set on the text is gone: the deck's body font applies
        Assert.Equal("Verdana", box.GetComputed(box.GetProps())!["font"]);

        Mutations.Set(doc.Root, Props(("palette", "paper")));
        slide = doc.Root.Children[0];
        Assert.Equal(("FFFFFF", "1D1D1F"), (slide.GetComputed(slide.GetProps())!["background"], Shape(slide, 1).GetComputed(Shape(slide, 1).GetProps())!["color"]));
        Assert.Empty(Errors(doc));
        using var back = Reopen(doc);
        Assert.Equal("paper", back.Root.GetProps()["palette"]);
    }

    [Fact]
    public void Align_and_distribute_line_shapes_up()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        foreach (var (x, y, w) in new[] { ("2cm", "3cm", "4cm"), ("10cm", "5cm", "2cm"), ("20cm", "4cm", "6cm") })
            Mutations.Add(slide, "shape", Props(("geometry", "rect"), ("x", x), ("y", y), ("w", w), ("h", "2cm")), null);
        string Box(int n, string k) => Registry.ToDisplay("shape", Shape(doc.Root.Children[0], n).GetProps())[k];

        Mutations.Set(slide, Props(("align", "top:/slide[1]/shape[1],shape[2],shape[3]")));
        Assert.Equal(["3cm", "3cm", "3cm"], new[] { Box(1, "y"), Box(2, "y"), Box(3, "y") });
        Mutations.Set(doc.Root.Children[0], Props(("distribute", "horizontal:shape[3],shape[1],shape[2]")));
        // 2..6, then a gap of (26 − 2 − 12) / 2 = 6: 12..14, then 20..26
        Assert.Equal(["2cm", "12cm", "20cm"], new[] { Box(1, "x"), Box(2, "x"), Box(3, "x") });
        Mutations.Set(doc.Root.Children[0], Props(("align", "center:shape[2]"))); // one shape: to the slide
        Assert.Equal("15.933cm", Box(2, "x")); // (33.867 − 2) / 2

        var bad = Assert.Throws<WriterException>(() => Mutations.Set(doc.Root.Children[0], Props(("align", "diagonal:shape[1]"))));
        Assert.Contains("left, center, right", bad.Hint);
        Assert.Throws<WriterException>(() => Mutations.Set(doc.Root.Children[0], Props(("align", "left"))));
    }

    [Fact]
    public void Overflowing_text_is_flagged_and_shrinks_to_fit()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "two")), null);
        var column = Shape(slide, 2);
        var html = string.Concat(Enumerable.Range(1, 7).Select(i => $"<p>第{i}点：市场规模持续扩大，用户增长超过预期，渠道结构发生了根本性的变化</p>"));
        column = Mutations.Set(column, Props(("html", html)));
        Assert.Equal("true", column.GetProps()["overflow"]);
        column = Mutations.Set(column, Props(("fit", "shrink")));
        Assert.False(column.GetProps().ContainsKey("overflow"));
        var size = double.Parse(column.Children[0].Children[0].GetProps()["size"], System.Globalization.CultureInfo.InvariantCulture);
        Assert.InRange(size, 10, 19); // down from the layout's 20pt, and every run alike
        Assert.All(column.Children.SelectMany(p => p.Children), r => Assert.Equal(size.ToString(System.Globalization.CultureInfo.InvariantCulture), r.GetProps()["size"]));
        Assert.False(Shape(slide, 1).GetProps().ContainsKey("overflow"));
    }
}
