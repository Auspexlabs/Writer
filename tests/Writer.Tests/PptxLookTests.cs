using Writer.Core;
using Writer.Formats.Pptx;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Tests;

/// <summary>What pptx nodes inherit ("computed"): theme colours with PowerPoint's modifiers, shape style references,
/// placeholder inheritance through the layout and master text styles, theme fonts, backgrounds and the colour map.
/// The blank deck's theme is Office's: accent1 4472C4, Calibri Light / Calibri, a bg1 background style.</summary>
public class PptxLookTests
{
    const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);
    static IReadOnlyDictionary<string, string> Computed(Node node) => node.GetComputed(node.GetProps()) ?? new Dictionary<string, string>();

    static (Document Doc, Node Slide) Deck(string layout = "Blank")
    {
        var doc = new PptxAdapter().Create();
        return (doc, Mutations.Add(doc.Root, "slide", Props(("layout", layout)), null));
    }

    [Theory]
    [InlineData("", "4472C4")]
    [InlineData("<a:lumMod val=\"75000\"/>", "2F5597")]                        // Accent 1, Darker 25%
    [InlineData("<a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/>", "8FAADC")] // Accent 1, Lighter 40%
    [InlineData("<a:lumMod val=\"20000\"/><a:lumOff val=\"80000\"/>", "DAE3F3")] // Accent 1, Lighter 80%
    public void Theme_colours_resolve_with_PowerPoints_modifiers(string modifiers, string expected)
    {
        var (doc, slide) = Deck();
        using var _ = doc;
        var shape = Mutations.Add(slide, "shape", Props(("geometry", "rect")), null);
        var spPr = ((P.Shape)shape.Anchor).ShapeProperties!;
        spPr.InsertAfter(new A.SolidFill($"<a:solidFill xmlns:a=\"{ANs}\"><a:schemeClr val=\"accent1\">{modifiers}</a:schemeClr></a:solidFill>"), spPr.GetFirstChild<A.PresetGeometry>());
        Assert.False(shape.GetProps().ContainsKey("fill"));
        Assert.Equal(expected, Computed(shape)["fill"]);
    }

    [Fact]
    public void Drawn_shapes_take_fill_line_and_text_from_their_style_references()
    {
        var (doc, slide) = Deck();
        using var _ = doc;
        var shape = Mutations.Add(slide, "shape", Props(("geometry", "ellipse")), null);
        var c = Computed(shape);
        Assert.Equal(("4472C4", "4472C4", "FFFFFF", "Calibri", "18"), (c["fill"], c["line"], c["color"], c["font"], c["size"]));
        Assert.Equal("Microsoft YaHei", c["fontEa"]);
    }

    [Fact]
    public void Text_boxes_have_no_fill_and_take_the_default_text_style()
    {
        var (doc, slide) = Deck();
        using var _ = doc;
        var box = Mutations.Add(slide, "shape", Props(("text", "Hello")), null);
        var c = Computed(box);
        Assert.False(c.ContainsKey("fill"));
        Assert.Equal(("000000", "Calibri", "18"), (c["color"], c["font"], c["size"]));
    }

    [Fact]
    public void Placeholders_inherit_through_the_layout_and_the_masters_text_styles()
    {
        var (doc, slide) = Deck("Title");
        using var _ = doc;
        var title = slide.Children.Single(n => n.GetProps().GetValueOrDefault("placeholder") == "title");
        var subtitle = slide.Children.Single(n => n.GetProps().GetValueOrDefault("placeholder") == "subtitle");
        var t = Computed(title);
        // colour and theme fonts from the master's titleStyle, size from the Title layout's own list style (60pt, not 44pt)
        Assert.Equal(("000000", "Calibri Light", "60"), (t["color"], t["font"], t["size"]));
        Assert.Equal("Microsoft YaHei", t["fontEa"]);
        var s = Computed(subtitle);
        Assert.Equal(("000000", "Calibri", "24"), (s["color"], s["font"], s["size"]));

        var (doc2, content) = Deck("Content");
        using var __ = doc2;
        var plainTitle = content.Children.Single(n => n.GetProps().GetValueOrDefault("placeholder") == "title");
        Assert.Equal("44", Computed(plainTitle)["size"]);
    }

    [Fact]
    public void Explicit_values_stay_in_props_and_are_not_repeated()
    {
        var (doc, slide) = Deck();
        using var _ = doc;
        var shape = Mutations.Add(slide, "shape", Props(("geometry", "rect"), ("text", "Hi")), null);
        shape = Mutations.Set(shape, Props(("fill", "FF0000"), ("color", "00FF00")));
        var c = Computed(shape);
        Assert.False(c.ContainsKey("fill"));
        Assert.False(c.ContainsKey("color"));
        Assert.Equal("4472C4", c["line"]);
    }

    [Fact]
    public void Slides_show_the_masters_background_through_the_colour_map()
    {
        var (doc, slide) = Deck("Title");
        using var _ = doc;
        Assert.Equal("FFFFFF", Computed(slide)["background"]);

        // a slide that swaps the map (dark on light becomes light on dark): bg1 → dk1, tx1 → lt1
        var sld = ((DocumentFormat.OpenXml.Packaging.SlidePart)slide.Anchor).Slide!;
        sld.ColorMapOverride = new P.ColorMapOverride(new A.OverrideColorMapping
        {
            Background1 = A.ColorSchemeIndexValues.Dark1, Text1 = A.ColorSchemeIndexValues.Light1,
            Background2 = A.ColorSchemeIndexValues.Dark2, Text2 = A.ColorSchemeIndexValues.Light2,
            Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3,
            Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6,
            Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
        });
        Assert.Equal("000000", Computed(slide)["background"]);
        var title = slide.Children.Single(n => n.GetProps().GetValueOrDefault("placeholder") == "title");
        Assert.Equal("FFFFFF", Computed(title)["color"]);

        Mutations.Set(slide, Props(("background", "123456")));
        Assert.False(Computed(doc.Root.Children[0]).ContainsKey("background"));
    }

    [Fact]
    public void Json_lists_computed_values_next_to_props_in_display_units()
    {
        var (doc, slide) = Deck();
        using var _ = doc;
        var shape = Mutations.Add(slide, "shape", Props(("geometry", "ellipse")), null);
        var json = System.Text.Json.JsonDocument.Parse(NodeJson.Serialize(shape, 0)).RootElement;
        Assert.Equal("4472C4", json.GetProperty("computed").GetProperty("fill").GetString());
        Assert.Equal("18pt", json.GetProperty("computed").GetProperty("size").GetString());
        Assert.False(json.GetProperty("props").TryGetProperty("fill", out var _fill));
    }
}
