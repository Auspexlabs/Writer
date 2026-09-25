using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Pptx;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Tests;

/// <summary>The shape gallery's file side: outlines with width, dash and arrowheads, shadows, gradients and rotation on p:sp;
/// lines and connectors as p:cxnSp that stick to shapes; groups as p:grpSp whose members read and move in slide units.</summary>
public class PptxShapesTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);
    static Document Reopen(Document doc) { var ms = new MemoryStream(); doc.Save(ms); ms.Position = 0; return new PptxAdapter().Open(ms); }
    static List<string> Errors(Document doc) => new OpenXmlValidator(FileFormatVersions.Office2016).Validate(((PptxDocument)doc).Package)
        .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToList();
    static string Cm(Node n, string k) => Registry.ToDisplay("shape", n.GetProps())[k];

    [Fact]
    public void A_shape_keeps_its_outline_shadow_gradient_rotation_and_aspect_lock()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        var star = Mutations.Add(slide, "shape", Props(("geometry", "star5"), ("x", "2cm"), ("y", "2cm"), ("w", "6cm"), ("h", "6cm"), ("text", "Hi"),
            ("line", "1F2937"), ("lineWidth", "2.25pt"), ("dash", "dash"), ("shadow", "true"), ("gradient", "4472C4,ED7D31,45"), ("rotation", "30"), ("lockAspect", "true")), null);
        Assert.Empty(Errors(doc));
        using var back = Reopen(doc);
        var p = back.Root.Children[0].Children.Single(c => c.Kind == "shape").GetProps();
        Assert.Equal(("star5", "1F2937", "2.25", "dash", "true", "4472C4,ED7D31,45", "30", "true"),
            (p["geometry"], p["line"], p["lineWidth"], p["dash"], p["shadow"], p["gradient"], p["rotation"], p["lockAspect"]));
        var sp = (P.Shape)back.Root.Children[0].Children.Single(c => c.Kind == "shape").Anchor;
        var spPr = sp.ShapeProperties!;
        Assert.Equal(["xfrm", "prstGeom", "gradFill", "ln", "effectLst"], spPr.ChildElements.Select(e => e.LocalName)); // DrawingML's order
        Assert.Equal(["solidFill", "prstDash"], spPr.GetFirstChild<A.Outline>()!.ChildElements.Select(e => e.LocalName));
        var shape = back.Root.Children[0].Children.Single(c => c.Kind == "shape");
        Mutations.Set(shape, Props(("gradient", ""), ("fill", "FF0000"), ("shadow", "false"), ("dash", "solid")));
        var q = shape.GetProps();
        Assert.Equal("FF0000", q["fill"]);
        Assert.False(q.ContainsKey("gradient") || q.ContainsKey("shadow") || q.ContainsKey("dash"));
    }

    [Fact]
    public void Connectors_are_cxnSp_with_arrowheads_that_stick_to_shapes()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        var a = Mutations.Add(slide, "shape", Props(("geometry", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "4cm"), ("h", "2cm")), null);
        var b = Mutations.Add(slide, "shape", Props(("geometry", "ellipse"), ("x", "12cm"), ("y", "8cm"), ("w", "4cm"), ("h", "2cm")), null);
        var line = Mutations.Add(slide, "connector", Props(("geometry", "straightConnector1"), ("x", "6cm"), ("y", "3cm"), ("w", "6cm"), ("h", "6cm"),
            ("line", "C00000"), ("lineWidth", "1.5pt"), ("tail", "triangle"), ("start", a.GetProps()["id"] + ",3"), ("end", b.GetProps()["id"] + ",1")), null);
        Assert.Equal("/slide[1]/connector[1]", line.Path);
        Assert.Empty(Errors(doc));
        using var back = Reopen(doc);
        var kinds = back.Root.Children[0].Children.Select(c => c.Kind).ToList();
        Assert.Equal(["shape", "shape", "connector"], kinds);
        var cxn = back.Root.Children[0].Children[2];
        var p = cxn.GetProps();
        Assert.Equal(("straightConnector1", "6cm", "6cm", "C00000", "1.5", "triangle", "2,3", "3,1"), (p["geometry"], Cm(cxn, "x"), Cm(cxn, "w"), p["line"], p["lineWidth"], p["tail"], p["start"], p["end"]));
        Assert.False(p.ContainsKey("head"));
        Assert.IsType<P.ConnectionShape>(cxn.Anchor);
        Mutations.Set(cxn, Props(("flipV", "true"), ("head", "oval"), ("end", ""), ("dash", "sysDot")));
        p = cxn.GetProps();
        Assert.Equal(("true", "oval", "sysDot"), (p["flipV"], p["head"], p["dash"]));
        Assert.False(p.ContainsKey("end"));
        var copy = Mutations.Copy(cxn, back.Root.Children[0], null);
        Assert.Equal("connector", copy.Kind);
        Assert.NotEqual(p["id"], copy.GetProps()["id"]);
        Assert.Empty(Errors(back));
    }

    [Fact]
    public void Groups_hold_members_in_slide_units_move_them_along_and_ungroup_in_place()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        Mutations.Add(slide, "shape", Props(("geometry", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "4cm"), ("h", "2cm"), ("text", "A")), null);
        Mutations.Add(slide, "shape", Props(("geometry", "ellipse"), ("x", "8cm"), ("y", "5cm"), ("w", "2cm"), ("h", "2cm"), ("text", "B")), null);
        Mutations.Add(slide, "shape", Props(("text", "loose"), ("x", "20cm"), ("y", "2cm"), ("w", "4cm"), ("h", "2cm")), null);
        var group = Mutations.Add(slide, "group", Props(("members", "shape[1],shape[2]")), null);
        Assert.Equal("/slide[1]/group[1]", group.Path);
        Assert.Equal(["shape", "group"], slide.Children.Select(c => c.Kind));
        Assert.Equal(("2cm", "2cm", "8cm", "5cm"), (Cm(group, "x"), Cm(group, "y"), Cm(group, "w"), Cm(group, "h")));
        Assert.Equal(["/slide[1]/group[1]/shape[1]", "/slide[1]/group[1]/shape[2]"], group.Children.Select(c => c.Path));
        Assert.Equal("8cm", Cm(group.Children[1], "x"));
        Assert.Empty(Errors(doc));

        // the group moves and halves: its members follow, and read where they now show
        group = Mutations.Set(group, Props(("x", "4cm"), ("w", "4cm"), ("h", "2.5cm")));
        Assert.Equal(("4cm", "2cm", "4cm", "2.5cm"), (Cm(group, "x"), Cm(group, "y"), Cm(group, "w"), Cm(group, "h")));
        var b = group.Children[1];
        Assert.Equal(("7cm", "3.5cm", "1cm", "1cm"), (Cm(b, "x"), Cm(b, "y"), Cm(b, "w"), Cm(b, "h")));
        // a member set in slide units lands there
        b = Mutations.Set(b, Props(("x", "6cm")));
        Assert.Equal("6cm", Cm(b, "x"));
        Assert.Equal("6cm", Registry.ToDisplay("shape", new Dictionary<string, string> { ["x"] = ((P.Shape)b.Anchor).ShapeProperties!.Transform2D!.Offset!.X!.Value.ToString() })["x"]); // child space: the group maps x' = 3 + x / 2

        using var back = Reopen(doc);
        var s = back.Root.Children[0];
        Assert.Equal(["shape", "group"], s.Children.Select(c => c.Kind));
        Assert.Equal("B", s.Children[1].Children[1].GetProps()["text"]);
        Mutations.Set(s.Children[1], Props(("ungroup", "true")));
        s = back.Root.Children[0];
        Assert.Equal(["shape", "shape", "shape"], s.Children.Select(c => c.Kind));
        Assert.Equal(["loose", "A", "B"], s.Children.Select(c => c.GetProps()["text"]));
        Assert.Equal(("6cm", "3.5cm", "1cm"), (Cm(s.Children[2], "x"), Cm(s.Children[2], "y"), Cm(s.Children[2], "w")));
        Assert.Empty(Errors(back));
    }

    [Fact]
    public void A_text_box_keeps_its_spacing_columns_autofit_direction_and_WordArt_effects()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        var box = Mutations.Add(slide, "shape", Props(("text", "One\nTwo"), ("lineSpacing", "1.5"), ("spaceBefore", "6pt"), ("spaceAfter", "3"), ("charSpacing", "1.5"),
            ("columns", "2"), ("autofit", "shrink:85"), ("direction", "eaVert"), ("textOutline", "1F2937"), ("textShadow", "true"), ("textGradient", "4472C4,ED7D31,90")), null);
        Mutations.Set(box.Children[1], Props(("list", "number"), ("level", "1")));
        Assert.Empty(Errors(doc));
        using var back = Reopen(doc);
        var shape = back.Root.Children[0].Children.Single(c => c.Kind == "shape");
        var p = shape.GetProps();
        Assert.Equal(("1.5", "6pt", "3pt", "1.5", "2", "shrink:85", "eaVert", "1F2937", "true", "4472C4,ED7D31,90"),
            (p["lineSpacing"], p["spaceBefore"], p["spaceAfter"], p["charSpacing"], p["columns"], p["autofit"], p["direction"], p["textOutline"], p["textShadow"], p["textGradient"]));
        var second = shape.Children[1].GetProps();
        Assert.Equal(("number", "1", "Two"), (second["list"], second["level"], second["text"]));
        var rPr = ((P.Shape)shape.Anchor).TextBody!.Elements<A.Paragraph>().First().Elements<A.Run>().First().RunProperties!;
        Assert.Equal(["ln", "gradFill", "effectLst"], rPr.ChildElements.Select(e => e.LocalName)); // DrawingML's order
        Mutations.Set(shape, Props(("text", "Typed again"))); // a rewrite keeps the settings: they ride on the first paragraph and run
        p = shape.GetProps();
        Assert.Equal(("1.5", "1.5", "4472C4,ED7D31,90"), (p["lineSpacing"], p["charSpacing"], p["textGradient"]));
        Mutations.Set(shape, Props(("textGradient", ""), ("autofit", "none"), ("columns", "1"), ("direction", "horz"), ("lineSpacing", "1")));
        p = shape.GetProps();
        Assert.Equal("4472C4", p["color"]); // the letters keep the gradient's first colour
        Assert.False(p.ContainsKey("textGradient") || p.ContainsKey("autofit") || p.ContainsKey("columns") || p.ContainsKey("direction") || p.ContainsKey("lineSpacing"));
        Assert.Empty(Errors(back));
    }

    [Fact]
    public void A_table_keeps_its_style_flags_column_widths_merges_and_cell_lines()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        var table = Mutations.Add(slide, "table", Props(("data", "[[\"a\",\"b\",\"c\"],[\"d\",\"e\",\"f\"],[\"g\",\"h\",\"i\"]]"), ("style", "LightStyle2Accent1"), ("header", "true"), ("banded", "false"), ("firstCol", "true"), ("widths", "[\"3cm\",\"5cm\",\"4cm\"]")), null);
        var p = table.GetProps();
        Assert.Equal(("LightStyle2Accent1", "true", "true", "[\"3cm\",\"5cm\",\"4cm\"]", "12cm"), (p["style"], p["header"], p["firstCol"], p["widths"], Registry.ToDisplay("table", p)["w"]));
        Assert.False(p.ContainsKey("banded"));
        Node Cell(int r, int c) => table.Children[r].Children[c];
        Mutations.Set(Cell(0, 0), Props(("colspan", "2"), ("rowspan", "2"), ("fill", "D9E2F3"), ("line", "1F2937")));
        Assert.Empty(Errors(doc));
        using var back = Reopen(doc);
        table = back.Root.Children[0].Children.Single(c => c.Kind == "table");
        var a = Cell(0, 0).GetProps();
        Assert.Equal(("2", "2", "D9E2F3", "1F2937"), (a["colspan"], a["rowspan"], a["fill"], a["line"]));
        Assert.Equal(["", "true", "true", "true"], new[] { Cell(0, 2), Cell(0, 1), Cell(1, 0), Cell(1, 1) }.Select(c => c.GetProps().GetValueOrDefault("covered", "")));
        Assert.Equal("", Cell(1, 1).GetProps()["text"]); // a covered cell keeps no text
        var tc = (A.TableCell)Cell(1, 1).Anchor;
        Assert.True(tc.HorizontalMerge?.Value == true && tc.VerticalMerge?.Value == true);
        Mutations.Set(Cell(0, 0), Props(("colspan", "1"), ("rowspan", "1")));
        Assert.All(new[] { Cell(0, 1), Cell(1, 0), Cell(1, 1) }, c => Assert.False(c.GetProps().ContainsKey("covered")));
        Mutations.Set(table, Props(("w", "24cm")));
        Assert.Equal("[\"6cm\",\"10cm\",\"8cm\"]", table.GetProps()["widths"]); // the columns keep their shares
        Assert.Equal("{69012ECD-51FC-41F1-AA8D-1B2483CD663E}", ((P.GraphicFrame)table.Anchor).Graphic!.GraphicData!.GetFirstChild<A.Table>()!.TableProperties!.GetFirstChild<A.TableStyleId>()!.Text); // Light Style 2 - Accent 1
        Mutations.Set(table, Props(("rows", "4"))); // a row added at the end starts clean: no span or merge mark comes along
        Assert.All(table.Children[3].Children, c => Assert.False(c.GetProps().ContainsKey("colspan") || c.GetProps().ContainsKey("covered")));
        Assert.Empty(Errors(back));
    }

    [Fact]
    public void A_deck_from_PowerPoint_reads_its_connectors_and_groups()
    {
        using var doc = new PptxAdapter().Open(new MemoryStream(File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("pptx"), "Mars-Settlement-Guide.pptx"))));
        var all = doc.Root.Children.SelectMany(s => s.Children).ToList();
        foreach (var n in all.Where(n => n.Kind is "connector" or "group")) Assert.True(n.GetProps().ContainsKey("w"), n.Path);
        Assert.All(all, n => Assert.Contains(n.Kind, new[] { "decor", "shape", "image", "table", "connector", "group" }));
    }
}
