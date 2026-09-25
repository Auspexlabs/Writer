using Writer.Core;
using Writer.Formats.Pptx;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class PptxTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    [Fact]
    public void Blank_deck_has_layouts_and_grows_slides()
    {
        using var doc = new PptxAdapter().Create();
        Assert.Equal("0", doc.Root.GetProps()["slides"]);
        var title = Mutations.Add(doc.Root, "slide", Props(("layout", "Title"), ("title", "Q4 Results")), null);
        Assert.Equal("/slide[1]", title.Path);
        Assert.Equal(("Q4 Results", "Title Slide"), (title.GetProps()["title"], title.GetProps()["layout"]));
        var shapes = title.Children.Where(c => c.Kind == "shape").ToList();
        Assert.Equal(2, shapes.Count);
        Assert.Equal("title", shapes[0].GetProps()["placeholder"]);
        Assert.Equal("subtitle", shapes[1].GetProps()["placeholder"]);
        Assert.Equal("2.709cm", Registry.ToDisplay("shape", shapes[0].GetProps())["x"]);

        var content = Mutations.Add(doc.Root, "slide", Props(("title", "Agenda")), null);
        Assert.Equal("Title and Content", content.GetProps()["layout"]);
        var blank = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), 1);
        Assert.Equal("/slide[1]", blank.Path);
        Assert.Empty(blank.Children);
        Assert.Equal(new[] { "Blank", "Title Slide", "Title and Content" }, doc.Root.Children.Select(s => s.GetProps()["layout"]));
        Assert.Equal("3", doc.Root.GetProps()["slides"]);
        var ex = Assert.Throws<WriterException>(() => Mutations.Add(doc.Root, "slide", Props(("layout", "Fancy")), null));
        Assert.Contains("Title Slide, Title and Content", ex.Hint);
        Assert.Contains("two, comparison", ex.Hint);

        var moved = Mutations.Move(PathResolver.Single(doc.Root, "/slide[3]"), doc.Root, 1);
        Assert.Equal("/slide[1]", moved.Path);
        Assert.Equal("Agenda", moved.GetProps()["title"]);
        PathResolver.Single(doc.Root, "/slide[2]").Remove();
        Assert.Equal(new[] { "Title and Content", "Title Slide" }, doc.Root.Children.Select(s => s.GetProps()["layout"]));

        using var reopened = new PptxAdapter().Open(new MemoryStream(Save(doc)));
        Assert.Equal("Agenda", reopened.Root.Children[0].GetProps()["title"]);
    }

    [Fact]
    public void Shapes_have_geometry_position_fill_and_text()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "Blank"), ("background", "1A1A2E")), null);
        Assert.Equal("1A1A2E", slide.GetProps()["background"]);
        var box = Mutations.Add(slide, "shape", Props(("md", "Revenue **grew** 25%\n\nSecond line"), ("x", "2cm"), ("y", "5cm"), ("w", "10cm"), ("h", "2cm"), ("font", "Arial"), ("size", "24"), ("color", "FFFFFF")), null);
        var shown = Registry.ToDisplay("shape", box.GetProps());
        Assert.Equal("textbox", shown["geometry"]);
        Assert.Equal(("2cm", "5cm", "10cm", "2cm"), (shown["x"], shown["y"], shown["w"], shown["h"]));
        Assert.Equal("Revenue grew 25%\nSecond line", shown["text"]);
        Assert.Equal(("Arial", "24pt", "FFFFFF"), (shown["font"], shown["size"], shown["color"]));
        Assert.Equal(2, box.Children.Count);
        Assert.Equal("true", box.Children[0].Children[1].GetProps()["bold"]);

        var rect = Mutations.Add(slide, "shape", Props(("geometry", "roundRect"), ("text", "Box"), ("fill", "4472C4"), ("line", "none")), null);
        var rectProps = rect.GetProps();
        Assert.Equal(("roundRect", "4472C4", "none"), (rectProps["geometry"], rectProps["fill"], rectProps["line"]));
        rect = Mutations.Set(rect, Props(("geometry", "ellipse"), ("fill", "none"), ("line", "FF0000"), ("w", "3cm")));
        Assert.Equal(("ellipse", "none", "FF0000", "1080000"), (rect.GetProps()["geometry"], rect.GetProps()["fill"], rect.GetProps()["line"], rect.GetProps()["w"]));
        Assert.Equal("/slide[1]/shape[2]", rect.Path);

        var title = Mutations.Set(slide, Props(("title", "Made on a blank slide")));
        Assert.Equal("Made on a blank slide", title.GetProps()["title"]);
        Assert.Equal(3, title.Children.Count);
        Assert.Equal("title", title.Children[2].GetProps()["placeholder"]);
    }

    [Fact]
    public void Slide_notes_visibility_and_transition_survive_save()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "Blank")), null);
        slide = Mutations.Set(slide, Props(("notes", "Discuss the figures"), ("hidden", "true"),
            ("transition", "fade"), ("duration", "700")));
        Assert.Equal("Discuss the figures", slide.GetProps()["notes"]);
        Assert.Equal("true", slide.GetProps()["hidden"]);
        Assert.Equal("fade", slide.GetProps()["transition"]);
        Assert.Equal("700", slide.GetProps()["duration"]);

        using var reopened = new PptxAdapter().Open(new MemoryStream(Save(doc)));
        var saved = reopened.Root.Children.Single();
        Assert.Equal("Discuss the figures", saved.GetProps()["notes"]);
        Assert.Equal("true", saved.GetProps()["hidden"]);
        Assert.Equal("fade", saved.GetProps()["transition"]);
        Assert.Equal("700", saved.GetProps()["duration"]);
        saved = Mutations.Set(saved, Props(("hidden", "false"), ("transition", "none")));
        Assert.DoesNotContain("hidden", saved.GetProps().Keys);
        Assert.DoesNotContain("transition", saved.GetProps().Keys);
    }

    [Fact]
    public void Paragraphs_and_runs_inside_shapes()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "Content"), ("title", "List")), null);
        var body = slide.Children[1];
        Assert.Equal("body", body.GetProps()["placeholder"]);
        var p1 = Mutations.Set(body.Children[0], Props(("text", "First"), ("list", "bullet")));
        var p2 = Mutations.Add(body, "paragraph", Props(("text", "Second"), ("list", "number"), ("level", "1"), ("align", "center")), null);
        Assert.Equal(("bullet", "0"), (p1.GetProps()["list"], p1.GetProps()["level"]));
        Assert.Equal(("number", "1", "center"), (p2.GetProps()["list"], p2.GetProps()["level"], p2.GetProps()["align"]));
        Assert.Equal("First\nSecond", Mutations.Refresh(body).Text);
        var run = Mutations.Add(p2, "run", Props(("text", " more"), ("bold", "true"), ("underline", "true"), ("color", "red"), ("size", "18"), ("link", "https://x.y")), null);
        var shown = Registry.ToDisplay("run", run.GetProps());
        Assert.Equal(("true", "true", "FF0000", "18pt", "https://x.y"), (shown["bold"], shown["underline"], shown["color"], shown["size"], shown["link"]));
        Assert.Equal("Second more", Mutations.Refresh(p2).Text);
        run = Mutations.Set(run, Props(("bold", "false"), ("link", ""), ("strike", "true"), ("font", "Georgia")));
        Assert.False(run.GetProps().ContainsKey("bold"));
        Assert.False(run.GetProps().ContainsKey("link"));
        Assert.Equal(("true", "Georgia"), (run.GetProps()["strike"], run.GetProps()["font"]));
        Mutations.Set(p1, Props(("list", "none")));
        Assert.False(Mutations.Refresh(p1).GetProps().ContainsKey("list"));
        p1.Remove();
        Assert.Single(Mutations.Refresh(body).Children);
    }

    [Fact]
    public void Tables_and_images_on_slides()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "Blank")), null);
        var table = Mutations.Add(slide, "table", Props(("data", "[[\"Name\",\"Score\"],[\"Ann\",\"90\"]]"), ("x", "1cm"), ("y", "2cm")), null);
        var props = Registry.ToDisplay("table", table.GetProps());
        Assert.Equal(("2", "2", "1cm", "2cm"), (props["rows"], props["cols"], props["x"], props["y"]));
        Assert.Equal("[[\"Name\",\"Score\"],[\"Ann\",\"90\"]]", props["data"]);
        table = Mutations.Set(table, Props(("rows", "3"), ("cols", "3")));
        Assert.Equal("[[\"Name\",\"Score\",\"\"],[\"Ann\",\"90\",\"\"],[\"\",\"\",\"\"]]", table.GetProps()["data"]);
        var row = Mutations.Set(PathResolver.Single(doc.Root, "/slide[1]/table[1]/row[3]"), Props(("data", "[\"Bob\",\"85\",\"x\"]")));
        Assert.Equal("[\"Bob\",\"85\",\"x\"]", row.GetProps()["data"]);
        var cell = Mutations.Set(PathResolver.Single(doc.Root, "//row[1]/cell[1]"), Props(("md", "**Name**"), ("fill", "D9E2F3"), ("align", "center")));
        Assert.Equal(("Name", "D9E2F3", "center"), (cell.Text, cell.GetProps()["fill"], cell.GetProps()["align"]));
        Assert.Equal("true", cell.Children[0].Children[0].GetProps()["bold"]);
        Mutations.Add(table, "row", Props(("data", "[\"Cy\",\"70\",\"y\"]")), 1);
        Assert.Equal("4", Mutations.Refresh(table).GetProps()["rows"]);
        Assert.Contains("Name\tScore\t\n", Views.Text(doc.Root));

        var png = Path.Combine(Path.GetTempPath(), $"writer-pptx-{Guid.NewGuid():N}.png");
        File.WriteAllBytes(png, FakePng(4, 2));
        try
        {
            var image = Mutations.Add(slide, "image", Props(("src", png), ("w", "8cm"), ("alt", "wide")), null);
            var shown = Registry.ToDisplay("image", image.GetProps());
            Assert.Equal(("8cm", "4cm", "wide"), (shown["w"], shown["h"], shown["alt"]));
            Assert.Contains("/media/", shown["src"]);
            Assert.NotNull(image.GetBinary());
            image = Mutations.Set(image, Props(("x", "3cm"), ("h", "2cm")));
            Assert.Equal(("3cm", "2cm"), (Registry.ToDisplay("image", image.GetProps())["x"], Registry.ToDisplay("image", image.GetProps())["h"]));
            Assert.Equal("/slide[1]/image[1]", image.Path);
        }
        finally
        {
            File.Delete(png);
        }
        Assert.Equal(new[] { "table", "image" }, Mutations.Refresh(slide).Children.Select(c => c.Kind));
    }

    [Fact]
    public void Raw_slide_and_shape_xml_round_trip()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "Blank")), null);
        Mutations.Add(slide, "shape", Props(("text", "hi")), null);
        var raw = PathResolver.Single(doc.Root, "/slide[1]/shape[1]").GetRaw();
        Assert.StartsWith("<p:sp", raw);
        PathResolver.Single(doc.Root, "/slide[1]/shape[1]").SetRaw(raw.Replace(">hi<", ">raw<"));
        Assert.Equal("raw", PathResolver.Single(doc.Root, "/slide[1]/shape[1]").Text);
        var slideRaw = slide.GetRaw();
        Assert.StartsWith("<p:sld", slideRaw);
        slide.SetRaw(slideRaw.Replace(">raw<", ">slide<"));
        Assert.Equal("slide", PathResolver.Single(doc.Root, "/slide[1]/shape[1]").Text);
    }

    [Fact]
    public void Sample_deck_reads_titles_positions_and_tables()
    {
        using var file = File.OpenRead(Path.Combine(FixtureDir("pptx"), "tables-merged.pptx"));
        using var doc = new PptxAdapter().Open(file);
        Assert.True(int.Parse(doc.Root.GetProps()["slides"]) >= 1);
        var tables = PathResolver.Query(doc.Root, "//table");
        Assert.NotEmpty(tables);
        Assert.True(int.Parse(tables[0].GetProps()["rows"]) >= 1);
        var shapes = PathResolver.Query(doc.Root, "//shape");
        Assert.Contains(shapes, s => s.GetProps().ContainsKey("x"));
        Assert.Contains("--- /slide[1] ---", Views.Text(doc.Root));
    }
}

public class PptxFidelityTests
{
    public static TheoryData<string> Files => new(TestDocs.Fixtures("pptx", "*.pptx").Select(f => Path.GetFileName(f)));

    [Theory]
    [MemberData(nameof(Files))]
    public void Open_read_everything_save_changes_nothing(string file)
    {
        var original = File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("pptx"), file));
        using var doc = new PptxAdapter().Open(new MemoryStream(original));
        _ = Views.Outline(doc.Root);
        _ = Views.Text(doc.Root);
        _ = NodeJson.Serialize(doc.Root, int.MaxValue);
        using var saved = new MemoryStream();
        doc.Save(saved);
        Assert.Null(PackageCompare.Diff(original, saved.ToArray()));
    }

    [Theory]
    [MemberData(nameof(Files))]
    public void Editing_one_shape_changes_only_its_slide(string file)
    {
        var original = File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("pptx"), file));
        using var doc = new PptxAdapter().Open(new MemoryStream(original));
        var shape = PathResolver.Query(doc.Root, "//shape").FirstOrDefault(s => s.Text!.Length > 0);
        if (shape is null) return;
        var slidePart = ((PptxDocument)doc).Presentation.SlideParts.First(p => ReferenceEquals(p, shape.Parent!.Anchor));
        var partName = slidePart.Uri.ToString().TrimStart('/');
        Mutations.Set(shape, new Dictionary<string, string> { ["text"] = "EDITED" });
        using var saved = new MemoryStream();
        doc.Save(saved);
        Assert.Null(PackageCompare.Diff(original, saved.ToArray(), partName));
    }
}
