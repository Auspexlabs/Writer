using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Html;
using Writer.Formats.Pptx;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Tests;

/// <summary>Layout and master decoration, speaker notes, hidden slides and transitions. decor.pptx is the Mars deck plus a decorated
/// master (bar, line, group, footer placeholder, solid background), a layout logo picture, a layout with showMasterSp="0" and one
/// with a picture background, a slide with showMasterSp="0", a spd="slow" fade, and notes/hidden/push+duration baked into slide 2.</summary>
public class PptxDecorTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);
    static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("pptx"), name));
    static Document Open(byte[] bytes) => new PptxAdapter().Open(new MemoryStream(bytes));
    static byte[] Save(Document doc) { var ms = new MemoryStream(); doc.Save(ms); return ms.ToArray(); }
    static Node Slide(Document doc, int n) => doc.Root.Children[n - 1];
    static SlidePart Part(Node slide) => (SlidePart)slide.Anchor;
    static string PartName(OpenXmlPart part) => part.Uri.ToString().TrimStart('/');
    static string Rels(string partName) => Path.GetDirectoryName(partName)!.Replace('\\', '/') + "/_rels/" + Path.GetFileName(partName) + ".rels";
    static PresentationPart Presentation(Document doc) => ((PptxDocument)doc).Presentation;

    /// <summary>Schema errors as strings, so a test can assert that an edit adds none to whatever the fixture already had.</summary>
    static HashSet<string> Errors(Document doc) =>
        new OpenXmlValidator(FileFormatVersions.Office2016).Validate(((PptxDocument)doc).Package)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToHashSet();

    static void AssertNoNewErrors(HashSet<string> before, Document doc) => Assert.Empty(Errors(doc).Except(before));

    [Fact]
    public void Decor_comes_first_and_leaves_shape_ordinals_alone()
    {
        using var mars = Open(Fixture("Mars-Settlement-Guide.pptx"));
        using var doc = Open(Fixture("decor.pptx"));
        for (var i = 1; i <= 5; i++)
        {
            var children = Slide(doc, i).Children;
            var decor = children.TakeWhile(c => c.Kind == "decor").ToList();
            Assert.Equal(Enumerable.Range(1, decor.Count).Select(n => $"/slide[{i}]/decor[{n}]"), decor.Select(d => d.Path));
            Assert.DoesNotContain(children.Skip(decor.Count), c => c.Kind == "decor");
            Assert.Equal(Slide(mars, i).Children.Select(c => (c.Path, c.Text)), children.Skip(decor.Count).Select(c => (c.Path, c.Text)));
        }
        var props = Slide(doc, 1).Children.Where(c => c.Kind == "decor").Select(c => c.GetProps()).ToList();
        Assert.Equal(["master", "master", "master", "layout"], props.Select(p => p["source"]));
        Assert.Equal(["shape", "shape", "shape", "image"], props.Select(p => p["type"]));
        Assert.Equal(["Master Bar", "Master Line", "Group Dot"], props.Take(3).Select(p => p["name"]));
        Assert.Equal(("MASTER BAR", "1F2A44", "0", "6400800", "12192000"), (props[0]["text"], props[0]["fill"], props[0]["x"], props[0]["y"], props[0]["w"]));
        Assert.Equal(("line", "FF5722"), (props[1]["geometry"], props[1]["line"]));
        // the dot is at (100,100) size 300 in a group whose child space 1000×500 maps onto 2286000×1143000 EMU at (9144000, 457200)
        Assert.Equal(("9372600", "685800", "685800", "685800"), (props[2]["x"], props[2]["y"], props[2]["w"], props[2]["h"]));
        Assert.Equal(("/ppt/media/decor-logo.png", "Logo", "457200"), (props[3]["src"], props[3]["alt"], props[3]["x"]));
        Assert.DoesNotContain(PathResolver.Query(doc.Root, "//decor"), d => d.GetProps().GetValueOrDefault("name") == "Footer Placeholder");
    }

    [Fact]
    public void Decor_respects_showMasterSp_and_the_background_falls_back_to_layout_then_master()
    {
        using var doc = Open(Fixture("decor.pptx"));
        Assert.DoesNotContain(Slide(doc, 3).Children, c => c.Kind == "decor");           // p:sld showMasterSp="0" hides layout and master shapes
        var layoutOnly = Slide(doc, 4).Children.Where(c => c.Kind == "decor").ToList();  // the layout's showMasterSp="0" hides the master's
        Assert.Equal(["LAYOUT TAG"], layoutOnly.Select(c => c.Text));
        Assert.Equal("layout", layoutOnly[0].GetProps()["source"]);
        Assert.Equal("101020", Slide(doc, 4).GetProps()["background"]);                  // neither the slide nor its layout has one: the master's
        Assert.Equal("080A1F", Slide(doc, 1).GetProps()["background"]);                  // the slide's own wins
        Assert.False(Slide(doc, 1).GetProps().ContainsKey("backgroundImage"));

        var picture = Slide(doc, 5);
        Assert.Equal("true", picture.GetProps()["backgroundImage"]);
        Assert.False(picture.GetProps().ContainsKey("background"));
        var bg = picture.Children[0].GetProps();
        Assert.Equal(("decor", "image", "layout", "true"), (picture.Children[0].Kind, bg["type"], bg["source"], bg["background"]));
        Assert.Equal(("0", "0", "12192000", "6858000"), (bg["x"], bg["y"], bg["w"], bg["h"]));
        Assert.Equal(["Master Bar", "Master Line", "Group Dot"], picture.Children.Skip(1).TakeWhile(c => c.Kind == "decor").Select(c => c.GetProps()["name"]));
    }

    [Fact]
    public void Decor_pictures_serve_their_bytes()
    {
        using var doc = Open(Fixture("decor.pptx"));
        foreach (var path in new[] { "/slide[1]/decor[4]", "/slide[5]/decor[1]" })
        {
            var binary = PathResolver.Single(doc.Root, path).GetBinary();
            Assert.NotNull(binary);
            Assert.Equal("image/png", binary.Value.ContentType);
            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, binary.Value.Data.Take(4));
        }
        Assert.Null(PathResolver.Single(doc.Root, "/slide[1]/decor[1]").GetBinary());
    }

    [Fact]
    public void Decor_is_read_only_and_reading_it_changes_nothing()
    {
        var original = Fixture("decor.pptx");
        using var doc = Open(original);
        var decor = PathResolver.Single(doc.Root, "/slide[1]/decor[1]");
        Assert.Equal(ErrorCode.Validation, Assert.Throws<WriterException>(() => Mutations.Set(decor, Props(("text", "x")))).Code);
        Assert.Equal(ErrorCode.FormatReadonly, Assert.Throws<WriterException>(() => decor.Remove()).Code);
        Assert.Equal(ErrorCode.FormatReadonly, Assert.Throws<WriterException>(() => decor.MoveTo(Slide(doc, 2), null)).Code);
        Assert.Equal(ErrorCode.FormatReadonly, Assert.Throws<WriterException>(() => decor.SetRaw("<p:sp/>")).Code);
        Assert.Throws<WriterException>(() => Mutations.Move(decor, Slide(doc, 2), null));   // the registry lists no parent: not addable, not movable
        Assert.Throws<WriterException>(() => Mutations.Add(Slide(doc, 1), "decor", Props(), null));
        Assert.Equal(4, Slide(doc, 1).Children.Count(c => c.Kind == "decor"));
        Assert.Null(PackageCompare.Diff(original, Save(doc)));
    }

    [Fact]
    public void Adding_a_shape_at_index_one_puts_it_before_the_slides_own_shapes()
    {
        using var doc = Open(Fixture("decor.pptx"));
        var added = Mutations.Add(Slide(doc, 1), "shape", Props(("text", "FIRST")), 1);
        Assert.Equal("/slide[1]/shape[1]", added.Path);
        var children = Slide(doc, 1).Children;
        Assert.Equal(4, children.TakeWhile(c => c.Kind == "decor").Count());
        Assert.Equal("FIRST", children[4].Text);
    }

    [Fact]
    public void Views_draw_decor_first_list_it_and_keep_it_out_of_text()
    {
        using var doc = Open(Fixture("decor.pptx"));
        var html = HtmlWriter.Render(doc);
        var first = html.IndexOf("<div class=\"shape\"", StringComparison.Ordinal);
        var second = html.IndexOf("<div class=\"shape\"", first + 1, StringComparison.Ordinal);
        Assert.Contains("MASTER BAR", html[first..second]);
        Assert.Contains("data:image/png;base64,", html);
        var outline = Views.Outline(doc.Root);
        Assert.Contains("/slide[1]/decor[1]  ", outline);
        Assert.Contains("source=master", outline);
        Assert.Contains("notes=Line one Line two", outline);
        Assert.DoesNotContain("MASTER BAR", Views.Text(doc.Root));
        Assert.Contains("\"kind\": \"decor\"", NodeJson.Serialize(doc.Root, int.MaxValue));
    }

    [Fact]
    public void Notes_hidden_and_transition_read_from_the_file()
    {
        using var doc = Open(Fixture("decor.pptx"));
        var p = Slide(doc, 2).GetProps();
        Assert.Equal(("Line one\nLine two", "true", "push", "1200"), (p["notes"], p["hidden"], p["transition"], p["duration"]));
        Assert.Equal("fade", Slide(doc, 1).GetProps()["transition"]);                     // spd="slow" only: no duration
        Assert.False(Slide(doc, 1).GetProps().ContainsKey("duration"));
        Assert.Equal("morph", Slide(doc, 3).GetProps()["transition"]);                    // p159:morph in mc:AlternateContent
        Assert.False(Slide(doc, 1).GetProps().ContainsKey("hidden"));
        Assert.False(Slide(doc, 1).GetProps().ContainsKey("notes"));
        Assert.DoesNotContain(Errors(doc), e => e.Contains("notes", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Notes_create_the_notes_master_once_the_way_powerpoint_does()
    {
        using var doc = new PptxAdapter().Create();
        Mutations.Add(doc.Root, "slide", Props(("layout", "Blank")), null);
        Mutations.Add(doc.Root, "slide", Props(("layout", "Blank")), null);
        var original = Save(doc);
        var before = Errors(doc);
        Assert.Null(Presentation(doc).NotesMasterPart);

        var slide = Mutations.Set(Slide(doc, 1), Props(("notes", "First line\nSecond line")));
        Assert.Equal("First line\nSecond line", slide.GetProps()["notes"]);
        var notes = Part(slide).NotesSlidePart!;
        var master = Presentation(doc).NotesMasterPart!;
        Assert.NotNull(master.ThemePart);
        Assert.Same(master, notes.NotesMasterPart);
        Assert.Same(Part(slide), notes.SlidePart);
        Assert.Equal(2, PptxDocument.NotesBody(notes)!.TextBody!.Elements<A.Paragraph>().Count());
        Assert.IsType<P.NotesMasterIdList>(Presentation(doc).Presentation!.SlideMasterIdList!.NextSibling());
        AssertNoNewErrors(before, doc);
        string[] touched = [PartName(notes), Rels(PartName(notes)), PartName(master), Rels(PartName(master)), PartName(master.ThemePart!),
            "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", Rels(PartName(Part(slide))), "[Content_Types].xml"];
        Assert.Null(PackageCompare.Diff(original, Save(doc), touched));

        Mutations.Set(Slide(doc, 2), Props(("notes", "Second slide")));
        Assert.Single(Presentation(doc).GetPartsOfType<NotesMasterPart>());
        using var reopened = Open(Save(doc));
        Assert.Equal(["First line\nSecond line", "Second slide"], reopened.Root.Children.Select(s => s.GetProps()["notes"]));
    }

    [Fact]
    public void Notes_replace_the_placeholder_text_and_keep_the_part_when_cleared()
    {
        var original = Fixture("decor.pptx");
        using var doc = Open(original);
        var before = Errors(doc);
        var slide = Mutations.Set(Slide(doc, 2), Props(("notes", "Rewritten")));
        var notes = Part(slide).NotesSlidePart!;
        Assert.Equal("Rewritten", slide.GetProps()["notes"]);
        Assert.Equal(2, notes.NotesSlide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Count()); // the slide image placeholder stays
        Assert.Null(PackageCompare.Diff(original, Save(doc), PartName(notes)));
        slide = Mutations.Set(slide, Props(("notes", "")));
        Assert.Equal("", slide.GetProps()["notes"]);
        Assert.NotNull(Part(slide).NotesSlidePart);
        AssertNoNewErrors(before, doc);
        Mutations.Set(Slide(doc, 1), Props(("notes", "")));                              // nothing to clear: no part is created
        Assert.Null(Part(Slide(doc, 1)).NotesSlidePart);
        Assert.False(Slide(doc, 1).GetProps().ContainsKey("notes"));
    }

    [Fact]
    public void Hidden_round_trips_and_changes_only_the_slide()
    {
        var original = Fixture("decor.pptx");
        using var doc = Open(original);
        var before = Errors(doc);
        var slide = Mutations.Set(Slide(doc, 1), Props(("hidden", "true")));
        Assert.Equal("true", slide.GetProps()["hidden"]);
        Assert.False(Part(slide).Slide!.Show!.Value);
        AssertNoNewErrors(before, doc);
        var saved = Save(doc);
        Assert.Null(PackageCompare.Diff(original, saved, PartName(Part(slide))));
        using var reopened = Open(saved);
        Assert.Equal("true", Slide(reopened, 1).GetProps()["hidden"]);
        var shown = Mutations.Set(Slide(reopened, 1), Props(("hidden", "false")));
        Assert.False(shown.GetProps().ContainsKey("hidden"));
        Assert.Null(Part(shown).Slide!.Show);
        Assert.False(Mutations.Set(Slide(reopened, 2), Props(("hidden", "no"))).GetProps().ContainsKey("hidden"));
    }

    [Theory]
    [InlineData("fade", "")]
    [InlineData("push", "dir=l")]
    [InlineData("wipe", "dir=r")]
    [InlineData("split", "orient=horz dir=out")]
    [InlineData("cover", "dir=l")]
    [InlineData("cut", "")]
    [InlineData("dissolve", "")]
    [InlineData("zoom", "")]
    [InlineData("random", "")]
    public void Every_transition_writes_powerpoints_element_and_survives_save(string name, string attributes)
    {
        var original = Fixture("decor.pptx");
        using var doc = Open(original);
        var before = Errors(doc);
        var slide = Mutations.Set(Slide(doc, 3), Props(("transition", name)));           // replaces the morph
        Assert.Equal(name, slide.GetProps()["transition"]);
        var type = Part(slide).Slide!.Transition!.ChildElements.Single();
        Assert.Equal((PptxTemplate.PNs, name), (type.NamespaceUri, type.LocalName));
        Assert.Equal(attributes, string.Join(" ", type.GetAttributes().Select(a => $"{a.LocalName}={a.Value}")));
        AssertNoNewErrors(before, doc);
        var saved = Save(doc);
        Assert.Null(PackageCompare.Diff(original, saved, PartName(Part(slide))));
        using var reopened = Open(saved);
        Assert.Equal(name, Slide(reopened, 3).GetProps()["transition"]);
        var none = Mutations.Set(Slide(reopened, 3), Props(("transition", "none")));
        Assert.False(none.GetProps().ContainsKey("transition"));
        Assert.Null(Part(none).Slide!.Transition);
        Assert.Empty(Part(none).Slide!.Elements<AlternateContent>());
    }

    [Fact]
    public void Duration_uses_the_p14_alternate_content_and_keeps_speed_attributes()
    {
        var original = Fixture("decor.pptx");
        using var doc = Open(original);
        var before = Errors(doc);
        var slide = Mutations.Set(Slide(doc, 1), Props(("transition", "wipe"), ("duration", "500")));  // slide 1 has spd="slow" advTm="3000"
        Assert.Equal(("wipe", "500"), (slide.GetProps()["transition"], slide.GetProps()["duration"]));
        var sld = Part(slide).Slide!;
        Assert.Null(sld.Transition);
        var alternate = Assert.Single(sld.Elements<AlternateContent>());
        var choice = Assert.IsType<AlternateContentChoice>(alternate.FirstChild);
        Assert.Equal("p14", choice.Requires?.Value);
        var chosen = Assert.IsType<P.Transition>(choice.FirstChild);
        var fallback = Assert.IsType<P.Transition>(Assert.IsType<AlternateContentFallback>(alternate.LastChild).FirstChild);
        Assert.Equal(("500", null), (chosen.Duration?.Value, fallback.Duration?.Value));
        foreach (var t in new[] { chosen, fallback })
        {
            Assert.Equal("slow", t.Speed?.InnerText);
            Assert.Equal("3000", t.AdvanceAfterTime?.InnerText);
            Assert.Equal("wipe", t.FirstChild!.LocalName);
        }
        Assert.True(alternate.NextSibling() is null or P.Timing or P.ExtensionList);
        AssertNoNewErrors(before, doc);
        var saved = Save(doc);
        Assert.Null(PackageCompare.Diff(original, saved, PartName(Part(slide))));
        using var reopened = Open(saved);
        var p = Slide(reopened, 1).GetProps();
        Assert.Equal(("wipe", "500"), (p["transition"], p["duration"]));
        var changed = Mutations.Set(Slide(reopened, 1), Props(("duration", "900")));     // duration alone keeps the effect
        Assert.Equal(("wipe", "900"), (changed.GetProps()["transition"], changed.GetProps()["duration"]));
        Assert.Equal(ErrorCode.Validation, Assert.Throws<WriterException>(() => Mutations.Set(changed, Props(("transition", "other")))).Code);
    }
}
