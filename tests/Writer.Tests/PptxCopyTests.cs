using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Cli;
using Writer.Core;
using Writer.Formats.Pptx;
using static Writer.Tests.TestDocs;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Tests;

/// <summary>copy: a slide or slide element copied within its deck, the XML deep-cloned with what it refers to; add --raw: an
/// element put back from the XML 'get --raw' printed.</summary>
public class PptxCopyTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);
    static readonly string Png = "data:image/png;base64," + Convert.ToBase64String(FakePng(4, 2));

    /// <summary>Slide 1 (Title layout): the title and subtitle placeholders, a picture, a table with a filled bold cell, notes.
    /// Slide 2: the Title layout too, nothing on it.</summary>
    static PptxDocument Deck()
    {
        var doc = (PptxDocument)new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "Title"), ("title", "Q4 Results"), ("notes", "Say hi")), null);
        Mutations.Add(slide, "image", Props(("src", Png), ("x", "2cm"), ("y", "8cm"), ("w", "4cm"), ("alt", "logo")), null);
        Mutations.Add(slide, "table", Props(("data", "[[\"Name\",\"Score\"],[\"Ann\",\"90\"]]")), null);
        Mutations.Set(PathResolver.Single(doc.Root, "/slide[1]/table[1]/row[1]/cell[1]"), Props(("fill", "D9E2F3"), ("md", "**Name**")));
        var second = Mutations.Add(doc.Root, "slide", Props(("layout", "Title")), null);
        foreach (var placeholder in second.Children.Where(c => c.Kind != "decor").Reverse()) placeholder.Remove();
        return doc;
    }

    static PptxDocument Reopen(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        ms.Position = 0;
        return (PptxDocument)new PptxAdapter().Open(ms);
    }

    static SlidePart PartOf(Node slide) => (SlidePart)slide.Anchor;
    static List<uint> Ids(Node slide) => PartOf(slide).Slide!.Descendants<P.NonVisualDrawingProperties>().Select(p => p.Id!.Value).ToList();
    static string WithoutIds(string xml) => Regex.Replace(xml, "<p:cNvPr id=\"\\d+\"", "<p:cNvPr");
    /// <summary>How a shape shows: its own props and what it inherits, without its drawing id.</summary>
    static Dictionary<string, string> Look(Node shape)
    {
        var props = shape.GetProps().Where(p => p.Key != "id").ToDictionary();
        foreach (var (k, v) in shape.GetComputed(shape.GetProps()) ?? new Dictionary<string, string>()) props.TryAdd(k, v);
        return props;
    }

    [Fact]
    public void A_slide_copy_is_the_same_slide_with_its_own_notes()
    {
        using var doc = Deck();
        var slide = PathResolver.Single(doc.Root, "/slide[1]");
        var copy = Mutations.Copy(slide, doc.Root, 2);
        Assert.Equal("/slide[2]", copy.Path);
        Assert.Equal(3, doc.Root.Children.Count);
        Assert.Equal(slide.GetRaw(), copy.GetRaw());
        Assert.NotEqual(slide.GetProps()["id"], copy.GetProps()["id"]);
        var (a, b) = (PartOf(slide), PartOf(copy));
        Assert.Same(a.SlideLayoutPart, b.SlideLayoutPart);
        Assert.Equal(a.ImageParts, b.ImageParts); // the pictures' images are shared, not copied
        Assert.NotSame(a.NotesSlidePart, b.NotesSlidePart);
        Assert.Same(b, b.NotesSlidePart!.SlidePart);
        Assert.Equal("Say hi", copy.GetProps()["notes"]);

        using var back = Reopen(doc);
        var (one, two) = (back.Root.Children[0], back.Root.Children[1]);
        Assert.Equal(one.GetRaw(), two.GetRaw());
        Assert.Equal(new[] { "decor", "shape", "shape", "image", "table" }, two.Children.Select(c => c.Kind)); // the Title Slide layout's accent rule first
        Assert.Equal(one.Children[3].GetBinary()!.Value.Data, two.Children[3].GetBinary()!.Value.Data);
        Assert.Equal(("Say hi", "Q4 Results"), (two.GetProps()["notes"], two.GetProps()["title"]));
        Assert.Equal("Say hi", one.GetProps()["notes"]);
    }

    [Fact]
    public void A_shape_copy_keeps_what_it_inherits_and_gets_a_fresh_id()
    {
        using var doc = Deck();
        var title = PathResolver.Single(doc.Root, "/slide[1]/shape[1]");
        Assert.Contains("size", Look(title).Keys); // inherited from the layout, not written on the shape
        Assert.DoesNotContain("size", title.GetProps().Keys);
        var here = Mutations.Copy(title, title.Parent!, null);
        var there = Mutations.Copy(title, PathResolver.Single(doc.Root, "/slide[2]"), null);
        foreach (var copy in new[] { here, there })
        {
            Assert.Equal("title", copy.GetProps()["placeholder"]);
            Assert.Equal(Look(title), Look(copy));
        }
        Assert.Equal("/slide[1]/shape[3]", here.Path);
        Assert.NotEqual(title.GetProps()["id"], here.GetProps()["id"]);
        Assert.Equal(Ids(title.Parent!).Distinct().Count(), Ids(title.Parent!).Count);

        using var back = Reopen(doc);
        Assert.Equal(Look(PathResolver.Single(back.Root, "/slide[1]/shape[1]")), Look(PathResolver.Single(back.Root, "/slide[2]/shape[1]")));
    }

    [Fact]
    public void A_picture_copy_brings_its_image_and_the_original_kept_for_reset()
    {
        using var doc = Deck();
        var pic = PathResolver.Single(doc.Root, "/slide[1]/image[1]");
        var part = PartOf(pic.Parent!);
        var shown = ((P.Picture)pic.Anchor).BlipFill!.Blip!.Embed!.Value!;
        var kept = part.AddImagePart("image/png", "wrOrig" + shown); // what 抠图 leaves: the original under wrOrig + the shown id
        kept.FeedData(new MemoryStream(FakePng(8, 8)));

        var here = Mutations.Copy(pic, pic.Parent!, null);
        Assert.Equal(shown, ((P.Picture)here.Anchor).BlipFill!.Blip!.Embed!.Value); // on its own slide it shares the relationship
        var there = Mutations.Copy(pic, PathResolver.Single(doc.Root, "/slide[2]"), null);
        var target = PartOf(there.Parent!);
        var embed = ((P.Picture)there.Anchor).BlipFill!.Blip!.Embed!.Value!;
        Assert.Same(part.GetPartById(shown), target.GetPartById(embed));
        Assert.Same(kept, target.GetPartById("wrOrig" + embed));
        Assert.Equal(("logo", "/slide[2]/image[1]"), (there.GetProps()["alt"], there.Path));

        using var back = Reopen(doc);
        var copies = back.Root.Children.SelectMany(s => s.Children).Where(c => c.Kind == "image").ToList();
        Assert.Equal(3, copies.Count);
        Assert.All(copies, c => Assert.Equal(FakePng(4, 2), c.GetBinary()!.Value.Data));
    }

    [Fact]
    public void A_table_copy_keeps_its_style_and_cells()
    {
        using var doc = Deck();
        var table = PathResolver.Single(doc.Root, "/slide[1]/table[1]");
        var copy = Mutations.Copy(table, PathResolver.Single(doc.Root, "/slide[2]"), 1);
        Assert.Equal(WithoutIds(table.GetRaw()), WithoutIds(copy.GetRaw()));
        Assert.Equal(("D9E2F3", "true"), (copy.Children[0].Children[0].GetProps()["fill"], copy.Children[0].Children[0].Children[0].Children[0].GetProps()["bold"]));
        Assert.Equal(table.GetProps()["data"], copy.GetProps()["data"]);
    }

    [Fact]
    public void A_removed_picture_comes_back_from_its_raw_xml_with_its_image()
    {
        using var deck = Deck();
        using var saved = Reopen(deck);
        var pic = PathResolver.Single(saved.Root, "/slide[1]/image[1]");
        var (raw, id) = (pic.GetRaw(), pic.GetProps()["id"]);
        pic.Remove();
        using var removed = Reopen(saved);
        var slide = PathResolver.Single(removed.Root, "/slide[1]");
        var back = Mutations.AddRaw(slide, raw, null);
        Assert.Equal(("image", id, "/slide[1]/image[1]"), (back.Kind, back.GetProps()["id"], back.Path));
        Assert.Equal(FakePng(4, 2), back.GetBinary()!.Value.Data);
        var twice = Mutations.AddRaw(slide, raw, null); // its id is taken now: the second one gets a fresh one
        Assert.NotEqual(id, twice.GetProps()["id"]);
        Assert.Throws<WriterException>(() => Mutations.AddRaw(slide, "<p:nvSpPr/>", null));
    }

    public static TheoryData<string> Decks => new(Fixtures("pptx", "*.pptx").Select(f => Path.GetFileName(f)));

    /// <summary>Every slide of each fixture copied right after it (their animations, charts, notes, groups, pictures and links to
    /// other slides): each copy has the same XML and shows the same, every part the deck had is unchanged, and the schema finds no
    /// kind of error the deck did not have.</summary>
    [Theory]
    [MemberData(nameof(Decks))]
    public void Copying_every_slide_of_any_deck_changes_nothing_else(string file)
    {
        var original = File.ReadAllBytes(Path.Combine(FixtureDir("pptx"), file));
        using var doc = (PptxDocument)new PptxAdapter().Open(new MemoryStream(original));
        var count = doc.Root.Children.Count;
        for (var i = 0; i < count; i++)
        {
            var slide = doc.Root.Children[2 * i];
            var copy = Mutations.Copy(slide, doc.Root, 2 * i + 2);
            Assert.Equal(slide.GetRaw(), copy.GetRaw());
            var (a, b) = (PartOf(slide), PartOf(copy));
            foreach (var mine in a.Parts)
            {
                var theirs = b.Parts.Where(p => p.RelationshipId == mine.RelationshipId).Select(p => p.OpenXmlPart).FirstOrDefault();
                if (mine.OpenXmlPart is SlideCommentsPart or PowerPointCommentPart) Assert.Null(theirs); // comments stay with the original
                else if (mine.OpenXmlPart is ImagePart or SlideLayoutPart or SlidePart) Assert.Same(mine.OpenXmlPart, theirs);
                else Assert.True(theirs is not null && !ReferenceEquals(mine.OpenXmlPart, theirs) && theirs.GetType() == mine.OpenXmlPart.GetType(), mine.OpenXmlPart.GetType().Name);
            }
        }
        var ms = new MemoryStream();
        doc.Save(ms);
        var saved = ms.ToArray();
        var (before, after) = (PackageCompare.Parts(original), PackageCompare.Parts(saved));
        foreach (var (name, bytes) in before)
        {
            if (name is "ppt/presentation.xml" or "ppt/_rels/presentation.xml.rels" or "[Content_Types].xml") continue;
            Assert.True(after.TryGetValue(name, out var now), name + " is gone");
            Assert.True(PackageCompare.IsXml(name) ? PackageCompare.Canonical(bytes) == PackageCompare.Canonical(now) : bytes.AsSpan().SequenceEqual(now), name + " changed");
        }
        Assert.Empty(SchemaErrors(saved).Except(SchemaErrors(original)));
        using var back = new PptxAdapter().Open(new MemoryStream(saved));
        string Content(Node s) => string.Join("\n", s.GetProps().Where(p => p.Key != "id").Select(p => p.Key + "=" + p.Value)
            .Concat(s.Children.Select(c => NodeJson.Serialize(c, int.MaxValue).Replace(s.Path + "/", "/slide/"))));
        for (var i = 0; i < count; i++)
        {
            var (one, two) = (back.Root.Children[2 * i], back.Root.Children[2 * i + 1]);
            Assert.Equal(one.GetRaw(), two.GetRaw());
            Assert.Equal(Content(one), Content(two));
        }
    }

    /// <summary>The kinds of schema error in a deck: what and where in the part, not which part.</summary>
    static HashSet<string> SchemaErrors(byte[] deck)
    {
        using var package = PresentationDocument.Open(new MemoryStream(deck), false);
        return new OpenXmlValidator().Validate(package).Select(e => e.Path?.XPath + " " + e.Description).ToHashSet();
    }

    [Fact]
    public void The_cli_copies_and_puts_back()
    {
        var dir = Directory.CreateTempSubdirectory("writer-copy").FullName;
        try
        {
            var file = Path.Combine(dir, "deck.pptx");
            var docx = Path.Combine(dir, "doc.docx");
            string Run(params string[] argv)
            {
                var (stdout, stderr) = (new StringWriter(), new StringWriter());
                var code = Runner.Run(argv, stdout, stderr);
                Assert.True(code == 0, stderr.ToString());
                return stdout.ToString();
            }
            Run("create", file);
            Run("add", file, "/", "--type", "slide", "--prop", "layout=Blank");
            Run("add", file, "/slide[1]", "--type", "shape", "--prop", "text=Hi");
            var copied = JsonDocument.Parse(Run("copy", file, "/slide[1]/shape[1]", "--to", "/slide[1]")).RootElement;
            Assert.Equal("/slide[1]/shape[2]", copied.GetProperty("path").GetString());
            Assert.Equal("/slide[2]", JsonDocument.Parse(Run("copy", file, "/slide[1]", "--to", "/", "--after", "/slide[1]")).RootElement.GetProperty("path").GetString());
            var raw = Run("get", file, "/slide[2]/shape[2]", "--raw").Trim();
            Run("remove", file, "/slide[2]/shape[2]");
            Assert.Equal("Hi", JsonDocument.Parse(Run("add", file, "/slide[2]", "--raw", raw)).RootElement.GetProperty("props").GetProperty("text").GetString());

            Run("create", docx);
            Run("add", docx, "/body", "--type", "paragraph", "--prop", "text=x");
            var err = new StringWriter();
            Assert.NotEqual(0, Runner.Run(["copy", docx, "/body/paragraph[1]", "--to", "/body"], new StringWriter(), err));
            Assert.Contains("UNSUPPORTED_KIND", err.ToString());
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }
}
