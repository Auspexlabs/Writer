using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Pptx;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace Writer.Tests;

/// <summary>A slide's animations as p:timing (PowerPoint's presets, click groups and steps, the build list), the morph transition, and
/// sections (p14:sectionLst).</summary>
public class PptxAnimTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);
    static Document Open(byte[] bytes) => new PptxAdapter().Open(new MemoryStream(bytes));
    static byte[] Save(Document doc) { var ms = new MemoryStream(); doc.Save(ms); return ms.ToArray(); }
    static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("pptx"), name));
    static HashSet<string> Errors(Document doc) => new OpenXmlValidator(FileFormatVersions.Office2016).Validate(((PptxDocument)doc).Package)
        .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToHashSet();
    static List<Dictionary<string, string>> Effects(Node slide) => JsonDocument.Parse(slide.GetProps()["animations"]).RootElement.EnumerateArray()
        .Select(e => e.EnumerateObject().ToDictionary(p => p.Name, p => p.Value.ValueKind == JsonValueKind.String ? p.Value.GetString()! : p.Value.GetRawText())).ToList();
    static string Of(List<Dictionary<string, string>> fx, string key) => string.Join(" ", fx.Select(f => f.GetValueOrDefault(key, "-")));

    [Fact]
    public void Animations_are_written_the_way_PowerPoint_writes_them_and_go_with_their_shape()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        string Id(string text) => Mutations.Add(slide, "shape", Props(("text", text), ("x", "2cm"), ("y", "2cm"), ("w", "4cm"), ("h", "2cm")), null).GetProps()["id"];
        var (a, b) = (Id("A"), Id("B"));
        Mutations.Set(slide, Props(("animations", $"[{{\"shape\":\"{a}\",\"effect\":\"fade\"}},{{\"shape\":\"{b}\",\"effect\":\"fly\",\"start\":\"with\",\"delay\":200}},"
            + $"{{\"shape\":\"{a}\",\"effect\":\"spin\",\"start\":\"after\",\"duration\":1000}},{{\"shape\":\"{b}\",\"effect\":\"fadeOut\",\"start\":\"click\"}}]")));
        Assert.Empty(Errors(doc));
        using var back = Open(Save(doc));
        slide = back.Root.Children[0];
        var fx = Effects(slide);
        Assert.Equal(("fade fly spin fadeOut", "click with after click", "500 500 1000 500", "0 200 0 0", $"{a} {b} {a} {b}"),
            (Of(fx, "effect"), Of(fx, "start"), Of(fx, "duration"), Of(fx, "delay"), Of(fx, "shape")));
        var timing = ((SlidePart)slide.Anchor).Slide!.Timing!;
        var main = timing.Descendants<P.CommonTimeNode>().Single(c => c.NodeType?.InnerText == "mainSeq");
        var groups = main.ChildTimeNodeList!.Elements<P.ParallelTimeNode>().ToList();
        Assert.Equal(2, groups.Count); // a click starts each
        Assert.Equal(["0", "700"], groups[0].CommonTimeNode!.ChildTimeNodeList!.Elements<P.ParallelTimeNode>().Select(s => s.CommonTimeNode!.StartConditionList!.GetFirstChild<P.Condition>()!.Delay!.Value!));
        Assert.Equal(Enumerable.Range(1, timing.Descendants<P.CommonTimeNode>().Count()).Select(i => (uint)i), timing.Descendants<P.CommonTimeNode>().Select(c => c.Id!.Value)); // numbered as PowerPoint does
        var spin = timing.Descendants<P.CommonTimeNode>().Single(c => c.PresetClass?.InnerText == "emph");
        Assert.Equal(("8", "1"), (spin.PresetId!.Value.ToString(), spin.GroupId!.Value.ToString())); // A's second effect
        Assert.Equal([$"{a}:0", $"{b}:0", $"{a}:1", $"{b}:1"], timing.BuildList!.Elements<P.BuildParagraph>().Select(p => $"{p.ShapeId!.Value}:{p.GroupId!.Value}"));

        // removing a shape takes its effects and builds along; an empty list removes the timing
        slide.Children.Single(c => c.GetProps().GetValueOrDefault("id") == b).Remove();
        using var after = Open(Save(back));
        Assert.Equal(("fade spin", "click after"), (Of(Effects(after.Root.Children[0]), "effect"), Of(Effects(after.Root.Children[0]), "start")));
        Assert.Empty(Errors(after));
        Mutations.Set(after.Root.Children[0], Props(("animations", "[]")));
        Assert.Null(((SlidePart)after.Root.Children[0].Anchor).Slide!.Timing);
        Assert.Throws<WriterException>(() => Mutations.Set(after.Root.Children[0], Props(("animations", "[{\"shape\":\"999\",\"effect\":\"fade\"}]"))));
    }

    [Fact]
    public void Effects_the_engine_does_not_model_are_read_and_written_back_unchanged()
    {
        using var doc = Open(Fixture("animations.pptx"));
        var before = Errors(doc);
        var slide = doc.Root.Children[5]; // fade, fly (after), wheel (with), wedge (after) twice
        var fx = Effects(slide);
        Assert.Equal(("fade fly other other other", "click after with after after", "entrance entrance entrance"),
            (Of(fx, "effect"), Of(fx, "start"), string.Join(" ", fx.Where(f => f["effect"] == "other").Select(f => f["class"]))));
        var json = JsonDocument.Parse(slide.GetProps()["animations"]).RootElement.EnumerateArray().Select(e => e.GetRawText()).ToList();
        (json[1], json[2]) = (json[2], json[1]); // the wheel before the fly
        Mutations.Set(slide, Props(("animations", "[" + string.Join(",", json) + "]")));
        using var back = Open(Save(doc));
        var again = Effects(back.Root.Children[5]);
        Assert.Equal("fade other fly other other", Of(again, "effect"));
        Assert.Equal("click with after after after", Of(again, "start"));
        Assert.Contains("presetID=\"21\"", again[1]["xml"]);
        Assert.Empty(Errors(back).Except(before));
    }

    [Fact]
    public void Morph_is_written_for_PowerPoint_2019_with_a_fade_for_older_versions()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank"), ("transition", "morph"), ("duration", "1500")), null);
        Assert.Equal(("morph", "1500"), (slide.GetProps()["transition"], slide.GetProps()["duration"]));
        var ac = ((SlidePart)slide.Anchor).Slide!.GetFirstChild<AlternateContent>()!;
        Assert.Equal("p159", ac.GetFirstChild<AlternateContentChoice>()!.Requires!.Value);
        Assert.Equal("fade", ac.GetFirstChild<AlternateContentFallback>()!.GetFirstChild<P.Transition>()!.ChildElements.Single().LocalName);
        Assert.Empty(Errors(doc));
        using var back = Open(Save(doc));
        Assert.Equal("morph", back.Root.Children[0].GetProps()["transition"]);
    }

    [Fact]
    public void Sections_run_from_their_first_slide_and_take_in_the_slides_put_before_the_next()
    {
        using var doc = new PptxAdapter().Create();
        var slides = Enumerable.Range(0, 4).Select(_ => Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null)).ToList();
        Mutations.Set(slides[2], Props(("section", "Results")));
        static string Of(Document d) => string.Join("|", d.Root.Children.Select(s => s.GetProps().GetValueOrDefault("section", "")));
        Assert.Equal("Default Section||Results|", Of(doc));
        Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), 3); // before Results' first slide: in the section before
        using var back = Open(Save(doc));
        Assert.Equal("Default Section|||Results|", Of(back));
        Assert.Equal([3, 2], ((PptxDocument)back).Presentation.Presentation!.PresentationExtensionList!.Descendants<P14.Section>().Select(s => s.SectionSlideIdList!.Count()));
        Assert.Empty(Errors(back));
        Mutations.Set(back.Root.Children[3], Props(("section", "Summary")));
        Assert.Equal("Default Section|||Summary|", Of(back));
        Mutations.Set(back.Root.Children[0], Props(("section", ""))); // the first goes: its slides join the next
        Assert.Equal("Summary||||", Of(back));
        Mutations.Set(back.Root.Children[0], Props(("section", "")));
        Assert.Null(((PptxDocument)back).Presentation.Presentation!.PresentationExtensionList);
    }
}
