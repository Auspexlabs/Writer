using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Pptx;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class PptxMediaCommentTests
{
    static Dictionary<string, string> P(params (string, string)[] p) => p.ToDictionary(x => x.Item1, x => x.Item2);
    [Fact]
    public void Motion_path_transition_direction_and_shared_master_edits_are_native()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", P(("layout", "blank")), null);
        Mutations.Add(doc.Root, "slide", P(("layout", "blank")), null);
        var shape = Mutations.Add(slide, "shape", P(("text", "Moving")), null);
        Mutations.Set(slide, P(("animations", "[{\"shape\":\"" + shape.GetProps()["id"] + "\",\"effect\":\"motion\",\"path\":\"M 0 0 C .1 -.2 .2 -.2 .3 0 E\",\"duration\":1700}]"), ("transitionDirection", "u"), ("transition", "push")));
        Mutations.Set(slide, P(("masterEdit", """{"source":"master","props":{"text":"Shared brand","x":"1cm","y":"1cm","size":"20pt"}}""")));
        using var stream = new MemoryStream(); doc.Save(stream);
        using var reopened = new PptxAdapter().Open(new MemoryStream(stream.ToArray()));
        var first = reopened.Root.Children[0];
        Assert.Equal("u", first.GetProps()["transitionDirection"]);
        Assert.Contains("motion", first.GetProps()["animations"]);
        Assert.Contains(".3 0 E", first.GetProps()["animations"]);
        Assert.All(reopened.Root.Children, s => Assert.Contains(s.Children, c => c.Kind == "decor" && c.GetProps().GetValueOrDefault("text") == "Shared brand"));
        using var objects = JsonDocument.Parse(first.GetProps()["masterObjects"]);
        var id = objects.RootElement.EnumerateArray().Single(o => o.TryGetProperty("text", out var t) && t.GetString() == "Shared brand").GetProperty("id").GetString();
        Mutations.Set(first, P(("masterEdit", "{\"source\":\"master\",\"id\":\"" + id + "\",\"props\":{\"text\":\"Updated brand\",\"x\":\"2cm\"}}")));
        Assert.All(reopened.Root.Children, s => Assert.Contains(s.Children, c => c.Kind == "decor" && c.GetProps().GetValueOrDefault("text") == "Updated brand"));
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((PptxDocument)reopened).Package).Select(e => e.Path?.XPath + ": " + e.Description));
    }
    [Fact]
    public void Embedded_media_and_comments_survive_save_copy_and_edit()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", P(("layout", "blank")), null);
        var bytes = new byte[] { 82, 73, 70, 70, 0, 0, 0, 0, 87, 65, 86, 69 };
        var image = Mutations.Add(slide, "image", P(("src", "data:image/png;base64," + Convert.ToBase64String(FakePng(4, 2))), ("media", "data:audio/wav;base64," + Convert.ToBase64String(bytes))), null);
        Assert.Equal("audio", image.GetProps()["mediaType"]); Assert.Equal(bytes, image.GetBinary()!.Value.Data);
        using var destination = new PptxAdapter().Create(); var target = Mutations.Add(destination.Root, "slide", P(("layout", "blank")), null);
        var imported = Mutations.Copy(image, target, null); Assert.Equal(bytes, imported.GetBinary()!.Value.Data);
        using var output = new MemoryStream(); destination.Save(output); using var opened = new PptxAdapter().Open(new MemoryStream(output.ToArray()));
        Assert.Equal(bytes, opened.Root.Children[0].Children.Single(c => c.Kind == "image").GetBinary()!.Value.Data);

        Mutations.Set(slide, P(("comments", """[{"author":"王五","text":"确认数据"}]""")));
        using var stream = new MemoryStream(); doc.Save(stream);
        using var reopened = new PptxAdapter().Open(new MemoryStream(stream.ToArray()));
        var old = reopened.Root.Children[0]; var copy = Mutations.Copy(old, reopened.Root, null);
        Assert.Equal(bytes, copy.Children.Single(c => c.Kind == "image").GetBinary()!.Value.Data);
        using var comments = JsonDocument.Parse(old.GetProps()["comments"]);
        var id = comments.RootElement[0].GetProperty("id").GetString();
        Mutations.Set(old, P(("comments", "[{\"id\":\"" + id + "\",\"author\":\"王五\",\"text\":\"已确认\"}]")));
        Assert.Contains("已确认", old.GetProps()["comments"]);
        Assert.Contains(id!, old.GetProps()["comments"]);
        Assert.Equal("[]", copy.GetProps()["comments"]);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((PptxDocument)reopened).Package).Select(e => e.Path?.XPath + ": " + e.Description));
        Mutations.Set(old, P(("comments", "[]"))); Assert.Equal("[]", old.GetProps()["comments"]);
    }
}
