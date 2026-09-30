using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Pptx;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class PptxMediaCommentTests
{
    [Fact] public void Caption_tracks_are_native_WebVTT_and_survive_copy_edit_delete_and_reopen()
    {
        using var doc=new PptxAdapter().Create();var slide=Mutations.Add(doc.Root,"slide",P(("layout","blank")),null);
        var image=Mutations.Add(slide,"image",P(("src","data:image/png;base64,"+Convert.ToBase64String(FakePng(4,2))),("media","data:audio/wav;base64,UklGRgAAAABXQVZF")),null);
        var spec=new System.Text.Json.Nodes.JsonObject{["display"]="slide",["tracks"]=new System.Text.Json.Nodes.JsonArray((System.Text.Json.Nodes.JsonNode)new System.Text.Json.Nodes.JsonObject{["label"]="中文",["lang"]="zh-CN",["text"]="WEBVTT\n\n00:00:00.000 --> 00:00:02.000\n字幕 & 测试\n"})};
        image.SetProp("captions",spec.ToJsonString());var expected=image.GetProps()["captions"];Assert.Contains("tracksInfo",image.GetRaw());
        using var stream=new MemoryStream();doc.Save(stream);using var opened=new PptxAdapter().Open(new MemoryStream(stream.ToArray()));var again=opened.Root.Children[0].Children.Single(n=>n.Kind=="image");Assert.Equal(expected,again.GetProps()["captions"]);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(((PptxDocument)opened).Package).Select(e=>e.Path?.XPath+": "+e.Description));
        using var targetDoc=new PptxAdapter().Create();var target=Mutations.Add(targetDoc.Root,"slide",P(("layout","blank")),null);var copy=Mutations.Copy(again,target,null);Assert.Equal(expected,copy.GetProps()["captions"]);
        spec["tracks"]![0]!["text"]="Invalid file";Assert.Throws<WriterException>(()=>again.SetProp("captions",spec.ToJsonString()));Assert.Equal(expected,again.GetProps()["captions"]);
        again.SetProp("captions","{\"display\":\"media\",\"tracks\":[]}");Assert.DoesNotContain("tracksInfo",again.GetRaw());Assert.Equal(expected,copy.GetProps()["captions"]);
        Assert.DoesNotContain(((PptxDocument)opened).Presentation.SlideParts.First().Parts,p=>p.OpenXmlPart.ContentType=="text/vtt");
    }
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
    [Fact]
    public void Media_trim_fades_playback_and_bookmarks_survive_copy_animation_changes_and_deletion()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", P(("layout", "blank")), null);
        var image = Mutations.Add(slide, "image", P(("src", "data:image/png;base64," + Convert.ToBase64String(FakePng(4, 2))), ("media", "data:audio/wav;base64,UklGRgAAAABXQVZF")), null);
        const string options = """{"trimStart":0.1,"trimEnd":0.2,"fadeIn":0.1,"fadeOut":0.1,"volume":65,"autoplay":true,"loop":true,"slideCount":3,"hideStopped":true,"bookmarks":[{"name":"Start","time":0.3}]}""";
        Mutations.Set(image, P(("playback", options)));
        var expected = image.GetProps()["playback"];
        var shape = Mutations.Add(slide, "shape", P(("text", "Animation stays independent")), null);
        Mutations.Set(slide, P(("animations", "[{\"shape\":\"" + shape.GetProps()["id"] + "\",\"effect\":\"fade\"}]")));
        Assert.Equal(expected, image.GetProps()["playback"]);
        using var target = new PptxAdapter().Create(); var targetSlide = Mutations.Add(target.Root, "slide", P(("layout", "blank")), null);
        var copied = Mutations.Copy(image, targetSlide, null); Assert.Equal(expected, copied.GetProps()["playback"]);
        using var stream = new MemoryStream(); doc.Save(stream); using var opened = new PptxAdapter().Open(new MemoryStream(stream.ToArray()));
        var loaded = opened.Root.Children[0].Children.Single(c => c.Kind == "image"); Assert.Equal(expected, loaded.GetProps()["playback"]);
        using var package = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(new MemoryStream(stream.ToArray()), false);
        var sp = package.PresentationPart!.SlideParts.First(); var trim = sp.Slide!.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.MediaTrim>().Single();
        Assert.Equal("100", trim.Start!.InnerText); Assert.Equal("200", trim.End!.InnerText);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((PptxDocument)opened).Package).Select(e => e.Path?.XPath + ": " + e.Description));
        Mutations.Set(loaded, P(("playback", """{"autoplay":false,"volume":0,"loop":false}""")));
        using var cleared = JsonDocument.Parse(loaded.GetProps()["playback"]); Assert.False(cleared.RootElement.GetProperty("autoplay").GetBoolean()); Assert.Equal(0, cleared.RootElement.GetProperty("volume").GetDouble());
        loaded.Remove(); using var deleted = new MemoryStream(); opened.Save(deleted);
        Assert.DoesNotContain(((PptxDocument)opened).Package.PresentationPart!.SlideParts.First().Slide!.Descendants<DocumentFormat.OpenXml.Presentation.CommonMediaNode>(), _ => true);
        var errors = string.Join("\n", new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((PptxDocument)target).Package).Select(e => e.Path?.XPath + ": " + e.Description)); Assert.True(errors.Length == 0, errors);
    }

}
