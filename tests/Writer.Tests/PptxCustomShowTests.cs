using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Pptx;

namespace Writer.Tests;

public class PptxCustomShowTests
{
    static Dictionary<string, string> Props(string key, string value) => new() { [key] = value };

    [Fact]
    public void Custom_shows_preserve_order_duplicates_and_remove_deleted_slides()
    {
        using var doc = new PptxAdapter().Create();
        var slides = Enumerable.Range(0, 3).Select(_ => Mutations.Add(doc.Root, "slide", Props("layout", "blank"), null)).ToArray();
        var ids = slides.Select(s => s.GetProps()["id"]).ToArray();
        var json = "[{\"name\":\"Sales\",\"slides\":[\"" + ids[2] + "\",\"" + ids[0] + "\",\"" + ids[2] + "\"]}]";
        Mutations.Set(doc.Root, Props("customShows", json));
        Assert.Equal(json, doc.Root.GetProps()["customShows"]);
        using var stream = new MemoryStream(); doc.Save(stream);
        using var reopened = new PptxAdapter().Open(new MemoryStream(stream.ToArray()));
        Assert.Equal(json, reopened.Root.GetProps()["customShows"]);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((PptxDocument)reopened).Package));
        reopened.Root.Children[2].Remove();
        using var pruned = new MemoryStream(); reopened.Save(pruned);
        using var parsed = JsonDocument.Parse(reopened.Root.GetProps()["customShows"]);
        Assert.Equal(ids[0], parsed.RootElement[0].GetProperty("slides")[0].GetString());
        Assert.Equal(1, parsed.RootElement[0].GetProperty("slides").GetArrayLength());
        reopened.Root.Children[0].Remove();
        using var empty = new MemoryStream(); reopened.Save(empty);
        Assert.Equal("[]", reopened.Root.GetProps()["customShows"]);
    }

    [Fact]
    public void Invalid_custom_show_updates_do_not_replace_valid_shows()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props("layout", "blank"), null);
        var show = "{\"name\":\"Demo\",\"slides\":[\"" + slide.GetProps()["id"] + "\"]}";
        Mutations.Set(doc.Root, Props("customShows", "[" + show + "]"));
        var before = doc.Root.GetProps()["customShows"];
        Assert.Throws<WriterException>(() => Mutations.Set(doc.Root, Props("customShows", "[" + show + "," + show + "]")));
        Assert.Equal(before, doc.Root.GetProps()["customShows"]);
        Assert.Throws<WriterException>(() => Mutations.Set(doc.Root, Props("customShows", "[{\"name\":\"Missing\",\"slides\":[\"9999\"]}]")));
        Assert.Equal(before, doc.Root.GetProps()["customShows"]);
    }
}
