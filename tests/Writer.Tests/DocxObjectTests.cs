using DocumentFormat.OpenXml;
using Writer.Core;
using static Writer.Tests.TestDocs;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Tests;

public class DocxObjectTests
{
    static W.EmbeddedObject Ole() => new("""
        <w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
          <v:shape id="_x0000_i1025" style="width:72pt;height:36pt"/><o:OLEObject ProgID="Equation.DSMT4"/>
        </w:object>
        """);

    [Fact]
    public void Standalone_and_inline_objects_survive_text_edits_and_delete_independently()
    {
        var p = new W.Paragraph(new W.Run(new W.Text("before")), new W.Run(Ole()), new W.Run(new W.Text("after")));
        using var doc = OpenDocx(Docx(new W.Paragraph(new W.Run(Ole())), p));
        var objects = PathResolver.Query(doc.Root, "//object").ToArray();
        Assert.Equal(2, objects.Length);
        Assert.Equal("/body/object[1]", objects[0].Path);
        Assert.Equal("/body/paragraph[1]/object[1]", objects[1].Path);
        Assert.Equal("6", objects[1].GetProps()["at"]);
        Assert.Equal("914400", objects[0].GetProps()["width"]);
        var original = objects[1].GetRaw();
        Mutations.Set(PathResolver.Single(doc.Root, "/body/paragraph[1]"), new Dictionary<string, string> { ["html"] = "before<b>edited</b>after" });
        Assert.Equal(original, PathResolver.Single(doc.Root, "/body/paragraph[1]/object[1]").GetRaw());
        using var stream = new MemoryStream(); doc.Save(stream);
        using var reopened = OpenDocx(stream.ToArray());
        PathResolver.Single(reopened.Root, "/body/object[1]").Remove();
        var inline = Assert.Single(PathResolver.Query(reopened.Root, "//object"));
        inline.Remove();
        Assert.Empty(PathResolver.Query(reopened.Root, "//object"));
        Assert.Equal("beforeeditedafter", PathResolver.Single(reopened.Root, "/body/paragraph[1]").Text);
    }

    [Fact]
    public void Object_offset_can_move_without_replacing_its_payload()
    {
        using var doc = OpenDocx(Docx(new W.Paragraph(new W.Run(new W.Text("abc")), new W.Run(Ole()), new W.Run(new W.Text("def")))));
        var obj = PathResolver.Single(doc.Root, "//object");
        var raw = obj.GetRaw();
        Mutations.Set(obj, new Dictionary<string, string> { ["at"] = "1" });
        Assert.Equal("1", obj.GetProps()["at"]);
        Assert.Equal(raw, obj.GetRaw());
        Assert.Equal("abcdef", PathResolver.Single(doc.Root, "/body/paragraph[1]").Text);
    }
}
