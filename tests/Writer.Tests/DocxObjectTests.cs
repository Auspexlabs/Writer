using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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

    [Fact]
    public void Saved_deletion_can_restore_original_preview_and_embedded_payload()
    {
        var ole = new W.EmbeddedObject("""
            <w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <v:shape id="_x0000_i1025" style="width:72pt;height:36pt"><v:imagedata r:id="rIdPreview"/></v:shape><o:OLEObject ProgID="Equation.DSMT4" r:id="rIdEquation" ShapeID="_x0000_i1025"/>
            </w:object>
            """);
        var png = FakePng(96, 48); byte[] payload = [1, 8, 7, 9, 2];
        using var doc = OpenDocx(Docx([new W.Paragraph(new W.Run(new W.Text("before")), new W.Run(ole), new W.Run(new W.Text("after")))], main =>
        {
            using var image = new MemoryStream(png); main.AddImagePart(ImagePartType.Png, "rIdPreview").FeedData(image);
            using var data = new MemoryStream(payload); main.AddEmbeddedObjectPart("application/vnd.ms-office.oleObject", "rIdEquation").FeedData(data);
        }));
        var obj = PathResolver.Single(doc.Root, "//object");
        var xml = obj.GetProps()["xml"];
        obj.Remove(); using var deleted = new MemoryStream(); doc.Save(deleted);
        using var restored = OpenDocx(deleted.ToArray());
        obj = Mutations.AddRaw(PathResolver.Single(restored.Root, "/body/paragraph[1]"), $"<w:r>{xml}</w:r>", null);
        Mutations.Set(obj, new Dictionary<string, string> { ["at"] = "6" });
        Assert.Equal("6", obj.GetProps()["at"]);
        Assert.Equal(png, obj.GetBinary()!.Value.Data);
        using var saved = new MemoryStream(); restored.Save(saved);
        using var package = WordprocessingDocument.Open(new MemoryStream(saved.ToArray()), false);
        using var bytes = new MemoryStream(); package.MainDocumentPart!.GetPartById("rIdEquation").GetStream().CopyTo(bytes);
        Assert.Equal(payload, bytes.ToArray());
    }

    [Fact]
    public void Copies_renumber_drawing_ids_and_block_inline_moves_return_the_new_path()
    {
        using var doc = OpenDocx(Docx(new W.Paragraph(new W.Run(Ole())), P("text")));
        var first = PathResolver.Single(doc.Root, "/body/object[1]");
        var copy = Mutations.AddRaw(PathResolver.Single(doc.Root, "/body"), first.GetRaw(), null);
        Assert.NotEqual(first.GetRaw(), copy.GetRaw());
        Assert.Contains("_x0000_i1026", copy.GetRaw());
        var moved = Mutations.Move(copy, PathResolver.Single(doc.Root, "/body/paragraph[1]"), null);
        Assert.Equal("/body/paragraph[1]/object[1]", moved.Path);
        moved = Mutations.Move(moved, PathResolver.Single(doc.Root, "/body"), 1);
        Assert.Equal("/body/object[1]", moved.Path);
        Assert.Equal(2, PathResolver.Query(doc.Root, "//object").Count());
    }
}
