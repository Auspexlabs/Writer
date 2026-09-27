using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Pptx;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Tests;

public class PptxEditorTests
{
    static Dictionary<string, string> Props(params (string Key, string Value)[] p) => p.ToDictionary(x => x.Key, x => x.Value);
    static Document Reopen(Document doc) { var stream = new MemoryStream(); doc.Save(stream); stream.Position = 0; return new PptxAdapter().Open(stream); }
    static void Valid(Document doc) => Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2016).Validate(((PptxDocument)doc).Package));

    [Fact]
    public void Text_box_formatting_and_dimensions_survive_reopening()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        var shape = Mutations.Add(slide, "shape", Props(("text", "One\nTwo")), null);
        Mutations.Set(shape, Props(("bold", "true"), ("italic", "true"), ("underline", "true"), ("align", "right"), ("verticalAlign", "bottom"), ("size", "24pt")));
        Mutations.Set(doc.Root, Props(("width", "25.4cm"), ("height", "19.05cm")));
        Valid(doc);
        using var back = Reopen(doc);
        var props = PathResolver.Single(back.Root, "/slide[1]/shape[1]").GetProps();
        Assert.Equal(("true", "true", "true", "right", "bottom", "24"), (props["bold"], props["italic"], props["underline"], props["align"], props["verticalAlign"], props["size"]));
        Assert.Equal("9144000", back.Root.GetProps()["width"]);
        var body = ((P.Shape)shape.Anchor).TextBody!;
        Assert.All(body.Elements<A.Paragraph>(), p => Assert.True(p.GetFirstChild<A.EndParagraphRunProperties>()!.Bold!.Value));
    }

    [Fact]
    public void Page_number_fields_follow_slide_order_and_keep_formatting()
    {
        using var doc = new PptxAdapter().Create();
        var first = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        var second = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        var number = Mutations.Add(second, "shape", Props(("field", "slideNumber"), ("size", "18pt"), ("bold", "true")), null);
        Assert.Equal("2", number.GetProps()["text"]);
        second.MoveTo(doc.Root, 1);
        Assert.Equal("1", number.GetProps()["text"]);
        var date = Mutations.Add(first, "shape", Props(("field", "date")), null);
        Assert.Equal("date", date.GetProps()["field"]);
        Valid(doc);
        using var back = Reopen(doc);
        var p = PathResolver.Single(back.Root, "/slide[1]/shape[1]").GetProps();
        Assert.Equal(("slideNumber", "1", "18", "true"), (p["field"], p["text"], p["size"], p["bold"]));
    }

    [Theory]
    [InlineData("charts-pie.pptx", "chart")]
    public void Graphic_frames_are_visible_copyable_and_removable_without_losing_parts(string file, string type)
    {
        using var doc = new PptxAdapter().Open(File.OpenRead(Path.Combine(TestDocs.FixtureDir("pptx"), file)));
        var frame = PathResolver.Query(doc.Root, "//object").First(n => n.GetProps()["type"] == type);
        Assert.True(frame.GetProps().ContainsKey(type));
        var original = frame.GetRaw();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        var copy = Mutations.Copy(frame, slide, null);
        Assert.Equal(frame.GetProps()[type], copy.GetProps()[type]);
        frame.Remove();
        using var back = Reopen(doc);
        Assert.Contains(PathResolver.Query(back.Root, "//object"), n => n.GetProps()[type] == copy.GetProps()[type]);
        Assert.Contains("graphicFrame", original);
        using var destination = new PptxAdapter().Create();
        var target = Mutations.Add(destination.Root, "slide", Props(("layout", "blank")), null);
        var imported = Mutations.Copy(copy, target, null);
        Assert.Equal(copy.GetProps()[type], imported.GetProps()[type]);
        using var reopened = Reopen(destination);
        Assert.Equal(copy.GetProps()[type], PathResolver.Single(reopened.Root, "/slide[1]/object[1]").GetProps()[type]);
    }

    [Fact]
    public void SmartArt_without_a_cached_drawing_exposes_its_node_text_and_keeps_the_part()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "blank")), null);
        var part = (SlidePart)slide.Anchor;
        var data = part.AddNewPart<DiagramDataPart>();
        using (var writer = new StreamWriter(data.GetStream(FileMode.Create))) writer.Write("""
            <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="1" type="node"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Plan</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/></dgm:dataModel>
            """);
        slide.AddRaw($"""
            <p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="3" name="Plan diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="{part.GetIdOfPart(data)}"/></a:graphicData></a:graphic></p:graphicFrame>
            """, null);
        using var back = Reopen(doc);
        var frame = PathResolver.Single(back.Root, "/slide[1]/object[1]");
        Assert.Equal("smartart", frame.GetProps()["type"]);
        Assert.Contains("Plan", frame.GetProps()["smartart"]);
        var other = Mutations.Add(back.Root, "slide", Props(("layout", "blank")), null);
        Assert.Equal(frame.GetProps()["smartart"], frame.CopyTo(other, null).GetProps()["smartart"]);
    }
}
