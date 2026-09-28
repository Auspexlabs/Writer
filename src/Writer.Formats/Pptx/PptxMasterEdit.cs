using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

static class PptxMasterEdit
{
    static P.ShapeTree? Tree(SlidePart slide, string source) => source == "layout" ? slide.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree : slide.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree;
    public static string Read(PptxDocument doc, SlidePart slide) => NodeJson.Compact(w =>
    {
        w.WriteStartArray();
        foreach (var source in new[] { "master", "layout" }) foreach (var shape in Tree(slide, source)?.Elements<P.Shape>() ?? [])
        {
            w.WriteStartObject(); w.WriteString("source", source);
            foreach (var (key, value) in Registry.ToDisplay("pptx", "shape", new PptxShape(doc, slide, shape).GetProps())) w.WriteString(key, value);
            w.WriteEndObject();
        }
        w.WriteEndArray();
    });
    public static void Write(PptxDocument doc, SlidePart slide, string json)
    {
        using var data = JsonDocument.Parse(json); var o = data.RootElement;
        var source = o.TryGetProperty("source", out var from) ? from.GetString() ?? "master" : "master";
        if (source is not ("master" or "layout")) throw new WriterException(ErrorCode.Validation, "Unknown master source", "Use master or layout.");
        var tree = Tree(slide, source) ?? throw new WriterException(ErrorCode.Validation, "Slide has no master or layout", "Apply a layout first.");
        var id = o.TryGetProperty("id", out var idValue) ? idValue.ToString() : "";
        var shape = tree.Elements<P.Shape>().FirstOrDefault(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value.ToString() == id);
        if (o.TryGetProperty("remove", out var remove) && remove.GetBoolean()) { shape?.Remove(); return; }
        var props = new Dictionary<string, string>();
        if (o.TryGetProperty("props", out var edits)) foreach (var p in edits.EnumerateObject())
        {
            if (p.Name is not ("text" or "font" or "size" or "color" or "bold" or "italic" or "fill" or "line" or "x" or "y" or "w" or "h" or "align" or "valign" or "geometry")) throw new WriterException(ErrorCode.Validation, "Unsupported master property: " + p.Name, "Edit text, geometry or formatting.");
            props[p.Name] = p.Value.ValueKind == JsonValueKind.String ? p.Value.GetString()! : p.Value.GetRawText();
        }
        var canonical = Registry.Normalize("pptx", "shape", props);
        if (shape is null)
        {
            if (id.Length > 0) throw new WriterException(ErrorCode.Validation, "Master shape not found", "Reload the document.");
            var created = PptxShape.New(doc, slide, canonical); shape = (P.Shape)created.Element;
            shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = tree.Descendants<P.NonVisualDrawingProperties>().Select(p => p.Id?.Value ?? 0).DefaultIfEmpty().Max() + 1;
            tree.Append(shape);
        }
        var node = new PptxShape(doc, slide, shape); foreach (var (key, value) in canonical) node.SetProp(key, value);
    }
}
