using System.Globalization;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>Layout gallery geometry comes from this deck, including custom placeholders and its master defaults.</summary>
static class PptxLayouts
{
    public static string Read(PptxDocument doc)
    {
        var layouts = new JsonArray();
        foreach (var layout in doc.Layouts)
        {
            var look = PptxLook.For(layout, doc.Presentation);
            var shapes = new JsonArray();
            foreach (var shape in layout.SlideLayout?.CommonSlideData?.ShapeTree?.Elements<P.Shape>() ?? [])
            {
                var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                if (ph is null || ph.Type?.InnerText is "dt" or "ftr" or "sldNum") continue;
                var master = layout.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree?.Elements<P.Shape>().FirstOrDefault(s => {
                    var candidate = s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                    return candidate is not null && (candidate.Index?.Value == ph.Index?.Value && ph.Index is not null || (candidate.Type?.InnerText ?? "body") == (ph.Type?.InnerText ?? "body"));
                });
                var transform = shape.ShapeProperties?.Transform2D ?? master?.ShapeProperties?.Transform2D;
                var props = look.Shape(shape, new Dictionary<string,string>());
                PptxText.ReadBox(shape.TextBody, props);
                props["placeholder"] = ph.Type?.InnerText switch { "ctrTitle" => "title", "subTitle" => "subtitle", null => "body", var t => t };
                PptxShape.AddBox(props, transform?.Offset?.X?.Value, transform?.Offset?.Y?.Value, transform?.Extents?.Cx?.Value, transform?.Extents?.Cy?.Value);
                var json = new JsonObject(); foreach (var (key, value) in props) json[key] = value;
                shapes.Add((JsonNode)json);
            }
            var bg = layout.SlideLayout?.CommonSlideData?.Background ?? layout.SlideMasterPart?.SlideMaster?.CommonSlideData?.Background;
            layouts.Add((JsonNode)new JsonObject { ["name"] = PptxDocument.LayoutName(layout), ["placeholders"] = shapes, ["background"] = look.Background(bg) });
        }
        return layouts.ToJsonString();
    }
}
