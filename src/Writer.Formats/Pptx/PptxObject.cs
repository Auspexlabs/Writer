using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

// Non-table graphic frames keep their original parts and markup. The shared Office
// renderer can draw cached charts and SmartArt; other frames still have a visible box.
sealed class PptxObject(PptxDocument doc, SlidePart slide, P.GraphicFrame frame) : Node
{
    public override string Kind => "object";
    public override object Anchor => frame;
    internal PptxDecor.Transform T { get; init; } = PptxDecor.Transform.Identity;
    static string? Attr(OpenXmlElement? e, string name) => e?.GetAttributes().FirstOrDefault(a => a.LocalName == name).Value;
    OpenXmlPart? Part(string? id) => id is not null && slide.TryGetPartById(id, out var part) ? part : null;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var data = frame.Graphic?.GraphicData;
        var uri = data?.Uri?.Value;
        var props = new Dictionary<string, string> { ["type"] = uri switch { OfficeGraphics.ChartUri or OfficeGraphics.ChartExUri => "chart", OfficeGraphics.DiagramUri => "smartart", _ => "drawing" } };
        var x = frame.Transform;
        PptxShape.AddBox(props, x?.Offset?.X?.Value, x?.Offset?.Y?.Value, x?.Extents?.Cx?.Value, x?.Extents?.Cy?.Value); T.Apply(props);
        var nv = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
        if (nv?.Id?.Value is { } id) props["id"] = id.ToString(CultureInfo.InvariantCulture);
        if (nv?.Name?.Value is { } name) props["name"] = name;
        if (nv?.Description?.Value is { } alt) props["alt"] = alt;
        var scheme = OfficeGraphics.Scheme(slide.SlideLayoutPart?.SlideMasterPart?.ThemePart);
        if (props["type"] == "chart" && Part(Attr(data?.ChildElements.FirstOrDefault(), "id")) is { } chart)
        {
            var json = uri == OfficeGraphics.ChartExUri ? OfficeGraphics.ChartExJson(chart, scheme) : OfficeGraphics.ChartJson(chart, scheme);
            if (json is not null) props["chart"] = json;
            if (OfficeGraphics.ChartTitle(json) is { } title) props["title"] = title;
        }
        if (props["type"] == "smartart")
        {
            var rels = data?.Descendants().FirstOrDefault(e => e.LocalName == "relIds");
            var dm = Part(Attr(rels, "dm"));
            var drawingId = dm is null ? null : OfficeGraphics.SmartArtDrawingId(dm);
            var drawing = drawingId is not null && dm!.TryGetPartById(drawingId, out var d) ? d : null;
            if (OfficeGraphics.SmartArtJson(drawing, dm, scheme) is { } json) props["smartart"] = json;
        }
        return props;
    }
    public override void SetProp(string name, string value)
    {
        if (name is not ("x" or "y" or "w" or "h")) return;
        var x = frame.Transform ??= new P.Transform(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 0, Cy = 0 });
        PptxOutline.SetBox(x.Offset!, x.Extents!, name, value, T);
    }
    public override string GetRaw() => frame.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(frame, raw, doc.Namespaces);
    public override void Remove() => frame.Remove();
    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent is not PptxSlide target) throw new WriterException(ErrorCode.Validation, "Graphic frames move between slides only", "Use --to /slide[n].");
        frame.Remove(); target.Insert(frame, index);
    }
    public override Node CopyTo(Node newParent, int? index) => newParent is PptxSlide target
        ? target.Place(PptxCopy.Element(frame, slide, target.Part), index) : base.CopyTo(newParent, index);
}
