using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>A shape, picture, connector or background picture a slide shows from its layout or master. Read-only: it projects
/// the same node classes as slide content over the layout or master element, so the props match the shape and image kinds.</summary>
sealed class PptxDecor(PptxDocument doc, SlidePart slide, OpenXmlPart part, OpenXmlElement element, string source, PptxDecor.Transform transform) : Node
{
    public override string Kind => "decor";
    public override object Anchor => element;

    /// <summary>What the slide shows underneath its own shapes, in drawing order: background picture, master shapes, layout shapes.
    /// Placeholders are skipped (the slide's own shapes stand in for them); showMasterSp="0" on the slide or layout hides what it says.</summary>
    internal static IEnumerable<Node> Of(PptxDocument doc, SlidePart slide)
    {
        if (Background(slide) is { } bg && bg.Background.BackgroundProperties?.GetFirstChild<A.BlipFill>() is not null)
            yield return new PptxDecor(doc, slide, bg.Part, bg.Background, bg.Source, Transform.Identity);
        if (slide.Slide?.ShowMasterShapes?.Value == false) yield break;
        var layout = slide.SlideLayoutPart;
        if (layout?.SlideMasterPart is { } master && layout.SlideLayout?.ShowMasterShapes?.Value != false)
            foreach (var (e, t) in Flatten(master.SlideMaster?.CommonSlideData?.ShapeTree, Transform.Identity)) yield return new PptxDecor(doc, slide, master, e, "master", t);
        if (layout is not null)
            foreach (var (e, t) in Flatten(layout.SlideLayout?.CommonSlideData?.ShapeTree, Transform.Identity)) yield return new PptxDecor(doc, slide, layout, e, "layout", t);
    }

    /// <summary>The effective background: the first p:bg on the slide, its layout, then its master.</summary>
    internal static (OpenXmlPart Part, P.Background Background, string Source)? Background(SlidePart slide)
    {
        if (slide.Slide?.CommonSlideData?.Background is { } own) return (slide, own, "slide");
        var layout = slide.SlideLayoutPart;
        if (layout?.SlideLayout?.CommonSlideData?.Background is { } fromLayout) return (layout, fromLayout, "layout");
        if (layout?.SlideMasterPart is { } master && master.SlideMaster?.CommonSlideData?.Background is { } fromMaster) return (master, fromMaster, "master");
        return null;
    }

    /// <summary>Non-placeholder shapes, pictures and connectors of a shape tree; groups contribute their members with the group's transform.</summary>
    static IEnumerable<(OpenXmlElement Element, Transform Transform)> Flatten(OpenXmlElement? tree, Transform t)
    {
        if (tree is null) yield break;
        foreach (var child in tree.ChildElements)
        {
            switch (child)
            {
                case P.Shape sp when sp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is null: yield return (sp, t); break;
                case P.Picture pic when pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is null: yield return (pic, t); break;
                case P.ConnectionShape cxn: yield return (cxn, t); break;
                case P.GroupShape group:
                    foreach (var inner in Flatten(group, t.Then(group.GroupShapeProperties?.TransformGroup))) yield return inner;
                    break;
            }
        }
    }

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        Dictionary<string, string> props;
        var type = "shape";
        switch (element)
        {
            case P.Shape sp:
                props = new(new PptxShape(doc, slide, sp).GetProps());
                break;
            case P.Picture pic:
                props = new(new PptxImage(doc, part, pic).GetProps());
                type = "image";
                break;
            case P.ConnectionShape cxn:
                props = new();
                var xfrm = cxn.ShapeProperties?.Transform2D;
                PptxShape.AddBox(props, xfrm?.Offset?.X?.Value, xfrm?.Offset?.Y?.Value, xfrm?.Extents?.Cx?.Value, xfrm?.Extents?.Cy?.Value);
                if (cxn.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText is { } preset) props["geometry"] = preset;
                if (cxn.ShapeProperties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value is { } line) props["line"] = line.ToUpperInvariant();
                if (cxn.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value is { Length: > 0 } name) props["name"] = name;
                break;
            default:
                var (width, height) = doc.SlideSize;
                props = new() { ["background"] = "true" };
                PptxShape.AddBox(props, 0, 0, width, height);
                if (Embed() is { } embed && part.TryGetPartById(embed, out var image)) props["src"] = image.Uri.ToString();
                type = "image";
                break;
        }
        props["source"] = source;
        props["type"] = type;
        transform.Apply(props);
        return props;
    }

    public override IReadOnlyDictionary<string, string>? GetComputed(IReadOnlyDictionary<string, string> props) =>
        element is P.Shape or P.ConnectionShape ? PptxLook.For(slide, doc.Presentation).Shape(element, props) : null;

    string? Embed() => (element as P.Background)?.BackgroundProperties?.GetFirstChild<A.BlipFill>()?.Blip?.Embed?.Value;

    public override (string ContentType, byte[] Data)? GetBinary() => element switch
    {
        P.Picture pic => new PptxImage(doc, part, pic).GetBinary(),
        P.Background => PptxImage.Bytes(part, Embed()),
        _ => null,
    };

    public override string GetRaw() => element.OuterXml;

    static WriterException ReadOnly() => new(ErrorCode.FormatReadonly, "Decor inherited from the layout or master is read-only (decor)",
        "Edit the layout in PowerPoint, or add a shape on the slide instead.");
    public override void SetProp(string name, string value) => throw ReadOnly();
    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index) => throw ReadOnly();
    public override void Remove() => throw ReadOnly();
    public override void MoveTo(Node newParent, int? index) => throw ReadOnly();
    public override void SetRaw(string raw) => throw ReadOnly();

    /// <summary>Maps a grouped shape's coordinates onto the slide, x' = Ax + Sx·x; nested groups compose.</summary>
    internal readonly record struct Transform(double Ax, double Ay, double Sx, double Sy)
    {
        public static readonly Transform Identity = new(0, 0, 1, 1);

        /// <summary>This transform after the group's own (child space → group space).</summary>
        public Transform Then(A.TransformGroup? group)
        {
            if (group?.Offset is not { } off || group.Extents is not { } ext) return this;
            var sx = Scale(ext.Cx?.Value, group.ChildExtents?.Cx?.Value);
            var sy = Scale(ext.Cy?.Value, group.ChildExtents?.Cy?.Value);
            var ax = (off.X?.Value ?? 0) - (group.ChildOffset?.X?.Value ?? 0) * sx;
            var ay = (off.Y?.Value ?? 0) - (group.ChildOffset?.Y?.Value ?? 0) * sy;
            return new(Ax + Sx * ax, Ay + Sy * ay, Sx * sx, Sy * sy);
        }

        static double Scale(long? extent, long? childExtent) => extent is long e && childExtent is long c && c > 0 ? e / (double)c : 1;

        public void Apply(Dictionary<string, string> props)
        {
            if (this == Identity) return;
            Map(props, "x", Ax, Sx);
            Map(props, "y", Ay, Sy);
            Map(props, "w", 0, Sx);
            Map(props, "h", 0, Sy);
        }

        static void Map(Dictionary<string, string> props, string name, double a, double s)
        {
            if (props.TryGetValue(name, out var v) && long.TryParse(v, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var emu))
                props[name] = ((long)Math.Round(a + s * emu)).ToString(CultureInfo.InvariantCulture);
        }
    }
}
