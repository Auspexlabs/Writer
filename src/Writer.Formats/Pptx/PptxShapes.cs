using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>The look shapes and connectors share: outline width, dash and arrowheads (a:ln), the outer shadow (a:effectLst),
/// a two-stop gradient (a:gradFill) and the rotation an xfrm carries.</summary>
static class PptxOutline
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    public static void Read(P.ShapeProperties? spPr, Dictionary<string, string> props)
    {
        if (spPr is null) return;
        if (spPr.GetFirstChild<A.Outline>() is { } ln)
        {
            if (ln.Width?.Value is { } w) props["lineWidth"] = (w / 12700.0).ToString("0.##", Inv);
            if (ln.GetFirstChild<A.PresetDash>()?.Val?.InnerText is { } dash) props["dash"] = dash;
            if (ln.GetFirstChild<A.HeadEnd>()?.Type?.InnerText is { } head) props["head"] = head;
            if (ln.GetFirstChild<A.TailEnd>()?.Type?.InnerText is { } tail) props["tail"] = tail;
        }
        if (spPr.GetFirstChild<A.EffectList>()?.GetFirstChild<A.OuterShadow>() is not null) props["shadow"] = "true";
        if (spPr.GetFirstChild<A.GradientFill>() is { } grad && GradientOf(grad) is { } g) props["gradient"] = g;
        if (spPr.Transform2D?.Rotation?.Value is { } rot && rot != 0) props["rotation"] = (rot / 60000).ToString(Inv);
        if (spPr.Transform2D?.HorizontalFlip?.Value == true) props["flipH"] = "true";
        if (spPr.Transform2D?.VerticalFlip?.Value == true) props["flipV"] = "true";
    }

    /// <summary>True when the property is one of this class's; it is then written.</summary>
    public static bool Set(P.ShapeProperties spPr, string name, string value)
    {
        switch (name)
        {
            case "lineWidth": Line(spPr).Width = (int)Math.Round(double.Parse(value, Inv) * 12700); return true;
            case "dash":
                var ln = Line(spPr);
                ln.RemoveAllChildren<A.PresetDash>();
                if (value != "solid") PptxText.InsertBeforeAny(ln, new A.PresetDash { Val = new A.PresetLineDashValues(value) }, e => e is A.Round or A.Bevel or A.Miter or A.HeadEnd or A.TailEnd or A.ExtensionList);
                return true;
            case "head":
                Line(spPr).RemoveAllChildren<A.HeadEnd>();
                if (value != "none") PptxText.InsertBeforeAny(Line(spPr), new A.HeadEnd { Type = new A.LineEndValues(value) }, e => e is A.TailEnd or A.ExtensionList);
                return true;
            case "tail":
                Line(spPr).RemoveAllChildren<A.TailEnd>();
                if (value != "none") PptxText.InsertBeforeAny(Line(spPr), new A.TailEnd { Type = new A.LineEndValues(value) }, e => e is A.ExtensionList);
                return true;
            case "shadow":
                spPr.RemoveAllChildren<A.EffectList>();
                if (value == "true") PptxText.InsertBeforeAny(spPr, Shadow(), e => e is A.EffectDag or A.Scene3DType or A.Shape3DType or A.ExtensionList);
                return true;
            case "gradient":
                foreach (var f in spPr.ChildElements.Where(e => e is A.SolidFill or A.NoFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill).ToList()) f.Remove();
                if (value.Length == 0 || value == "none") return true;
                PptxText.InsertBeforeAny(spPr, Gradient(value, "gradient"), e => e is A.Outline or A.EffectList or A.EffectDag or A.Scene3DType or A.Shape3DType or A.ExtensionList);
                return true;
            case "rotation":
                Transform(spPr).Rotation = ((int.Parse(value, Inv) % 360) + 360) % 360 * 60000;
                return true;
            case "flipH": Transform(spPr).HorizontalFlip = value == "true" ? true : null; return true;
            case "flipV": Transform(spPr).VerticalFlip = value == "true" ? true : null; return true;
        }
        return false;
    }

    /// <summary>"A,B,angle": a two-stop linear gradient (angle in degrees, 90 = top to bottom).</summary>
    public static A.GradientFill Gradient(string value, string prop)
    {
        var parts = value.Split(',', StringSplitOptions.TrimEntries);
        if (parts.Length < 2) throw new WriterException(ErrorCode.Validation, $"{prop}: '{value}' is not two colours", $"Give the start and end colour and an angle, e.g. {prop}=4472C4,ED7D31,90.");
        var angle = parts.Length > 2 && int.TryParse(parts[2], NumberStyles.AllowLeadingSign, Inv, out var deg) ? ((deg % 360) + 360) % 360 : 90;
        return new A.GradientFill(
            new A.GradientStopList(
                new A.GradientStop(new A.RgbColorModelHex { Val = Units.ParseColor(parts[0]) }) { Position = 0 },
                new A.GradientStop(new A.RgbColorModelHex { Val = Units.ParseColor(parts[1]) }) { Position = 100000 }),
            new A.LinearGradientFill { Angle = angle * 60000, Scaled = false }) { RotateWithShape = true };
    }

    /// <summary>The "A,B,angle" a gradFill reads as, or null when it has fewer than two plain colours.</summary>
    public static string? GradientOf(A.GradientFill grad)
    {
        var stops = grad.GradientStopList?.Elements<A.GradientStop>().Select(s => s.RgbColorModelHex?.Val?.Value).ToList() ?? [];
        return stops.Count >= 2 && stops[0] is not null && stops[^1] is not null
            ? $"{stops[0]!.ToUpperInvariant()},{stops[^1]!.ToUpperInvariant()},{(grad.GetFirstChild<A.LinearGradientFill>()?.Angle?.Value ?? 5400000) / 60000}" : null;
    }

    /// <summary>Office's outer shadow, offset downwards.</summary>
    public static A.EffectList Shadow() => new(new A.OuterShadow(new A.PresetColor(new A.Alpha { Val = 40000 }) { Val = A.PresetColorValues.Black })
        { BlurRadius = 50800L, Distance = 38100L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Top, RotateWithShape = false });

    /// <summary>The spPr's a:ln, made when missing, in its place before the effects.</summary>
    public static A.Outline Line(P.ShapeProperties spPr)
    {
        if (spPr.GetFirstChild<A.Outline>() is { } line) return line;
        var ln = new A.Outline();
        PptxText.InsertBeforeAny(spPr, ln, e => e is A.EffectList or A.EffectDag or A.Scene3DType or A.Shape3DType or A.ExtensionList);
        return ln;
    }

    static A.Transform2D Transform(P.ShapeProperties spPr) =>
        spPr.Transform2D ?? (A.Transform2D)spPr.PrependChild(new A.Transform2D(new A.Offset { X = 914400L, Y = 914400L }, new A.Extents { Cx = 4572000L, Cy = 914400L }));

    /// <summary>x, y, w or h onto an xfrm, in slide units (the transform maps a grouped element's child space to the slide).</summary>
    public static void SetBox(A.Offset off, A.Extents ext, string name, string value, PptxDecor.Transform t)
    {
        var emu = t.Unmap(name, long.Parse(value, Inv));
        switch (name)
        {
            case "x": off.X = emu; break;
            case "y": off.Y = emu; break;
            case "w": ext.Cx = emu; break;
            default: ext.Cy = emu; break;
        }
    }

    public static P.ShapeStyle Style(A.SchemeColorValues fill, uint fillIdx, A.SchemeColorValues font) => new(
        new A.LineReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = fillIdx == 0 ? 1U : 2U },
        new A.FillReference(new A.SchemeColor { Val = fill }) { Index = fillIdx },
        new A.EffectReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
        new A.FontReference(new A.SchemeColor { Val = font }) { Index = A.FontCollectionIndexValues.Minor });
}

/// <summary>A p:cxnSp: a line or connector, straight or bent, with arrowheads, that may stick to shapes (stCxn / endCxn name the
/// shape's drawing id and connection site: 0 top, 1 left, 2 bottom, 3 right on a rectangle).</summary>
sealed class PptxConnector(PptxDocument doc, SlidePart slide, P.ConnectionShape cxn) : Node
{
    public override string Kind => "connector";
    public override object Anchor => cxn;
    internal PptxDecor.Transform T { get; init; } = PptxDecor.Transform.Identity;

    P.NonVisualDrawingProperties? Nv => cxn.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties;
    P.NonVisualConnectorShapeDrawingProperties Links => cxn.NonVisualConnectionShapeProperties!.NonVisualConnectorShapeDrawingProperties ??= new P.NonVisualConnectorShapeDrawingProperties();

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string>();
        var spPr = cxn.ShapeProperties;
        if (spPr?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText is { } preset) props["geometry"] = preset;
        var xfrm = spPr?.Transform2D;
        PptxShape.AddBox(props, xfrm?.Offset?.X?.Value, xfrm?.Offset?.Y?.Value, xfrm?.Extents?.Cx?.Value, xfrm?.Extents?.Cy?.Value);
        T.Apply(props);
        if (spPr?.GetFirstChild<A.Outline>() is { } line)
        {
            if (line.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value is { } color) props["line"] = color.ToUpperInvariant();
            else if (line.GetFirstChild<A.NoFill>() is not null) props["line"] = "none";
        }
        PptxOutline.Read(spPr, props);
        if (Links.StartConnection is { Id.Value: var sid, Index.Value: var sidx }) props["start"] = $"{sid},{sidx}";
        if (Links.EndConnection is { Id.Value: var eid, Index.Value: var eidx }) props["end"] = $"{eid},{eidx}";
        if (Nv?.Name?.Value is { Length: > 0 } name) props["name"] = name;
        if (Nv?.Id?.Value is { } id) props["id"] = id.ToString(CultureInfo.InvariantCulture);
        return props;
    }

    public override IReadOnlyDictionary<string, string>? GetComputed(IReadOnlyDictionary<string, string> props) =>
        PptxLook.For(slide, doc.Presentation).Shape(cxn, props);

    public override void SetProp(string name, string value)
    {
        var spPr = cxn.ShapeProperties ??= new P.ShapeProperties();
        if (PptxOutline.Set(spPr, name, value)) return;
        switch (name)
        {
            case "x" or "y" or "w" or "h":
                var xfrm = spPr.Transform2D ?? (A.Transform2D)spPr.PrependChild(new A.Transform2D(new A.Offset { X = 914400L, Y = 914400L }, new A.Extents { Cx = 3657600L, Cy = 0L }));
                PptxOutline.SetBox(xfrm.Offset!, xfrm.Extents!, name, value, T);
                break;
            case "geometry":
                spPr.RemoveAllChildren<A.PresetGeometry>();
                var geometry = new A.PresetGeometry(new A.AdjustValueList()) { Preset = new A.ShapeTypeValues(value) };
                if (spPr.Transform2D is { } after) spPr.InsertAfter(geometry, after);
                else spPr.PrependChild(geometry);
                break;
            case "line":
                var ln = PptxOutline.Line(spPr);
                PptxText.SetColor(ln, value);
                if (value == "none") ln.PrependChild(new A.NoFill());
                break;
            case "start" or "end":
                var links = Links;
                if (name == "start") links.StartConnection = null; else links.EndConnection = null;
                if (value.Length == 0 || value == "none") break;
                var parts = value.Split(',', StringSplitOptions.TrimEntries);
                if (parts.Length != 2 || !uint.TryParse(parts[0], out var sid) || !uint.TryParse(parts[1], out var sidx))
                    throw new WriterException(ErrorCode.Validation, $"{name}: '{value}' is not a shape id and site", "Give the shape's id and its connection site, e.g. start=4,3 (0 top, 1 left, 2 bottom, 3 right).");
                if (name == "start") links.StartConnection = new A.StartConnection { Id = sid, Index = sidx };
                else links.EndConnection = new A.EndConnection { Id = sid, Index = sidx };
                break;
        }
    }

    public static (OpenXmlElement Element, string[] Consumed) New(PptxDocument doc, SlidePart slide, IReadOnlyDictionary<string, string> props)
    {
        var geometry = props.GetValueOrDefault("geometry") ?? "straightConnector1";
        var id = PptxDocument.NextShapeId(slide);
        var element = new P.ConnectionShape(
            new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = "Connector " + id },
                new P.NonVisualConnectorShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = PptxShape.Emu(props, "x", 914400L), Y = PptxShape.Emu(props, "y", 914400L) },
                    new A.Extents { Cx = PptxShape.Emu(props, "w", 3657600L), Cy = PptxShape.Emu(props, "h", 0L) }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = new A.ShapeTypeValues(geometry) }),
            PptxOutline.Style(A.SchemeColorValues.Accent1, 0U, A.SchemeColorValues.Text1));
        _ = doc;
        return (element, ["geometry", "x", "y", "w", "h"]);
    }

    public override string GetRaw() => cxn.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(cxn, raw, doc.Namespaces);
    public override void Remove() => cxn.Remove();

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent is not PptxSlide target) throw new WriterException(ErrorCode.Validation, "Connectors move between slides only", "Use --to /slide[n].");
        cxn.Remove();
        target.Insert(cxn, index);
    }

    public override Node CopyTo(Node newParent, int? index) => newParent is PptxSlide target
        ? target.Place(PptxCopy.Element(cxn, slide, target.Part), index)
        : throw new WriterException(ErrorCode.Validation, "Connectors are copied onto slides", "Use --to /slide[n].");
}

/// <summary>A p:grpSp. Its members are its children, their boxes given in slide units (the group's child space mapped through
/// its transform), so they read and move as if they lay on the slide; moving or resizing the group carries them along.</summary>
sealed class PptxGroup(PptxDocument doc, SlidePart slide, P.GroupShape group) : Node
{
    public override string Kind => "group";
    public override object Anchor => group;
    internal PptxDecor.Transform T { get; init; } = PptxDecor.Transform.Identity;

    A.TransformGroup Xfrm => (group.GroupShapeProperties ??= new P.GroupShapeProperties()).TransformGroup
        ??= new A.TransformGroup(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 0L, Cy = 0L }, new A.ChildOffset { X = 0L, Y = 0L }, new A.ChildExtents { Cx = 0L, Cy = 0L });
    PptxDecor.Transform Inner => T.Then(group.GroupShapeProperties?.TransformGroup);

    protected override IEnumerable<Node> ProjectChildren() => Members(doc, slide, group, Inner);

    /// <summary>The shapes, pictures, connectors, tables and groups inside a group, their boxes mapped by t onto the slide.</summary>
    internal static IEnumerable<Node> Members(PptxDocument doc, SlidePart slide, OpenXmlElement tree, PptxDecor.Transform t)
    {
        foreach (var child in tree.ChildElements)
        {
            switch (child)
            {
                case P.Shape sp: yield return new PptxShape(doc, slide, sp) { T = t }; break;
                case P.Picture pic: yield return new PptxImage(doc, slide, pic) { T = t }; break;
                case P.ConnectionShape cxn: yield return new PptxConnector(doc, slide, cxn) { T = t }; break;
                case P.GraphicFrame frame when PptxTable.TableOf(frame) is not null: yield return new PptxTable(doc, slide, frame) { T = t }; break;
                case P.GroupShape inner: yield return new PptxGroup(doc, slide, inner) { T = t }; break;
            }
        }
    }

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string>();
        var xfrm = group.GroupShapeProperties?.TransformGroup;
        PptxShape.AddBox(props, xfrm?.Offset?.X?.Value, xfrm?.Offset?.Y?.Value, xfrm?.Extents?.Cx?.Value, xfrm?.Extents?.Cy?.Value);
        T.Apply(props);
        if (xfrm?.Rotation?.Value is { } rot && rot != 0) props["rotation"] = (rot / 60000).ToString(CultureInfo.InvariantCulture);
        var nv = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties;
        if (nv?.Name?.Value is { Length: > 0 } name) props["name"] = name;
        if (nv?.Id?.Value is { } id) props["id"] = id.ToString(CultureInfo.InvariantCulture);
        return props;
    }

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "x" or "y" or "w" or "h":
                var xfrm = Xfrm;
                if ((xfrm.ChildExtents?.Cx?.Value ?? 0) == 0 && (xfrm.ChildExtents?.Cy?.Value ?? 0) == 0)
                {   // a group without a child space of its own: its members' coordinates are the slide's
                    xfrm.ChildOffset = new A.ChildOffset { X = xfrm.Offset?.X?.Value ?? 0, Y = xfrm.Offset?.Y?.Value ?? 0 };
                    xfrm.ChildExtents = new A.ChildExtents { Cx = xfrm.Extents?.Cx?.Value ?? 0, Cy = xfrm.Extents?.Cy?.Value ?? 0 };
                }
                PptxOutline.SetBox(xfrm.Offset!, xfrm.Extents!, name, value, T);
                break;
            case "rotation": Xfrm.Rotation = ((int.Parse(value, CultureInfo.InvariantCulture) % 360) + 360) % 360 * 60000; break;
            case "ungroup": if (value == "true") Ungroup(); break;
        }
    }

    /// <summary>The members go back onto the slide where the group stood, at the places the group showed them; the group goes.</summary>
    void Ungroup()
    {
        var t = Inner;
        foreach (var member in group.ChildElements.Where(e => e is P.Shape or P.Picture or P.ConnectionShape or P.GraphicFrame or P.GroupShape).ToList())
        {
            var (off, ext) = member switch
            {
                P.Shape sp => (sp.ShapeProperties?.Transform2D?.Offset, sp.ShapeProperties?.Transform2D?.Extents),
                P.Picture pic => (pic.ShapeProperties?.Transform2D?.Offset, pic.ShapeProperties?.Transform2D?.Extents),
                P.ConnectionShape cxn => (cxn.ShapeProperties?.Transform2D?.Offset, cxn.ShapeProperties?.Transform2D?.Extents),
                P.GraphicFrame frame => (frame.Transform?.Offset, frame.Transform?.Extents),
                P.GroupShape inner => (inner.GroupShapeProperties?.TransformGroup?.Offset, inner.GroupShapeProperties?.TransformGroup?.Extents),
                _ => (null, null),
            };
            if (off is not null && ext is not null)
            {
                (off.X, off.Y) = ((long)Math.Round(t.Ax + t.Sx * (off.X?.Value ?? 0)), (long)Math.Round(t.Ay + t.Sy * (off.Y?.Value ?? 0)));
                (ext.Cx, ext.Cy) = ((long)Math.Round(t.Sx * (ext.Cx?.Value ?? 0)), (long)Math.Round(t.Sy * (ext.Cy?.Value ?? 0)));
            }
            member.Remove();
            group.Parent!.InsertBefore(member, group);
        }
        group.Remove();
    }

    /// <summary>A group of the slide's shapes named in members (paths from the outline, or shape[2] relative to the slide), in
    /// their drawing order, around the box they cover; with no members, an empty group at the box given.</summary>
    public static (OpenXmlElement Element, string[] Consumed) New(PptxDocument doc, SlidePart slide, PptxSlide owner, IReadOnlyDictionary<string, string> props)
    {
        var id = PptxDocument.NextShapeId(slide);
        var element = new P.GroupShape(
            new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = id, Name = "Group " + id }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 0L, Cy = 0L }, new A.ChildOffset { X = 0L, Y = 0L }, new A.ChildExtents { Cx = 0L, Cy = 0L })));
        var members = new List<Node>();
        foreach (var item in (props.GetValueOrDefault("members") ?? "").Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
        {
            var node = item.StartsWith("/slide", StringComparison.Ordinal) ? PathResolver.Single(owner.Parent ?? owner, item) : PathResolver.Single(owner, "/" + item.TrimStart('/'));
            if (!ReferenceEquals(node.Parent?.Anchor, owner.Anchor) || node.Kind is not ("shape" or "image" or "connector" or "table" or "group"))
                throw new WriterException(ErrorCode.Validation, $"members: {item} is not a shape, picture, connector, table or group on {owner.Path}", "Name this slide's own shapes, e.g. members=shape[2],shape[3].");
            if (!members.Any(m => ReferenceEquals(m.Anchor, node.Anchor))) members.Add(node);
        }
        var xfrm = element.GroupShapeProperties!.TransformGroup!;
        long x = PptxShape.Emu(props, "x", 914400L), y = PptxShape.Emu(props, "y", 914400L), w = PptxShape.Emu(props, "w", 914400L), h = PptxShape.Emu(props, "h", 914400L);
        if (members.Count > 0)
        {
            var boxes = new List<(long X, long Y, long W, long H)>();
            foreach (var m in members)
            {
                var p = m.GetProps();
                long L(string k) => long.TryParse(p.GetValueOrDefault(k), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var v) ? v : 0;
                if (m is PptxShape && ((P.Shape)m.Anchor).ShapeProperties?.Transform2D is null) m.SetProp("x", L("x").ToString(CultureInfo.InvariantCulture)); // a placeholder's box becomes its own
                boxes.Add((L("x"), L("y"), L("w"), L("h")));
            }
            (x, y) = (boxes.Min(b => b.X), boxes.Min(b => b.Y));
            (w, h) = (boxes.Max(b => b.X + b.W) - x, boxes.Max(b => b.Y + b.H) - y);
            foreach (var m in members.OrderBy(m => ((OpenXmlElement)m.Anchor).Parent!.ChildElements.ToList().IndexOf((OpenXmlElement)m.Anchor)))
            {
                var e = (OpenXmlElement)m.Anchor;
                e.Remove();
                element.Append(e);
            }
        }
        (xfrm.Offset!.X, xfrm.Offset.Y, xfrm.Extents!.Cx, xfrm.Extents.Cy) = (x, y, w, h);
        (xfrm.ChildOffset!.X, xfrm.ChildOffset.Y, xfrm.ChildExtents!.Cx, xfrm.ChildExtents.Cy) = (x, y, w, h);
        _ = doc;
        return (element, ["members", "x", "y", "w", "h"]);
    }

    public override string GetRaw() => group.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(group, raw, doc.Namespaces);
    public override void Remove() => group.Remove();

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent is not PptxSlide target) throw new WriterException(ErrorCode.Validation, "Groups move between slides only", "Use --to /slide[n].");
        group.Remove();
        target.Insert(group, index);
    }

    public override Node CopyTo(Node newParent, int? index) => newParent is PptxSlide target
        ? target.Place(PptxCopy.Element(group, slide, target.Part), index)
        : throw new WriterException(ErrorCode.Validation, "Groups are copied onto slides", "Use --to /slide[n].");
}
