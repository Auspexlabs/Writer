using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>A picture. A paragraph whose only content is one inline picture is a block of its own (/body/image[n]); a picture
/// in any other paragraph is that paragraph's child (/body/paragraph[n]/image[k]), usually floating: wp:anchor, placed by
/// positionH / positionV offsets or alignments from the page, margin, column or paragraph, with its text wrapping. Moving a
/// picture into a paragraph makes it float there; moving it to the body makes it a block again, inline.</summary>
sealed class DocxImage(DocxDocument doc, W.Drawing drawing, W.Paragraph? block = null) : PictureNode
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    /// <summary>A4 content width with 2.54cm margins, in EMU.</summary>
    const long MaxWidthEmu = 9026L * 635;
    /// <summary>Word's gap between a floating picture and the text around it, left and right (0.125in).</summary>
    const uint SideGap = 114300;

    public override object Anchor => drawing;

    public static DocxImage Block(DocxDocument doc, W.Paragraph p) => new(doc, p.Descendants<W.Drawing>().First(), p);

    public static bool IsPictureParagraph(W.Paragraph p) =>
        DocxRuns.ParagraphText(p).Length == 0
        && p.Descendants<W.Drawing>().Count() == 1
        && p.Descendants<PIC.Picture>().Count() == 1
        && p.Descendants<DW.Inline>().Any();

    /// <summary>The pictures in a paragraph's own runs (not in text boxes or alternate content).</summary>
    public static IEnumerable<W.Drawing> In(W.Paragraph p) =>
        p.Descendants<W.Drawing>().Where(d => d.Parent is W.Run && d.Ancestors<W.Paragraph>().First() == p && PicOf(d) is not null);

    static PIC.Picture? PicOf(W.Drawing d) =>
        ((OpenXmlElement?)d.GetFirstChild<DW.Inline>() ?? d.GetFirstChild<DW.Anchor>())?.GetFirstChild<A.Graphic>()?.GraphicData?.GetFirstChild<PIC.Picture>();

    DW.Anchor? Floating => drawing.GetFirstChild<DW.Anchor>();
    OpenXmlCompositeElement Holder => (OpenXmlCompositeElement?)drawing.GetFirstChild<DW.Inline>() ?? drawing.GetFirstChild<DW.Anchor>()!;
    DW.DocProperties? DocPr => Holder.GetFirstChild<DW.DocProperties>();
    protected override OpenXmlCompositeElement Pic => PicOf(drawing)!;
    protected override OpenXmlPart Owner => doc.Main;

    protected override (long X, long Y, long W, long H) Frame
    {
        get
        {
            var extent = Holder.GetFirstChild<DW.Extent>();
            return (0, 0, extent?.Cx?.Value ?? 0, extent?.Cy?.Value ?? 0);
        }
        set
        {
            if (Holder.GetFirstChild<DW.Extent>() is { } extent) (extent.Cx, extent.Cy) = (value.W, value.H);
            if (drawing.Descendants<A.Extents>().FirstOrDefault() is { } extents) (extents.Cx, extents.Cy) = (value.W, value.H);
        }
    }

    /// <summary>Word reserves room around a picture for what it draws past its frame: the rotated corners, border and shadow.</summary>
    protected override void Reframe()
    {
        var (l, t, r, b) = Overhang();
        var holder = Holder;
        var effect = holder.GetFirstChild<DW.EffectExtent>();
        if (effect is null)
        {
            effect = new DW.EffectExtent();
            holder.InsertAfter(effect, holder.GetFirstChild<DW.Extent>());
        }
        (effect.LeftEdge, effect.TopEdge, effect.RightEdge, effect.BottomEdge) = (l, t, r, b);
    }

    protected override void OwnProps(Dictionary<string, string> props)
    {
        var (_, _, w, h) = Frame;
        if (Holder.GetFirstChild<DW.Extent>() is not null)
        {
            props["width"] = w.ToString(Inv);
            props["height"] = h.ToString(Inv);
        }
        if (DocPr?.Description?.Value is { Length: > 0 } alt) props["alt"] = alt;
        if (DocPr?.Id?.Value is { } id) props["id"] = id.ToString(Inv);
        if (Floating is not { } a)
        {
            props["wrap"] = "inline";
            return;
        }
        ReadPlace(a, props);
    }

    /// <summary>A floating object's wrap and place (a picture's, a shape's).</summary>
    internal static void ReadPlace(DW.Anchor a, Dictionary<string, string> props)
    {
        props["wrap"] = WrapOf(a);
        if (a.HorizontalPosition is { } hp)
        {
            if (hp.RelativeFrom?.InnerText is { } from) props["xFrom"] = from;
            if (hp.PositionOffset?.Text is { } x) props["x"] = x.Trim();
            else if (hp.HorizontalAlignment?.Text is { } align) props["xAlign"] = align.Trim();
        }
        if (a.VerticalPosition is { } vp)
        {
            if (vp.RelativeFrom?.InnerText is { } from) props["yFrom"] = from;
            if (vp.PositionOffset?.Text is { } y) props["y"] = y.Trim();
            else if (vp.VerticalAlignment?.Text is { } align) props["yAlign"] = align.Trim();
        }
    }

    static string WrapOf(DW.Anchor a) =>
        a.GetFirstChild<DW.WrapSquare>() is not null ? "square"
        : a.GetFirstChild<DW.WrapTight>() is not null ? "tight"
        : a.GetFirstChild<DW.WrapThrough>() is not null ? "through"
        : a.GetFirstChild<DW.WrapTopBottom>() is not null ? "topBottom"
        : a.BehindDoc?.Value == true ? "behind" : "front";

    public override string GetRaw() => block?.OuterXml ?? drawing.Parent!.OuterXml;

    /// <summary>A new picture paragraph from an image file. Size defaults to the pixel size at 96 dpi, capped at the page width.</summary>
    public static (OpenXmlElement Element, string[] Consumed) New(DocxDocument doc, IReadOnlyDictionary<string, string> props)
    {
        var drawing = NewDrawing(doc, props);
        return (new W.Paragraph(new W.Run(drawing)), ["src", "width", "height", "alt"]);
    }

    /// <summary>A picture added under a paragraph floats there (square wrap, at the top left of the column, unless props say otherwise).</summary>
    public static Node AddTo(DocxDocument doc, Node paragraph, W.Paragraph p, IReadOnlyDictionary<string, string> props, int? index)
    {
        var drawing = NewDrawing(doc, props);
        DocxBlocks.InsertAt(paragraph, p, new W.Run(drawing), index);
        var node = new DocxImage(doc, drawing);
        node.Float();
        foreach (var (name, value) in props)
            if (name is not ("src" or "width" or "height" or "alt")) node.SetProp(name, value);
        return node;
    }

    static W.Drawing NewDrawing(DocxDocument doc, IReadOnlyDictionary<string, string> props)
    {
        var src = props.GetValueOrDefault("src")
            ?? throw new WriterException(ErrorCode.Validation, "An image needs src", "Example: --prop src=chart.png");
        var (relId, info, name) = Embed(doc, src);
        var (cx, cy) = Size(info, props.GetValueOrDefault("width"), props.GetValueOrDefault("height"));
        return BuildDrawing(NextId(doc), name, relId, cx, cy, props.GetValueOrDefault("alt"));
    }

    static (string RelId, ImageInfo Info, string Name) Embed(DocxDocument doc, string src)
    {
        var (bytes, name) = ImageInfo.Load(src);
        var info = ImageInfo.Read(bytes);
        var part = doc.Main.AddImagePart(info.ContentType);
        part.FeedData(new MemoryStream(bytes));
        return (doc.Main.GetIdOfPart(part), info, name);
    }

    static (long Cx, long Cy) Size(ImageInfo info, string? width, string? height)
    {
        var aspect = info.Height / (double)Math.Max(1, info.Width);
        if (width is not null && height is not null)
            return (long.Parse(width, Inv), long.Parse(height, Inv));
        if (width is not null)
        {
            var w = long.Parse(width, Inv);
            return (w, (long)Math.Round(w * aspect));
        }
        if (height is not null)
        {
            var h = long.Parse(height, Inv);
            return ((long)Math.Round(h / aspect), h);
        }
        var natural = Math.Min(info.Width * Units.EmuPerPx, MaxWidthEmu);
        return (natural, (long)Math.Round(natural * aspect));
    }

    internal static uint NextId(DocxDocument doc) =>
        (doc.Main.Document!.Descendants<DW.DocProperties>().Select(d => d.Id?.Value).Max() ?? 0u) + 1;

    static W.Drawing BuildDrawing(uint id, string name, string relId, long cx, long cy, string? alt) => new(
        new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.DocProperties { Id = id, Name = name, Description = string.IsNullOrEmpty(alt) ? null : new StringValue(alt) }, // a plain null would still write descr=""
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(
                new PIC.Picture(
                    new PIC.NonVisualPictureProperties(
                        new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(new A.Blip { Embed = relId }, new A.Stretch(new A.FillRectangle())),
                    new PIC.ShapeProperties(
                        new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });

    protected override void SetOwnProp(string name, string value)
    {
        switch (name)
        {
            case "width" or "height":
                var emu = long.Parse(value, Inv);
                Frame = name == "width" ? Frame with { W = emu } : Frame with { H = emu };
                Reframe();
                break;
            case "alt":
                if (DocPr is { } docPr) docPr.Description = value.Length > 0 ? new StringValue(value) : null;
                break;
            case "wrap" when value == "inline":
                Inline();
                break;
            case "wrap" or "x" or "xAlign" or "y" or "yAlign" or "xFrom" or "yFrom":
                SetPlace(Float(), name, value);
                break;
        }
    }

    /// <summary>A floating object's wrap, offsets, alignments and what they measure from.</summary>
    internal static void SetPlace(DW.Anchor a, string name, string value)
    {
        switch (name)
        {
            case "wrap":
                SetWrap(a, value);
                break;
            case "x" or "xAlign":
                var h = a.HorizontalPosition!;
                h.RemoveAllChildren();
                h.Append(name == "x" ? new DW.PositionOffset(value) : new DW.HorizontalAlignment(value));
                break;
            case "y" or "yAlign":
                var v = a.VerticalPosition!;
                v.RemoveAllChildren();
                v.Append(name == "y" ? new DW.PositionOffset(value) : new DW.VerticalAlignment(value));
                break;
            case "xFrom":
                a.HorizontalPosition!.RelativeFrom = new DW.HorizontalRelativePositionValues(value);
                break;
            case "yFrom":
                a.VerticalPosition!.RelativeFrom = new DW.VerticalRelativePositionValues(value);
                break;
        }
    }

    // ---- floating: wp:inline ⇄ wp:anchor, the wrap and the place ----

    /// <summary>The picture's wp:anchor, made from its wp:inline when it is inline: square wrap, at the top left of the column where
    /// its paragraph starts, in front of the floating objects already there. A block picture floats in its own paragraph.</summary>
    DW.Anchor Float()
    {
        if (Floating is { } anchor) return anchor;
        var inline = drawing.GetFirstChild<DW.Inline>()!;
        anchor = new DW.Anchor
        {
            DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = SideGap, DistanceFromRight = SideGap, SimplePos = false,
            RelativeHeight = NextZ(), BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true,
            AnchorId = inline.AnchorId, EditId = inline.EditId,
        };
        anchor.Append(new DW.SimplePosition { X = 0L, Y = 0L },
            new DW.HorizontalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.HorizontalRelativePositionValues.Column },
            new DW.VerticalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph });
        foreach (var child in inline.ChildElements.ToList())
        {
            child.Remove();
            anchor.Append(child);
        }
        anchor.InsertBefore(new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides }, anchor.GetFirstChild<DW.DocProperties>());
        inline.InsertBeforeSelf(anchor);
        inline.Remove();
        return anchor;
    }

    /// <summary>Back in the line of text: extent, effect extent, docPr, frame locks and graphic stay; the place and the wrap go.</summary>
    void Inline()
    {
        if (Floating is not { } anchor) return;
        var inline = new DW.Inline { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U, AnchorId = anchor.AnchorId, EditId = anchor.EditId };
        foreach (var child in anchor.ChildElements.ToList())
        {
            if (child is not (DW.Extent or DW.EffectExtent or DW.DocProperties or DW.NonVisualGraphicFrameDrawingProperties or A.Graphic)) continue;
            child.Remove();
            inline.Append(child);
        }
        anchor.InsertBeforeSelf(inline);
        anchor.Remove();
    }

    static void SetWrap(DW.Anchor anchor, string mode)
    {
        foreach (var old in anchor.ChildElements.Where(e => e is DW.WrapNone or DW.WrapSquare or DW.WrapTight or DW.WrapThrough or DW.WrapTopBottom).ToList()) old.Remove();
        OpenXmlElement wrap = mode switch
        {
            "square" => new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
            "tight" => new DW.WrapTight(Outline()) { WrapText = DW.WrapTextValues.BothSides },
            "through" => new DW.WrapThrough(Outline()) { WrapText = DW.WrapTextValues.BothSides },
            "topBottom" => new DW.WrapTopBottom(),
            _ => new DW.WrapNone(),
        };
        anchor.BehindDoc = mode == "behind";
        anchor.InsertBefore(wrap, anchor.GetFirstChild<DW.DocProperties>());
    }

    /// <summary>The whole frame as a wrap polygon (21600 is its full width and height), as Word writes it for a new tight wrap.</summary>
    static DW.WrapPolygon Outline() => new(new DW.StartPoint { X = 0L, Y = 0L }, new DW.LineTo { X = 0L, Y = 21600L }, new DW.LineTo { X = 21600L, Y = 21600L },
        new DW.LineTo { X = 21600L, Y = 0L }, new DW.LineTo { X = 0L, Y = 0L }) { Edited = false };

    /// <summary>A z-order above every floating object in the document (Word counts up from 251658240).</summary>
    uint NextZ() => NextZ(doc);
    internal static uint NextZ(DocxDocument doc) => Math.Max(251658240u, doc.Main.Document!.Descendants<DW.Anchor>().Select(a => a.RelativeHeight?.Value ?? 0u).DefaultIfEmpty(0u).Max()) + 1;

    // ---- where it lives ----

    public override void Remove()
    {
        if (block is not null)
        {
            DocxBlocks.Detach(block);
            return;
        }
        var run = drawing.Parent;
        drawing.Remove();
        if (run is W.Run r && !r.ChildElements.Any(e => e is not W.RunProperties)) Cut(r);
    }

    /// <summary>Into a paragraph: the picture floats there, at the end of its runs. Into the body or a cell: a paragraph of its own,
    /// inline. The paragraph a block picture leaves goes with it.</summary>
    public override void MoveTo(Node newParent, int? index)
    {
        var container = DocxBlocks.ContainerOf(newParent);
        if (block is not null && container is not W.Paragraph)
        {
            DocxBlocks.Move(newParent, block, index);
            return;
        }
        if (block is not null && container.Ancestors().Contains(block))
            throw new WriterException(ErrorCode.Validation, "Cannot move an element into itself", "Pick another target.");
        var run = TakeRun();
        if (container is W.Paragraph)
        {
            DocxBlocks.InsertAt(newParent, container, run, index);
            Float();
            return;
        }
        Inline();
        DocxBlocks.InsertAt(newParent, container, new W.Paragraph(run), index);
    }

    /// <summary>The picture's run, out of its place: its own run when that holds nothing else, else a new one with its formatting.</summary>
    W.Run TakeRun()
    {
        var old = drawing.Parent as W.Run;
        if (block is not null) DocxBlocks.Detach(block);
        if (old is not null && old.ChildElements.All(e => e is W.RunProperties || ReferenceEquals(e, drawing)))
        {
            Cut(old);
            return old;
        }
        drawing.Remove();
        var run = new W.Run();
        if (old?.RunProperties is { } rp) run.Append(rp.CloneNode(true));
        run.Append(drawing);
        return run;
    }

    /// <summary>Takes a run out of its paragraph, and the link or revision mark it leaves empty.</summary>
    internal static void Cut(W.Run run)
    {
        var parent = run.Parent;
        run.Remove();
        while (parent is W.Hyperlink or W.InsertedRun or W.DeletedRun && !parent.HasChildren && parent.Parent is { } up)
        {
            parent.Remove();
            parent = up;
        }
    }

    public override void SetRaw(string raw)
    {
        if (block is not null) DocxBlocks.ReplaceRaw(doc, block, raw);
        else DocxBlocks.ReplaceRaw(doc, drawing.Parent!, raw);
    }
}
