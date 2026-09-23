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

/// <summary>A paragraph whose only content is one picture.</summary>
sealed class DocxImage(DocxDocument doc, W.Paragraph p) : PictureNode
{
    /// <summary>A4 content width with 2.54cm margins, in EMU.</summary>
    const long MaxWidthEmu = 9026L * 635;

    public override object Anchor => p;

    public static bool IsPictureParagraph(W.Paragraph p) =>
        DocxRuns.ParagraphText(p).Length == 0
        && p.Descendants<W.Drawing>().Count() == 1
        && p.Descendants<PIC.Picture>().Count() == 1;

    W.Drawing Drawing => p.Descendants<W.Drawing>().First();
    protected override OpenXmlCompositeElement Pic => Drawing.Descendants<PIC.Picture>().First();
    protected override OpenXmlPart Owner => doc.Main;

    protected override (long X, long Y, long W, long H) Frame
    {
        get
        {
            var extent = Drawing.Descendants<DW.Extent>().FirstOrDefault();
            return (0, 0, extent?.Cx?.Value ?? 0, extent?.Cy?.Value ?? 0);
        }
        set
        {
            if (Drawing.Descendants<DW.Extent>().FirstOrDefault() is { } extent) (extent.Cx, extent.Cy) = (value.W, value.H);
            if (Drawing.Descendants<A.Extents>().FirstOrDefault() is { } extents) (extents.Cx, extents.Cy) = (value.W, value.H);
        }
    }

    /// <summary>Word reserves room around an inline picture for what it draws past its frame: the rotated corners, border and shadow.</summary>
    protected override void Reframe()
    {
        var (l, t, r, b) = Overhang();
        var holder = (OpenXmlCompositeElement?)Drawing.GetFirstChild<DW.Inline>() ?? Drawing.GetFirstChild<DW.Anchor>();
        if (holder is null) return;
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
        if (Drawing.Descendants<DW.Extent>().Any())
        {
            props["width"] = w.ToString(CultureInfo.InvariantCulture);
            props["height"] = h.ToString(CultureInfo.InvariantCulture);
        }
        if (Drawing.Descendants<DW.DocProperties>().FirstOrDefault()?.Description?.Value is { Length: > 0 } alt) props["alt"] = alt;
    }

    public override string GetRaw() => p.OuterXml;

    /// <summary>A new picture paragraph from an image file. Size defaults to the pixel size at 96 dpi, capped at the page width.</summary>
    public static (OpenXmlElement Element, string[] Consumed) New(DocxDocument doc, IReadOnlyDictionary<string, string> props)
    {
        var src = props.GetValueOrDefault("src")
            ?? throw new WriterException(ErrorCode.Validation, "An image needs src", "Example: --prop src=chart.png");
        var (relId, info, name) = Embed(doc, src);
        var (cx, cy) = Size(info, props.GetValueOrDefault("width"), props.GetValueOrDefault("height"));
        var drawing = BuildDrawing(NextId(doc), name, relId, cx, cy, props.GetValueOrDefault("alt"));
        return (new W.Paragraph(new W.Run(drawing)), ["src", "width", "height", "alt"]);
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
            return (long.Parse(width, CultureInfo.InvariantCulture), long.Parse(height, CultureInfo.InvariantCulture));
        if (width is not null)
        {
            var w = long.Parse(width, CultureInfo.InvariantCulture);
            return (w, (long)Math.Round(w * aspect));
        }
        if (height is not null)
        {
            var h = long.Parse(height, CultureInfo.InvariantCulture);
            return ((long)Math.Round(h / aspect), h);
        }
        var natural = Math.Min(info.Width * Units.EmuPerPx, MaxWidthEmu);
        return (natural, (long)Math.Round(natural * aspect));
    }

    static uint NextId(DocxDocument doc) =>
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
                var emu = long.Parse(value, CultureInfo.InvariantCulture);
                Frame = name == "width" ? Frame with { W = emu } : Frame with { H = emu };
                Reframe();
                break;
            case "alt":
                if (Drawing.Descendants<DW.DocProperties>().FirstOrDefault() is { } docPr) docPr.Description = value.Length > 0 ? new StringValue(value) : null;
                break;
        }
    }

    public override void Remove() => DocxBlocks.Detach(p);
    public override void MoveTo(Node newParent, int? index) => DocxBlocks.Move(newParent, p, index);
    public override void SetRaw(string raw) => DocxBlocks.ReplaceRaw(doc, p, raw);
}
