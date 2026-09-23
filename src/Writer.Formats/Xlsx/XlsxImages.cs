using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Writer.Formats.Xlsx;

/// <summary>A picture on a worksheet: an xdr:pic in the sheet's drawing part, placed by the anchor around it.</summary>
sealed class XlsxImage(XlsxDocument doc, XlsxSheet sheet, Xdr.Picture pic) : PictureNode
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    public override object Anchor => pic;
    protected override OpenXmlCompositeElement Pic => pic;
    protected override OpenXmlPart Owner => sheet.Part.DrawingsPart!;
    OpenXmlCompositeElement Holder => (OpenXmlCompositeElement)pic.Parent!;

    /// <summary>Excel places the picture by its anchor; the picture's own a:xfrm follows so other readers agree.</summary>
    protected override (long X, long Y, long W, long H) Frame
    {
        get
        {
            var box = XlsxAnchors.Read(Holder, sheet.Grid);
            return (box.X, box.Y, box.W, box.H);
        }
        set
        {
            var box = new XlsxAnchors.Box(Math.Max(0, value.X), Math.Max(0, value.Y), Math.Max(1, value.W), Math.Max(1, value.H));
            XlsxAnchors.Write(Holder, box, sheet.Grid);
            if (pic.ShapeProperties?.Transform2D is { Offset: { } off, Extents: { } ext })
                (off.X, off.Y, ext.Cx, ext.Cy) = (box.X, box.Y, box.W, box.H);
        }
    }

    protected override void OwnProps(Dictionary<string, string> props)
    {
        var (x, y, w, h) = Frame;
        props["x"] = x.ToString(Inv);
        props["y"] = y.ToString(Inv);
        props["w"] = w.ToString(Inv);
        props["h"] = h.ToString(Inv);
        var nv = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
        if (nv?.Description?.Value is { Length: > 0 } alt) props["alt"] = alt;
        if (nv?.Id?.Value is { } id) props["id"] = id.ToString(Inv);
    }

    protected override void SetOwnProp(string name, string value)
    {
        switch (name)
        {
            case "x" or "y" or "w" or "h":
                var emu = long.Parse(value, Inv);
                var f = Frame;
                Frame = name switch { "x" => f with { X = emu }, "y" => f with { Y = emu }, "w" => f with { W = emu }, _ => f with { H = emu } };
                break;
            case "alt":
                if (pic.NonVisualPictureProperties?.NonVisualDrawingProperties is { } nv) nv.Description = value.Length > 0 ? value : null;
                break;
        }
    }

    public override string GetRaw() => pic.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(pic, raw, doc.Namespaces.Concat(XlsxAnchors.Namespaces));

    public override void Remove()
    {
        var drawings = sheet.Part.DrawingsPart!;
        var embed = pic.BlipFill?.Blip?.Embed?.Value;
        Holder.Remove();
        if (embed is not null && !drawings.WorksheetDrawing!.Descendants<A.Blip>().Any(b => b.Embed?.Value == embed) && drawings.TryGetPartById(embed, out _))
            drawings.DeletePart(embed);
        if (drawings.WorksheetDrawing?.HasChildren == true) return;
        sheet.Part.Worksheet!.GetFirstChild<Drawing>()?.Remove();
        sheet.Part.DeletePart(drawings);
    }

    public override void MoveTo(Node newParent, int? index) =>
        throw new WriterException(ErrorCode.Validation, "A picture stays on its sheet", "Change x and y to move it on the sheet.");
}

/// <summary>Finds the pictures a sheet already has and adds new ones.</summary>
static class XlsxImages
{
    public static IEnumerable<XlsxImage> Of(XlsxDocument doc, XlsxSheet sheet)
    {
        if (sheet.Part.DrawingsPart?.WorksheetDrawing is not { } drawing) yield break;
        foreach (var anchor in drawing.ChildElements.OfType<OpenXmlCompositeElement>())
            if (anchor.GetFirstChild<Xdr.Picture>() is { } pic) yield return new XlsxImage(doc, sheet, pic);
    }

    /// <summary>A picture from a file or data URL, at its pixel size (96 dpi, at most 15 cm wide) to the right of the used range unless placed.</summary>
    public static XlsxImage Add(XlsxDocument doc, XlsxSheet sheet, IReadOnlyDictionary<string, string> props)
    {
        var src = props.GetValueOrDefault("src") ?? throw new WriterException(ErrorCode.Validation, "An image needs src", "Example: --prop src=logo.png");
        var (bytes, name) = ImageInfo.Load(src);
        var info = ImageInfo.Read(bytes);
        var drawings = XlsxCharts.Drawings(sheet);
        var drawing = drawings.WorksheetDrawing!;
        var image = drawings.AddImagePart(info.ContentType);
        image.FeedData(new MemoryStream(bytes));
        var aspect = info.Height / (double)Math.Max(1, info.Width);
        var width = Math.Min(info.Width * Units.EmuPerPx, 15 * Units.EmuPerCm);
        if (props.TryGetValue("w", out var w) && !props.ContainsKey("h")) width = long.Parse(w, CultureInfo.InvariantCulture);
        var height = props.TryGetValue("h", out var h) && !props.ContainsKey("w") ? long.Parse(h, CultureInfo.InvariantCulture) : (long)Math.Round(width * aspect);
        if (props.ContainsKey("h") && !props.ContainsKey("w")) width = (long)Math.Round(height / aspect);
        var box = XlsxCharts.DefaultBox(sheet, props, width, height);
        var id = drawing.Descendants<Xdr.NonVisualDrawingProperties>().Select(p => p.Id?.Value ?? 0u).DefaultIfEmpty(1u).Max() + 1;
        var pic = new Xdr.Picture(
            new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties { Id = id, Name = name, Description = props.GetValueOrDefault("alt") is { Length: > 0 } alt ? alt : null },
                new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
            new Xdr.BlipFill(new A.Blip { Embed = drawings.GetIdOfPart(image) }, new A.Stretch(new A.FillRectangle())),
            new Xdr.ShapeProperties(
                new A.Transform2D(new A.Offset { X = box.X, Y = box.Y }, new A.Extents { Cx = box.W, Cy = box.H }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        var anchor = new Xdr.TwoCellAnchor(new Xdr.FromMarker(), new Xdr.ToMarker(), pic, new Xdr.ClientData()) { EditAs = Xdr.EditAsValues.OneCell };
        drawing.Append(anchor);
        XlsxAnchors.Write(anchor, box, sheet.Grid);
        return new XlsxImage(doc, sheet, pic);
    }
}
