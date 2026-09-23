using System.Diagnostics;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace Writer.Formats.Common;

/// <summary>
/// A DrawingML picture: pic:pic in Word, p:pic in PowerPoint, xdr:pic in Excel. The picture tools work alike in the three
/// formats and write Office's own markup — a:srcRect, a:xfrm, a:lum, a:grayscl, a:alphaModFix, a:ln, a:outerShdw, a:prstGeom —
/// so Word, PowerPoint and Excel show what the editor shows. A subclass supplies the frame and the format's own props.
/// </summary>
abstract class PictureNode : Node
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /// <summary>Relationship id prefix of the picture kept for reset after a background removal: kept + the id the picture shows now.
    /// Office leaves a relationship nothing refers to alone (Word drops it on its next save, and reset then only drops adjustments).</summary>
    const string KeptPrefix = "wrOrig";

    public override string Kind => "image";

    /// <summary>The picture element; its blipFill and spPr hold everything the tools write.</summary>
    protected abstract OpenXmlCompositeElement Pic { get; }
    /// <summary>The part whose relationships hold the image: document, slide or drawing part.</summary>
    protected abstract OpenXmlPart Owner { get; }
    /// <summary>The frame in EMU. Word lays inline pictures out in the text, so X and Y stay 0 there.</summary>
    protected abstract (long X, long Y, long W, long H) Frame { get; set; }
    /// <summary>Props the format keeps itself: size or position, alt text, id.</summary>
    protected abstract void OwnProps(Dictionary<string, string> props);
    protected abstract void SetOwnProp(string name, string value);
    /// <summary>Runs after an edit that changes how far the picture draws beyond its frame (rotation, border, shadow).</summary>
    protected virtual void Reframe() { }

    OpenXmlCompositeElement Child(string localName) => Pic.ChildElements.OfType<OpenXmlCompositeElement>().FirstOrDefault(e => e.LocalName == localName)
        ?? throw new WriterException(ErrorCode.FormatError, $"This picture has no {localName}", "Remove it and add the picture again.");
    OpenXmlCompositeElement BlipFill => Child("blipFill");
    OpenXmlCompositeElement SpPr => Child("spPr");
    A.Blip? Blip => BlipFill.GetFirstChild<A.Blip>();
    A.Transform2D? Xfrm => SpPr.GetFirstChild<A.Transform2D>();
    A.OuterShadow? Shadow => SpPr.GetFirstChild<A.EffectList>()?.GetFirstChild<A.OuterShadow>();

    OpenXmlPart? ImagePart(string? relId) => relId is not null && Owner.TryGetPartById(relId, out var part) ? part : null;

    public sealed override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string>();
        var image = ImagePart(Blip?.Embed?.Value);
        if (image is not null) props["src"] = image.Uri.ToString();
        OwnProps(props);
        if (image is not null)
        {
            using var stream = image.GetStream(FileMode.Open, FileAccess.Read);
            props["bytes"] = stream.Length.ToString(Inv);
        }
        ReadLook(props);
        return props;
    }

    public override (string ContentType, byte[] Data)? GetBinary() => ImagePart(Blip?.Embed?.Value) is { } part ? (part.ContentType, Read(part)) : null;

    static byte[] Read(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    void ReadLook(Dictionary<string, string> props)
    {
        var crop = Crop();
        if (crop != default) props["crop"] = string.Join(",", new[] { crop.L, crop.T, crop.R, crop.B }.Select(v => (v * 100).ToString("0.###", Inv)));
        if (Xfrm is { } xfrm)
        {
            if (Degrees(xfrm) is var deg and not 0) props["rotation"] = deg.ToString(Inv);
            if (xfrm.HorizontalFlip?.Value == true) props["flipH"] = "true";
            if (xfrm.VerticalFlip?.Value == true) props["flipV"] = "true";
        }
        if (Blip is { } blip)
        {
            if (blip.GetFirstChild<A.LuminanceEffect>() is { } lum)
            {
                if ((lum.Brightness?.Value ?? 0) is var b and not 0) props["brightness"] = Percent(b);
                if ((lum.Contrast?.Value ?? 0) is var c and not 0) props["contrast"] = Percent(c);
            }
            if (blip.GetFirstChild<A.Grayscale>() is not null) props["grayscale"] = "true";
            if (blip.GetFirstChild<A.AlphaModulationFixed>() is { } alpha && (alpha.Amount?.Value ?? 100000) < 100000)
                props["transparency"] = (100 - (int)Math.Round((alpha.Amount?.Value ?? 100000) / 1000.0)).ToString(Inv);
        }
        if (SpPr.GetFirstChild<A.Outline>() is { } ln && ln.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value is { } color)
        {
            props["line"] = color.ToUpperInvariant();
            props["lineWidth"] = Math.Round((ln.Width?.Value ?? 9525) / (double)Units.EmuPerPt, 2).ToString("0.##", Inv);
        }
        if (Shadow is not null) props["shadow"] = "true";
        if (SpPr.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText is { } shape && shape != "rect") props["geometry"] = shape;
    }

    static string Percent(int thousandths) => ((int)Math.Round(thousandths / 1000.0)).ToString(Inv);
    static int Degrees(A.Transform2D xfrm) => (int)Math.Round((xfrm.Rotation?.Value ?? 0) / 60000.0) % 360;

    public sealed override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "src": Replace(value); break;
            case "background": Show(Vision.Cutout(Current()), keepOriginal: true); break;
            case "compress": Compress(value switch { "web" => 150, "email" => 96, _ => 220 }); break;
            case "reset": if (value == "true") Reset(); break;
            case "crop": SetCrop(ParseCrop(value)); break;
            case "rotation":
                var deg = ((int.Parse(value, Inv) % 360) + 360) % 360;
                EnsureXfrm().Rotation = deg == 0 ? null : deg * 60000;
                Reframe();
                break;
            case "flipH": EnsureXfrm().HorizontalFlip = value == "true" ? true : null; break;
            case "flipV": EnsureXfrm().VerticalFlip = value == "true" ? true : null; break;
            case "brightness" or "contrast": SetLum(name, int.Parse(value, Inv)); break;
            case "grayscale":
                var blip = RequireBlip();
                blip.RemoveAllChildren<A.Grayscale>();
                if (value == "true") AddEffect(blip, new A.Grayscale());
                break;
            case "transparency":
                var b = RequireBlip();
                b.RemoveAllChildren<A.AlphaModulationFixed>();
                if (int.Parse(value, Inv) is var t and > 0) AddEffect(b, new A.AlphaModulationFixed { Amount = (100 - t) * 1000 });
                break;
            case "line" or "lineWidth": SetLine(name, value); Reframe(); break;
            case "shadow": SetShadow(value == "true"); Reframe(); break;
            case "geometry": SetGeometry(value); break;
            default: SetOwnProp(name, value); break;
        }
    }

    A.Blip RequireBlip() => Blip ?? throw new WriterException(ErrorCode.Validation, "This picture has no image data", "Add the picture again.");

    /// <summary>A colour effect goes before the blip's extension list; the effects are a repeating choice, which the SDK's AddChild would replace.</summary>
    static void AddEffect(A.Blip blip, OpenXmlElement effect)
    {
        if (blip.GetFirstChild<A.BlipExtensionList>() is { } extensions) blip.InsertBefore(effect, extensions);
        else blip.AppendChild(effect);
    }

    A.Transform2D EnsureXfrm()
    {
        if (Xfrm is { } xfrm) return xfrm;
        var (x, y, w, h) = Frame;
        var created = new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = w, Cy = h });
        SpPr.PrependChild(created);
        return created;
    }

    // ---- crop: a:srcRect in thousandths of a percent; the frame keeps the picture's scale and the visible part stays put ----

    readonly record struct Edges(double L, double T, double R, double B);

    Edges Crop() => BlipFill.GetFirstChild<A.SourceRectangle>() is { } r
        ? new((r.Left?.Value ?? 0) / 100000.0, (r.Top?.Value ?? 0) / 100000.0, (r.Right?.Value ?? 0) / 100000.0, (r.Bottom?.Value ?? 0) / 100000.0)
        : default;

    static Edges ParseCrop(string value)
    {
        var s = value.Trim();
        if (s is "" or "0" or "none") return default;
        var parts = s.Split(',');
        var n = new double[4];
        if (parts.Length != 4 || parts.Where((p, i) => !double.TryParse(p.Trim().TrimEnd('%'), NumberStyles.Float, Inv, out n[i])).Any()
            || n.Any(v => v is < 0 or >= 100 || !double.IsFinite(v)) || n[0] + n[2] > 99 || n[1] + n[3] > 99)
            throw new WriterException(ErrorCode.Validation, $"crop: '{value}' is not valid",
                "Expected left,top,right,bottom in percent of the picture, each side leaving at least 1% visible. Example: crop=10,0,10,0; 0,0,0,0 removes the crop.");
        return new(n[0] / 100, n[1] / 100, n[2] / 100, n[3] / 100);
    }

    void SetCrop(Edges c)
    {
        var old = Crop();
        if (old == c) return;
        var (x, y, w, h) = Frame;
        double fw = w / Math.Max(0.01, 1 - old.L - old.R), fh = h / Math.Max(0.01, 1 - old.T - old.B); // the whole picture at its current scale
        double nw = fw * (1 - c.L - c.R), nh = fh * (1 - c.T - c.B);
        // the centre moves by what the edges took, turned with the frame, so the part still shown stays where it was
        double dx = fw * ((c.L - c.R) - (old.L - old.R)) / 2, dy = fh * ((c.T - c.B) - (old.T - old.B)) / 2;
        var xfrm = Xfrm;
        if (xfrm?.HorizontalFlip?.Value == true) dx = -dx;
        if (xfrm?.VerticalFlip?.Value == true) dy = -dy;
        var a = (xfrm?.Rotation?.Value ?? 0) / 60000.0 * Math.PI / 180;
        double cx = x + w / 2.0 + dx * Math.Cos(a) - dy * Math.Sin(a), cy = y + h / 2.0 + dx * Math.Sin(a) + dy * Math.Cos(a);
        WriteCrop(c);
        Frame = ((long)Math.Round(cx - nw / 2), (long)Math.Round(cy - nh / 2), Math.Max(1, (long)Math.Round(nw)), Math.Max(1, (long)Math.Round(nh)));
        Reframe();
    }

    void WriteCrop(Edges c)
    {
        var fill = BlipFill;
        fill.RemoveAllChildren<A.SourceRectangle>();
        if (c == default) return;
        static Int32Value? At(double v) => v == 0 ? null : (int)Math.Round(v * 100000);
        var rect = new A.SourceRectangle { Left = At(c.L), Top = At(c.T), Right = At(c.R), Bottom = At(c.B) };
        if (Blip is { } blip) fill.InsertAfter(rect, blip);
        else fill.PrependChild(rect);
    }

    // ---- colour, border, shadow, shape ----

    void SetLum(string name, int value)
    {
        var blip = RequireBlip();
        var lum = blip.GetFirstChild<A.LuminanceEffect>();
        if (lum is null)
        {
            if (value == 0) return;
            lum = new A.LuminanceEffect();
            AddEffect(blip, lum);
        }
        if (name == "brightness") lum.Brightness = value == 0 ? null : value * 1000;
        else lum.Contrast = value == 0 ? null : value * 1000;
        if (lum.Brightness is null && lum.Contrast is null) lum.Remove();
    }

    void SetLine(string name, string value)
    {
        var spPr = SpPr;
        var ln = spPr.GetFirstChild<A.Outline>();
        if (name == "line" && value == "none")
        {
            ln?.Remove();
            return;
        }
        if (ln is null || ln.GetFirstChild<A.SolidFill>() is null)
        {
            ln?.Remove();
            ln = new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "000000" })) { Width = 9525 };
            spPr.AddChild(ln);
        }
        if (name == "line")
        {
            var fill = ln.GetFirstChild<A.SolidFill>()!;
            fill.RemoveAllChildren();
            fill.Append(new A.RgbColorModelHex { Val = value });
        }
        else ln.Width = (int)Math.Round(double.Parse(value, Inv) * Units.EmuPerPt);
    }

    void SetShadow(bool on)
    {
        var spPr = SpPr;
        var effects = spPr.GetFirstChild<A.EffectList>();
        effects?.RemoveAllChildren<A.OuterShadow>();
        if (on)
        {
            if (effects is null)
            {
                effects = new A.EffectList();
                spPr.AddChild(effects);
            }
            // PowerPoint's "Offset: Bottom Right" outer shadow
            effects.AddChild(new A.OuterShadow(new A.PresetColor(new A.Alpha { Val = 40000 }) { Val = A.PresetColorValues.Black })
            { BlurRadius = 50800, Distance = 38100, Direction = 2700000, Alignment = A.RectangleAlignmentValues.TopLeft, RotateWithShape = false });
        }
        else if (effects is { HasChildren: false }) effects.Remove();
    }

    void SetGeometry(string shape)
    {
        var spPr = SpPr;
        spPr.RemoveAllChildren<A.PresetGeometry>();
        spPr.RemoveAllChildren<A.CustomGeometry>();
        spPr.AddChild(new A.PresetGeometry(new A.AdjustValueList()) { Preset = new A.ShapeTypeValues(shape) });
    }

    /// <summary>How far the rotated frame, half the border and the shadow reach past the frame on each side, in EMU (Word's effect extent).</summary>
    protected (long L, long T, long R, long B) Overhang()
    {
        var (_, _, w, h) = Frame;
        var a = (Xfrm?.Rotation?.Value ?? 0) / 60000.0 * Math.PI / 180;
        double side = (Math.Abs(w * Math.Cos(a)) + Math.Abs(h * Math.Sin(a)) - w) / 2, top = (Math.Abs(w * Math.Sin(a)) + Math.Abs(h * Math.Cos(a)) - h) / 2;
        var ln = SpPr.GetFirstChild<A.Outline>();
        var line = ln?.GetFirstChild<A.SolidFill>() is null ? 0 : (ln.Width?.Value ?? 9525) / 2.0;
        double reach = 0;
        if (Shadow is { } s) reach = (s.Distance?.Value ?? 0) * Math.Cos(Math.PI / 4) + (s.BlurRadius?.Value ?? 0) / 2.0;
        long R(double v) => (long)Math.Round(v);
        return (R(side + line), R(top + line), R(side + line + reach), R(top + line + reach));
    }

    // ---- the image itself: replace, background removal, compress, reset ----

    byte[] Current() => ImagePart(Blip?.Embed?.Value) is { } part ? Read(part)
        : throw new WriterException(ErrorCode.Validation, "This picture has no image data", "Add the picture again.");

    /// <summary>A new picture from a file or data URL: the width stays, the height follows the new picture, the crop goes.</summary>
    void Replace(string src)
    {
        var (bytes, _) = ImageInfo.Load(src);
        var info = ImageInfo.Read(bytes);
        WriteCrop(default);
        Show(bytes, keepOriginal: false);
        var f = Frame;
        Frame = f with { H = Math.Max(1, (long)Math.Round(f.W * info.Height / (double)Math.Max(1, info.Width))) };
        Reframe();
    }

    /// <summary>Points the picture at a new image part with these bytes. The part shown before is kept for reset when
    /// <paramref name="keepOriginal"/> (unless an earlier original is kept already); otherwise the kept original goes too.
    /// A part is deleted once nothing refers to it.</summary>
    void Show(byte[] bytes, bool keepOriginal)
    {
        var blip = RequireBlip();
        var old = blip.Embed?.Value;
        var id = Owner.GetIdOfPart(NewPart(bytes, null));
        blip.Embed = id;
        if (old is null) return;
        var kept = KeptPart(old);
        if (!keepOriginal)
        {
            if (kept is not null) Owner.DeletePart(KeptPrefix + old);
        }
        // a package holds one relationship per target, so the kept picture's relationship is renamed to follow the picture shown
        else if (kept is not null) Owner.ChangeIdOfPart(kept, KeptPrefix + id);
        else if (ImagePart(old) is { } shown)
        {
            if (InUse(old)) NewPart(Read(shown), KeptPrefix + id); // another picture shows it too: keep a copy
            else
            {
                Owner.ChangeIdOfPart(shown, KeptPrefix + id);
                return;
            }
        }
        Drop(old);
    }

    OpenXmlPart NewPart(byte[] bytes, string? id)
    {
        var type = ImageInfo.Read(bytes).ContentType;
        OpenXmlPart part = (Owner, id) switch
        {
            (MainDocumentPart main, null) => main.AddImagePart(type),
            (MainDocumentPart main, _) => main.AddImagePart(type, id),
            (SlidePart slide, null) => slide.AddImagePart(type),
            (SlidePart slide, _) => slide.AddImagePart(type, id),
            (DrawingsPart drawings, null) => drawings.AddImagePart(type),
            (DrawingsPart drawings, _) => drawings.AddImagePart(type, id),
            _ => throw new WriterException(ErrorCode.UnsupportedKind, "Pictures here cannot be changed", "Edit pictures in the document body, on slides or on sheets."),
        };
        part.FeedData(new MemoryStream(bytes));
        return part;
    }

    OpenXmlPart? KeptPart(string? shownId) => shownId is not null && Owner.TryGetPartById(KeptPrefix + shownId, out var part) ? part : null;

    bool InUse(string relId) => Owner.RootElement?.Descendants().Any(e => e.GetAttributes().Any(a => a.NamespaceUri == RelNs && a.Value == relId)) == true;

    /// <summary>Deletes the relationship (and the part, when no other relationship keeps it) unless the markup still uses it.</summary>
    void Drop(string relId)
    {
        if (!InUse(relId) && Owner.TryGetPartById(relId, out _)) Owner.DeletePart(relId);
    }

    void Compress(int ppi)
    {
        var bytes = Current();
        var info = ImageInfo.Read(bytes);
        var crop = Crop();
        var (_, _, w, h) = Frame;
        int maxW = (int)Math.Ceiling(w / (double)Units.EmuPerInch * ppi), maxH = (int)Math.Ceiling(h / (double)Units.EmuPerInch * ppi);
        double shownW = info.Width * (1 - crop.L - crop.R), shownH = info.Height * (1 - crop.T - crop.B);
        if (crop == default && shownW <= maxW && shownH <= maxH) return; // already no bigger than it is shown
        var smaller = Vision.Compress(bytes, info.ContentType == "image/jpeg" ? ".jpg" : ".png", Math.Max(1, maxW), Math.Max(1, maxH),
            crop == default ? null : string.Join(",", new[] { crop.L, crop.T, crop.R, crop.B }.Select(v => v.ToString("0.#####", Inv))));
        if (smaller.Length >= bytes.Length) return;
        WriteCrop(default); // the cropped areas are gone from the picture itself, the frame shows the same part
        Show(smaller, keepOriginal: false);
    }

    void Reset()
    {
        SetCrop(default);
        if (Xfrm is { } xfrm)
        {
            xfrm.Rotation = null;
            xfrm.HorizontalFlip = null;
            xfrm.VerticalFlip = null;
        }
        if (Blip is { } blip)
            foreach (var effect in blip.ChildElements.Where(e => e is not A.BlipExtensionList).ToList()) effect.Remove();
        SpPr.RemoveAllChildren<A.Outline>();
        SetShadow(false);
        SetGeometry("rect");
        if (Blip?.Embed?.Value is { } shown && KeptPart(shown) is { } original)
        {
            var id = "R" + Guid.NewGuid().ToString("N")[..16];
            Owner.ChangeIdOfPart(original, id);
            Blip!.Embed = id;
            Drop(shown);
        }
        Reframe();
    }
}

/// <summary>
/// The writer-vision helper: Apple Vision background removal and ImageIO re-encoding in a small native program shipped next
/// to the engine. Its command line is in desktop/vision/README.md; a Windows build keeps the same contract.
/// </summary>
static class Vision
{
    public static byte[] Cutout(byte[] image) =>
        Run("cutout", image, ".png", [], "抠图需要 macOS 14 以上的 Apple Vision", "抠图失败");

    public static byte[] Compress(byte[] image, string ext, int maxWidth, int maxHeight, string? crop) =>
        Run("compress", image, ext, crop is null ? ["--max", $"{maxWidth}x{maxHeight}"] : ["--max", $"{maxWidth}x{maxHeight}", "--crop", crop],
            "压缩图片需要图片助手", "压缩图片失败");

    /// <summary>Next to the engine's executable, then WRITER_VISION, then on PATH.</summary>
    public static string? Find()
    {
        var name = OperatingSystem.IsWindows() ? "writer-vision.exe" : "writer-vision";
        var local = Path.Combine(AppContext.BaseDirectory, name);
        if (File.Exists(local)) return local;
        if (Environment.GetEnvironmentVariable("WRITER_VISION") is { Length: > 0 } env && File.Exists(env)) return env;
        return (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Select(dir => Path.Combine(dir, name)).FirstOrDefault(File.Exists);
    }

    static byte[] Run(string verb, byte[] image, string outExt, string[] args, string missing, string failed)
    {
        var helper = Find() ?? throw new WriterException(ErrorCode.Io, $"{missing}（找不到 writer-vision 助手）",
            "The Writer app for macOS ships it. From source: node desktop/scripts/build-vision.mjs on macOS 14+, then set WRITER_VISION to the binary it prints.");
        var dir = Directory.CreateTempSubdirectory("writer-vision-");
        try
        {
            var input = Path.Combine(dir.FullName, "in." + ImageInfo.Read(image).Extension);
            var output = Path.Combine(dir.FullName, "out" + outExt);
            File.WriteAllBytes(input, image);
            var start = new ProcessStartInfo(helper) { RedirectStandardError = true, RedirectStandardOutput = true, UseShellExecute = false };
            foreach (var arg in new[] { verb, input, output }.Concat(args)) start.ArgumentList.Add(arg);
            using var process = Process.Start(start) ?? throw new WriterException(ErrorCode.Io, $"{failed}：writer-vision 无法启动", "Check that the helper is executable.");
            var stderr = process.StandardError.ReadToEndAsync();
            _ = process.StandardOutput.ReadToEndAsync();
            if (!process.WaitForExit(TimeSpan.FromMinutes(2)))
            {
                process.Kill(entireProcessTree: true);
                throw new WriterException(ErrorCode.Io, $"{failed}：超时", "Try a smaller picture.");
            }
            if (process.ExitCode == 2) throw new WriterException(ErrorCode.Validation, $"{failed}：图片里没有找到主体", "Vision found no person, animal or object to keep; the picture stays as it was.");
            if (process.ExitCode != 0 || !File.Exists(output))
                throw new WriterException(ErrorCode.Io, $"{failed}：{stderr.Result.Trim()}", "The picture stays as it was.");
            return File.ReadAllBytes(output);
        }
        finally
        {
            try { dir.Delete(recursive: true); } catch (IOException) { }
        }
    }
}
