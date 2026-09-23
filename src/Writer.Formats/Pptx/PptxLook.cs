using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>What a slide element looks like once PowerPoint's inheritance is applied: theme colours through the colour map,
/// colour modifiers (lumMod, lumOff, tint, shade…), shape style references, placeholder inheritance from the layout and
/// master, the master's text styles, the presentation's default text style and theme fonts. Explicit values stay in a node's
/// props; this supplies the rest as "computed" so a UI can draw a deck made from a real PowerPoint template.
/// Works on element local names and attributes rather than SDK types, so unknown or newer markup is simply skipped.</summary>
sealed class PptxLook
{
    readonly OpenXmlElement? _scheme, _fonts, _format, _defaultText;
    readonly SlideMasterPart? _master;
    readonly SlideLayoutPart? _layout;
    readonly Dictionary<string, string> _map = new()
    {
        ["bg1"] = "lt1", ["tx1"] = "dk1", ["bg2"] = "lt2", ["tx2"] = "dk2",
    };

    PptxLook(OpenXmlPart part, PresentationPart? presentation)
    {
        var slide = part as SlidePart;
        _layout = slide?.SlideLayoutPart ?? part as SlideLayoutPart;
        _master = _layout?.SlideMasterPart ?? part as SlideMasterPart;
        var theme = _master?.ThemePart?.Theme;
        var elements = Child(theme, "themeElements");
        _scheme = Child(elements, "clrScheme");
        _fonts = Child(elements, "fontScheme");
        _format = Child(elements, "fmtScheme");
        _defaultText = Child(presentation?.Presentation, "defaultTextStyle");
        // The colour map: the master's p:clrMap, overridden by the layout's and then the slide's a:overrideClrMapping.
        Remap(Child(_master?.SlideMaster, "clrMap"));
        Remap(Child(Child(_layout?.SlideLayout, "clrMapOvr"), "overrideClrMapping"));
        Remap(Child(Child(slide?.Slide, "clrMapOvr"), "overrideClrMapping"));
    }

    public static PptxLook For(OpenXmlPart part, PresentationPart? presentation) => new(part, presentation);

    void Remap(OpenXmlElement? map)
    {
        if (map is null) return;
        foreach (var attribute in map.GetAttributes()) _map[attribute.LocalName] = attribute.Value ?? attribute.LocalName;
    }

    static OpenXmlElement? Child(OpenXmlElement? parent, string localName) =>
        parent?.ChildElements.FirstOrDefault(e => e.LocalName == localName);

    static string? Attr(OpenXmlElement? element, string localName) =>
        element?.GetAttributes().FirstOrDefault(a => a.LocalName == localName && string.IsNullOrEmpty(a.NamespaceUri)).Value;

    static int? IntAttr(OpenXmlElement? element, string localName) =>
        int.TryParse(Attr(element, localName), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var v) ? v : null;

    // ---------- colours ----------

    static readonly HashSet<string> ColorElements = ["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"];

    static readonly Dictionary<string, string> Presets = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000", ["white"] = "FFFFFF", ["red"] = "FF0000", ["green"] = "008000", ["blue"] = "0000FF", ["yellow"] = "FFFF00",
        ["gray"] = "808080", ["grey"] = "808080", ["darkGray"] = "A9A9A9", ["lightGray"] = "D3D3D3", ["orange"] = "FFA500", ["purple"] = "800080",
        ["cyan"] = "00FFFF", ["magenta"] = "FF00FF", ["navy"] = "000080", ["maroon"] = "800000", ["silver"] = "C0C0C0", ["teal"] = "008080",
    };

    /// <summary>RRGGBB for a colour element (srgbClr, schemeClr, sysClr, prstClr, scrgbClr, hslClr) with its modifiers applied.
    /// <paramref name="placeholder"/> is what schemeClr val="phClr" stands for inside a theme style.</summary>
    public string? Color(OpenXmlElement? color, string? placeholder = null)
    {
        if (color is null) return null;
        string? rgb = color.LocalName switch
        {
            "srgbClr" => Attr(color, "val"),
            "sysClr" => Attr(color, "lastClr") ?? (Attr(color, "val") == "window" ? "FFFFFF" : "000000"),
            "prstClr" => Presets.GetValueOrDefault(Attr(color, "val") ?? ""),
            "schemeClr" => Attr(color, "val") is "phClr" ? placeholder : Scheme(Attr(color, "val")),
            "scrgbClr" => FromLinear(IntAttr(color, "r") ?? 0, IntAttr(color, "g") ?? 0, IntAttr(color, "b") ?? 0),
            "hslClr" => FromHsl((IntAttr(color, "hue") ?? 0) / 60000.0 / 360, (IntAttr(color, "sat") ?? 0) / 100000.0, (IntAttr(color, "lum") ?? 0) / 100000.0),
            _ => null,
        };
        if (rgb is null || rgb.Length != 6) return rgb;
        return Modify(rgb.ToUpperInvariant(), color);
    }

    string? Scheme(string? name)
    {
        if (name is null || _scheme is null) return null;
        var slot = _map.GetValueOrDefault(name, name);
        var entry = Child(_scheme, slot);
        return Color(entry?.ChildElements.FirstOrDefault(e => ColorElements.Contains(e.LocalName)));
    }

    /// <summary>The first colour element among a container's children (solidFill, fgClr, a style reference…).</summary>
    public string? FirstColor(OpenXmlElement? container, string? placeholder = null) =>
        Color(container?.ChildElements.FirstOrDefault(e => ColorElements.Contains(e.LocalName)), placeholder);

    /// <summary>The fill a properties element (spPr, bgPr, rPr, ln) declares: a colour, "none", or null when it declares none.
    /// Gradients report their first stop and patterns their foreground; pictures report null.</summary>
    public string? Fill(OpenXmlElement? properties, string? placeholder = null)
    {
        if (properties is null) return null;
        foreach (var child in properties.ChildElements)
        {
            var fill = FillElement(child, placeholder, out var declared);
            if (declared) return fill;
        }
        return null;
    }

    string? FillElement(OpenXmlElement element, string? placeholder, out bool declared)
    {
        declared = true;
        switch (element.LocalName)
        {
            case "noFill": return "none";
            case "solidFill": return FirstColor(element, placeholder);
            case "gradFill":
                var stop = Child(Child(element, "gsLst"), "gs");
                return FirstColor(stop, placeholder);
            case "pattFill": return FirstColor(Child(element, "fgClr"), placeholder);
            case "blipFill": return null;
            case "grpFill": return null;
            default:
                declared = false;
                return null;
        }
    }

    /// <summary>A shape style reference (fillRef, lnRef, fontRef, bgRef): the theme's style at idx with phClr replaced by the reference's own colour.</summary>
    public string? StyleRef(OpenXmlElement? reference)
    {
        if (reference is null) return null;
        var own = FirstColor(reference);
        if (reference.LocalName == "fontRef") return own;
        var index = IntAttr(reference, "idx") ?? 0;
        if (index == 0) return "none";
        OpenXmlElement? list = reference.LocalName switch
        {
            "lnRef" => Child(_format, "lnStyleLst"),
            _ when index >= 1001 => Child(_format, "bgFillStyleLst"),
            _ => Child(_format, "fillStyleLst"),
        };
        var position = (index >= 1001 ? index - 1001 : index - 1);
        var style = list?.ChildElements.ElementAtOrDefault(position);
        if (style is null) return own;
        if (reference.LocalName == "lnRef") return Fill(style, own) ?? own;
        var fill = FillElement(style, own, out var declared);
        return declared ? fill ?? own : own;
    }

    static string Modify(string rgb, OpenXmlElement color)
    {
        var modifiers = color.ChildElements.Where(e => e.LocalName is "lumMod" or "lumOff" or "tint" or "shade" or "satMod" or "satOff").ToList();
        if (modifiers.Count == 0) return rgb;
        double r = Convert.ToInt32(rgb[..2], 16) / 255.0, g = Convert.ToInt32(rgb[2..4], 16) / 255.0, b = Convert.ToInt32(rgb[4..], 16) / 255.0;
        foreach (var m in modifiers)
        {
            var value = (IntAttr(m, "val") ?? 100000) / 100000.0;
            switch (m.LocalName)
            {
                case "tint":
                    (r, g, b) = (Encode(Decode(r) * value + (1 - value)), Encode(Decode(g) * value + (1 - value)), Encode(Decode(b) * value + (1 - value)));
                    break;
                case "shade":
                    (r, g, b) = (Encode(Decode(r) * value), Encode(Decode(g) * value), Encode(Decode(b) * value));
                    break;
                default:
                    var (h, s, l) = ToHsl(r, g, b);
                    switch (m.LocalName)
                    {
                        case "lumMod": l *= value; break;
                        case "lumOff": l += value; break;
                        case "satMod": s *= value; break;
                        case "satOff": s += value; break;
                    }
                    (r, g, b) = HslToRgb(h, Math.Clamp(s, 0, 1), Math.Clamp(l, 0, 1));
                    break;
            }
        }
        return Hex(r, g, b);
    }

    static double Decode(double c) => c <= 0.04045 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
    static double Encode(double c) => Math.Clamp(c <= 0.0031308 ? c * 12.92 : 1.055 * Math.Pow(c, 1 / 2.4) - 0.055, 0, 1);

    static string FromLinear(int r, int g, int b) => Hex(Encode(r / 100000.0), Encode(g / 100000.0), Encode(b / 100000.0));

    static string FromHsl(double h, double s, double l)
    {
        var (r, g, b) = HslToRgb(h, s, l);
        return Hex(r, g, b);
    }

    static string Hex(double r, double g, double b) =>
        string.Concat(new[] { r, g, b }.Select(c => ((int)Math.Round(Math.Clamp(c, 0, 1) * 255)).ToString("X2", CultureInfo.InvariantCulture)));

    static (double H, double S, double L) ToHsl(double r, double g, double b)
    {
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var l = (max + min) / 2;
        if (max == min) return (0, 0, l);
        var d = max - min;
        var s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        double h;
        if (max == r) h = (g - b) / d + (g < b ? 6 : 0);
        else if (max == g) h = (b - r) / d + 2;
        else h = (r - g) / d + 4;
        return (h / 6, s, l);
    }

    static (double R, double G, double B) HslToRgb(double h, double s, double l)
    {
        if (s == 0) return (l, l, l);
        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;
        return (Hue(p, q, h + 1.0 / 3), Hue(p, q, h), Hue(p, q, h - 1.0 / 3));
    }

    static double Hue(double p, double q, double t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6) return p + (q - p) * 6 * t;
        if (t < 0.5) return q;
        if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
        return p;
    }

    // ---------- fonts ----------

    /// <summary>A typeface, with theme references (+mj-lt, +mn-ea…) replaced by the theme's major/minor fonts. Null when empty.</summary>
    public string? Font(string? typeface)
    {
        if (string.IsNullOrEmpty(typeface)) return null;
        if (!typeface.StartsWith('+')) return typeface;
        var family = typeface.StartsWith("+mj", StringComparison.Ordinal) ? "majorFont" : "minorFont";
        var script = typeface.EndsWith("-ea", StringComparison.Ordinal) ? "ea" : typeface.EndsWith("-cs", StringComparison.Ordinal) ? "cs" : "latin";
        var resolved = Attr(Child(Child(_fonts, family), script), "typeface");
        return string.IsNullOrEmpty(resolved) ? null : resolved;
    }

    // ---------- shapes ----------

    /// <summary>The appearance of a shape or connector that its own properties do not state: fill, line, text colour, fonts, size.</summary>
    public Dictionary<string, string> Shape(OpenXmlElement shape, IReadOnlyDictionary<string, string> props)
    {
        var result = new Dictionary<string, string>();
        var placeholder = (shape as P.Shape)?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
        var (layoutShape, masterShape) = placeholder is null ? (null, null) : Inherited(placeholder);
        var style = Child(shape, "style");

        if (!props.ContainsKey("fill"))
        {
            var fill = Fill(Child(shape, "spPr")) ?? Fill(Child(layoutShape, "spPr")) ?? Fill(Child(masterShape, "spPr"))
                ?? StyleRef(Child(style, "fillRef"));
            if (fill is not null) result["fill"] = fill;
        }
        if (!props.ContainsKey("line"))
        {
            var line = Fill(Child(Child(shape, "spPr"), "ln")) ?? Fill(Child(Child(layoutShape, "spPr"), "ln"))
                ?? Fill(Child(Child(masterShape, "spPr"), "ln")) ?? StyleRef(Child(style, "lnRef"));
            if (line is not null) result["line"] = line;
        }

        var level = IntAttr(Child(Child(shape, "txBody")?.ChildElements.FirstOrDefault(e => e.LocalName == "p"), "pPr"), "lvl") ?? 0;
        var sources = new List<OpenXmlElement?>
        {
            FirstRunProperties(shape),
            LevelDefaults(Child(Child(shape, "txBody"), "lstStyle"), level),
        };
        var fontRef = Child(style, "fontRef");
        if (placeholder is not null)
        {
            sources.Add(LevelDefaults(Child(Child(layoutShape, "txBody"), "lstStyle"), level));
            sources.Add(LevelDefaults(Child(Child(masterShape, "txBody"), "lstStyle"), level));
            sources.Add(LevelDefaults(MasterTextStyle(placeholder), level));
        }
        else
        {
            sources.Add(LevelDefaults(_defaultText, level));
            sources.Add(LevelDefaults(Child(Child(_master?.SlideMaster, "txStyles"), "otherStyle"), level));
        }

        if (!props.ContainsKey("color"))
        {
            string? color = null;
            foreach (var rPr in sources.Take(2)) if ((color = Fill(rPr)) is not null) break;
            color ??= FirstColor(fontRef);
            if (color is null) foreach (var rPr in sources.Skip(2)) if ((color = Fill(rPr)) is not null) break;
            if (color is not null && color != "none") result["color"] = color;
        }
        if (!props.ContainsKey("font"))
        {
            var font = FirstFont(sources.Take(2), "latin") ?? FontRef(fontRef, "lt") ?? FirstFont(sources.Skip(2), "latin");
            if (font is not null) result["font"] = font;
        }
        var ea = FirstFont(sources.Take(2), "ea") ?? FontRef(fontRef, "ea") ?? FirstFont(sources.Skip(2), "ea");
        if (ea is not null && ea != (props.GetValueOrDefault("font") ?? result.GetValueOrDefault("font"))) result["fontEa"] = ea;
        if (!props.ContainsKey("size"))
        {
            var size = sources.Select(s => IntAttr(s, "sz")).FirstOrDefault(s => s is not null);
            if (size is not null) result["size"] = PptxText.Points(size);
        }
        // bold as the list and master styles set it; a run's own bold is already in the paragraphs' html
        if (sources.Skip(1).Select(s => Attr(s, "b")).FirstOrDefault(b => b is not null) is { } bold)
            result["bold"] = bold is "1" or "true" ? "true" : "false";
        return result;
    }

    string? FirstFont(IEnumerable<OpenXmlElement?> sources, string script)
    {
        foreach (var rPr in sources)
            if (Font(Attr(Child(rPr, script), "typeface")) is { } font) return font;
        return null;
    }

    string? FontRef(OpenXmlElement? fontRef, string script)
    {
        var idx = Attr(fontRef, "idx");
        if (idx is not ("major" or "minor")) return null;
        return Font((idx == "major" ? "+mj-" : "+mn-") + script);
    }

    static OpenXmlElement? FirstRunProperties(OpenXmlElement shape) =>
        Child(shape, "txBody")?.ChildElements.Where(e => e.LocalName == "p")
            .SelectMany(p => p.ChildElements.Where(e => e.LocalName == "r")).Select(r => Child(r, "rPr")).FirstOrDefault(r => r is not null);

    static OpenXmlElement? LevelDefaults(OpenXmlElement? listStyle, int level) =>
        Child(Child(listStyle, $"lvl{Math.Clamp(level, 0, 8) + 1}pPr"), "defRPr");

    OpenXmlElement? MasterTextStyle(P.PlaceholderShape placeholder)
    {
        var styles = Child(_master?.SlideMaster, "txStyles");
        var type = placeholder.Type?.InnerText;
        return type is "title" or "ctrTitle" ? Child(styles, "titleStyle")
            : type is "dt" or "ftr" or "sldNum" ? Child(styles, "otherStyle")
            : Child(styles, "bodyStyle");
    }

    /// <summary>The layout and master placeholders a slide placeholder inherits from: matched by idx, then by type.</summary>
    (OpenXmlElement? Layout, OpenXmlElement? Master) Inherited(P.PlaceholderShape placeholder)
    {
        var type = placeholder.Type?.InnerText ?? "body";
        var idx = placeholder.Index?.Value;
        var fromLayout = Match(_layout?.SlideLayout?.CommonSlideData?.ShapeTree, type, idx);
        var layoutType = (fromLayout as P.Shape)?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.InnerText ?? type;
        var fromMaster = Match(_master?.SlideMaster?.CommonSlideData?.ShapeTree, layoutType, null);
        return (fromLayout, fromMaster);
    }

    static OpenXmlElement? Match(OpenXmlElement? tree, string type, uint? idx)
    {
        if (tree is null) return null;
        var shapes = tree.Descendants<P.Shape>().Select(s => (Shape: s, Ph: s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape))
            .Where(x => x.Ph is not null).ToList();
        if (idx is not null && shapes.FirstOrDefault(x => x.Ph!.Index?.Value == idx).Shape is { } byIndex) return byIndex;
        static string Family(string t) => t is "ctrTitle" ? "title" : t is "subTitle" or "obj" ? "body" : t;
        return shapes.FirstOrDefault(x => (x.Ph!.Type?.InnerText ?? "body") == type).Shape
            ?? shapes.FirstOrDefault(x => Family(x.Ph!.Type?.InnerText ?? "body") == Family(type)).Shape;
    }

    // ---------- backgrounds ----------

    /// <summary>The effective background colour of a slide: its own p:bg, else its layout's, else its master's — solid colours,
    /// theme colours and background style references. Null for picture backgrounds and when none is declared.</summary>
    public string? Background(P.Background? background)
    {
        if (background is null) return null;
        if (Child(background, "bgPr") is { } properties) return Fill(properties);
        return StyleRef(Child(background, "bgRef"));
    }
}
