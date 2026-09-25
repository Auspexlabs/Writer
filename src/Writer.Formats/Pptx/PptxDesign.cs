using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>The edits that make a deck look designed rather than drawn by hand: a palette and fonts for the whole deck, a layout
/// change that moves placeholders the way PowerPoint does, shapes aligned and spaced evenly, text shrunk to fit its box.</summary>
static class PptxDesign
{
    // ---------- palette and fonts ----------

    /// <summary>The palette the deck wears: its first master's colour scheme when that is one of <see cref="PptxTemplate.Palettes"/>.</summary>
    public static string? Palette(PptxDocument doc)
    {
        var name = Theme(doc.Presentation.SlideMasterParts.FirstOrDefault())?.ThemeElements?.ColorScheme?.Name?.Value;
        return name is not null && name.StartsWith("Writer ", StringComparison.Ordinal)
            ? PptxTemplate.Palettes.FirstOrDefault(p => p.Name == name[7..])?.Key : null;
    }

    /// <summary>Every master's theme takes the palette's colours and fonts, and its background comes from the theme again; a dark
    /// palette swaps the colour map. What slides and shapes set for themselves stays theirs.</summary>
    public static void ApplyPalette(PptxDocument doc, PptxTemplate.Palette palette)
    {
        foreach (var master in doc.Presentation.SlideMasterParts)
        {
            if (Theme(master)?.ThemeElements is not { } elements || master.SlideMaster is not { } sm) continue;
            elements.ColorScheme = new A.ColorScheme(PptxTemplate.ColorSchemeXml(palette));
            elements.FontScheme = new A.FontScheme(PptxTemplate.FontSchemeXml("Writer " + palette.Name, palette.Heading, palette.Body));
            sm.ColorMap = ColorMap(palette.Dark);
            var data = sm.CommonSlideData!;
            if (data.Background?.BackgroundProperties?.GetFirstChild<A.BlipFill>() is null)
                data.Background = new P.Background(new P.BackgroundStyleReference(new A.SchemeColor { Val = A.SchemeColorValues.Background1 }) { Index = 1001U });
        }
    }

    /// <summary>The master's colour map: text on background as the theme names them, or swapped for a dark palette.</summary>
    static P.ColorMap ColorMap(bool dark) => new()
    {
        Background1 = dark ? A.ColorSchemeIndexValues.Dark1 : A.ColorSchemeIndexValues.Light1,
        Text1 = dark ? A.ColorSchemeIndexValues.Light1 : A.ColorSchemeIndexValues.Dark1,
        Background2 = dark ? A.ColorSchemeIndexValues.Dark2 : A.ColorSchemeIndexValues.Light2,
        Text2 = dark ? A.ColorSchemeIndexValues.Light2 : A.ColorSchemeIndexValues.Dark2,
        Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3,
        Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6,
        Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
    };

    /// <summary>"heading,body": the theme fonts of the deck.</summary>
    public static string? Fonts(PptxDocument doc)
    {
        var fonts = Theme(doc.Presentation.SlideMasterParts.FirstOrDefault())?.ThemeElements?.FontScheme;
        var major = fonts?.MajorFont?.LatinFont?.Typeface?.Value;
        var minor = fonts?.MinorFont?.LatinFont?.Typeface?.Value;
        return major is null && minor is null ? null : major + "," + minor;
    }

    /// <summary>Headings in the first font and everything else in the second, on every slide: the theme fonts change, and fonts set
    /// on the text itself (what a user or an earlier edit picked) are taken off so the theme's apply.</summary>
    public static void SetFonts(PptxDocument doc, string value)
    {
        var parts = value.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length is < 1 or > 2)
            throw new WriterException(ErrorCode.Validation, $"fonts: '{value}' is not valid", "Give the heading font and the body font, e.g. fonts=\"Noto Serif SC,Noto Sans SC\".");
        var (heading, body) = (parts[0], parts.Length > 1 ? parts[1] : parts[0]);
        foreach (var master in doc.Presentation.SlideMasterParts)
            if (Theme(master)?.ThemeElements is { } elements)
                elements.FontScheme = new A.FontScheme(PptxTemplate.FontSchemeXml(elements.FontScheme?.Name?.Value ?? "Writer", heading, body));
        foreach (var slide in doc.Slides)
            foreach (var rPr in slide.Slide!.Descendants().Where(e => e is A.RunProperties or A.EndParagraphRunProperties or A.DefaultRunProperties).ToList())
                foreach (var font in rPr.ChildElements.Where(e => e is A.LatinFont or A.EastAsianFont).ToList()) font.Remove();
    }

    static A.Theme? Theme(SlideMasterPart? master) => master?.ThemePart?.Theme;

    // ---------- layout ----------

    /// <summary>Which placeholders stand in for each other when a slide changes layout: titles, text and content (a subtitle
    /// counts as text), anything else by its own type.</summary>
    internal static string Family(string? type) => type switch
    {
        "title" or "ctrTitle" => "title",
        null or "body" or "obj" or "subTitle" => "body",
        _ => type,
    };

    /// <summary>A slide takes another layout as PowerPoint does it: each of its placeholders (in drawing order) becomes the first
    /// free one of the same family in the new layout and moves where that one is; one left over keeps its place when it has text
    /// and goes when it is empty; the new layout's placeholders nobody took are added empty. slide.ui's relayout does the same.</summary>
    public static void ChangeLayout(SlidePart slide, SlideLayoutPart layout)
    {
        var tree = PptxDocument.ShapeTree(slide);
        // where each placeholder is now, taken before the layout it inherits its place from changes
        var own = tree.Elements<P.Shape>()
            .Select(s => (Shape: s, Ph: PlaceholderOf(s)))
            .Where(x => x.Ph is not null && !PptxTemplate.IsFooter(x.Ph))
            .Select(x => (x.Shape, x.Ph, Box: PptxShape.Box(x.Shape, slide)))
            .ToList();
        if (slide.SlideLayoutPart is { } current && !ReferenceEquals(current, layout)) slide.DeletePart(current);
        if (slide.SlideLayoutPart is null) slide.AddPart(layout);
        var slots = PptxTemplate.LayoutPlaceholders(layout).ToList();
        var taken = new bool[slots.Count];
        foreach (var (shape, ph, box) in own)
        {
            var i = Enumerable.Range(0, slots.Count).FirstOrDefault(k => !taken[k] && Family(slots[k].Ph.Type?.InnerText) == Family(ph!.Type?.InnerText), -1);
            if (i >= 0)
            {
                taken[i] = true;
                var fresh = (P.PlaceholderShape)slots[i].Ph.CloneNode(true);
                fresh.HasCustomPrompt = null;
                shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape = fresh;
                shape.ShapeProperties?.Transform2D?.Remove();
            }
            else if (PptxText.BodyText(shape.TextBody).Trim().Length == 0) shape.Remove();
            else if (shape.ShapeProperties?.Transform2D is null && box is { } b)
                (shape.ShapeProperties ??= new P.ShapeProperties()).PrependChild(new A.Transform2D(new A.Offset { X = b.X, Y = b.Y }, new A.Extents { Cx = b.W, Cy = b.H }));
        }
        for (var k = 0; k < slots.Count; k++)
            if (!taken[k]) tree.Append(PptxTemplate.EmptyPlaceholder(PptxDocument.NextShapeId(slide), slots[k].Name, slots[k].Ph));
    }

    internal static P.PlaceholderShape? PlaceholderOf(P.Shape shape) => shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;

    // ---------- align and distribute ----------

    /// <summary>"left:shape[2],image[1]": the edge (left, center, right, top, middle, bottom) and the slide's shapes, pictures or
    /// tables by path. One lines up with the slide, several with the box around them.</summary>
    public static void Align(Node slide, long width, long height, string value)
    {
        var (edge, nodes) = Parse(slide, value, "align", "left:/slide[1]/shape[2],/slide[1]/shape[3]");
        var boxes = nodes.Select(BoxOf).ToList();
        var (x0, y0, x1, y1) = nodes.Count == 1 ? (0L, 0L, width, height) : (boxes.Min(b => b.X), boxes.Min(b => b.Y), boxes.Max(b => b.X + b.W), boxes.Max(b => b.Y + b.H));
        for (var i = 0; i < nodes.Count; i++)
        {
            var b = boxes[i];
            var (name, v) = edge switch
            {
                "left" => ("x", x0),
                "center" => ("x", (x0 + x1 - b.W) / 2),
                "right" => ("x", x1 - b.W),
                "top" => ("y", y0),
                "middle" => ("y", (y0 + y1 - b.H) / 2),
                "bottom" => ("y", y1 - b.H),
                _ => throw new WriterException(ErrorCode.Validation, $"align: '{edge}' is not an edge", "Use left, center, right, top, middle or bottom, e.g. align=\"left:/slide[1]/shape[2],/slide[1]/shape[3]\"."),
            };
            nodes[i].SetProp(name, v.ToString(CultureInfo.InvariantCulture));
        }
    }

    /// <summary>"horizontal:shape[2],shape[3],shape[4]": the first and last (by position) stay and the gaps between all of them
    /// become equal; one or two are spaced evenly across the slide.</summary>
    public static void Distribute(Node slide, long width, long height, string value)
    {
        var (axis, nodes) = Parse(slide, value, "distribute", "horizontal:/slide[1]/shape[2],/slide[1]/shape[3],/slide[1]/shape[4]");
        if (axis is not ("horizontal" or "vertical"))
            throw new WriterException(ErrorCode.Validation, $"distribute: '{axis}' is not a direction", "Use horizontal or vertical, e.g. distribute=\"horizontal:/slide[1]/shape[2],/slide[1]/shape[3],/slide[1]/shape[4]\".");
        var h = axis == "horizontal";
        var items = nodes.Select(n => (Node: n, Box: BoxOf(n))).OrderBy(x => h ? x.Box.X : x.Box.Y).ToList();
        long Start((long X, long Y, long W, long H) b) => h ? b.X : b.Y;
        long Size((long X, long Y, long W, long H) b) => h ? b.W : b.H;
        var total = items.Sum(x => Size(x.Box));
        long from, gap;
        if (items.Count >= 3)
        {
            from = Start(items[0].Box);
            var to = items.Max(x => Start(x.Box) + Size(x.Box));
            gap = (to - from - total) / (items.Count - 1);
        }
        else
        {
            var span = h ? width : height;
            gap = (span - total) / (items.Count + 1);
            from = gap;
        }
        foreach (var (node, box) in items)
        {
            node.SetProp(h ? "x" : "y", from.ToString(CultureInfo.InvariantCulture));
            from += Size(box) + gap;
        }
    }

    static (string Head, List<Node> Nodes) Parse(Node slide, string value, string prop, string example)
    {
        var colon = value.IndexOf(':');
        var hint = $"Example: {prop}=\"{example}\" (paths from 'view <file> outline').";
        if (colon < 0) throw new WriterException(ErrorCode.Validation, $"{prop}: '{value}' names no shapes", hint);
        var nodes = new List<Node>();
        var root = slide.Parent ?? slide;
        foreach (var item in value[(colon + 1)..].Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
        {
            // a full path from the outline, or one relative to the slide (shape[2])
            var node = item.StartsWith("/slide", StringComparison.Ordinal) ? PathResolver.Single(root, item) : PathResolver.Single(slide, "/" + item.TrimStart('/'));
            if (!ReferenceEquals(node.Parent?.Anchor, slide.Anchor) || node.Kind is not ("shape" or "image" or "table" or "connector" or "group"))
                throw new WriterException(ErrorCode.Validation, $"{prop}: {item} is not a shape, picture or table on {slide.Path}", "Name this slide's shapes, pictures or tables.");
            if (!nodes.Any(n => ReferenceEquals(n.Anchor, node.Anchor))) nodes.Add(node);
        }
        if (nodes.Count == 0) throw new WriterException(ErrorCode.Validation, $"{prop}: '{value}' names no shapes", hint);
        return (value[..colon].Trim().ToLowerInvariant(), nodes);
    }

    static (long X, long Y, long W, long H) BoxOf(Node node)
    {
        var p = node.GetProps();
        long L(string k) => long.TryParse(p.GetValueOrDefault(k), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var v) ? v : 0;
        return (L("x"), L("y"), L("w"), L("h"));
    }

    // ---------- text that fits ----------

    /// <summary>How tall a shape's text is, in points, at scale times its sizes (sizes without a run value are
    /// <paramref name="size"/>): lines estimated from character widths (CJK one em, Latin about half) against the box width
    /// less its insets and each paragraph's indent.</summary>
    // ponytail: an estimate without font metrics — within a line or so for body text; measure with the real fonts if it ever misleads
    static double TextHeight(P.Shape shape, double widthPt, double size, double scale)
    {
        var bodyPr = shape.TextBody?.BodyProperties;
        double Inset(Int32Value? v, int fallback) => (v?.Value ?? fallback) / 12700.0;
        var inner = widthPt - Inset(bodyPr?.LeftInset, 91440) - Inset(bodyPr?.RightInset, 91440);
        var height = Inset(bodyPr?.TopInset, 45720) + Inset(bodyPr?.BottomInset, 45720);
        foreach (var p in PptxText.Paragraphs(shape.TextBody))
        {
            var runSize = p.Elements<A.Run>().Select(r => r.RunProperties?.FontSize?.Value).Where(s => s is not null).Select(s => s!.Value / 100.0).DefaultIfEmpty(size).Max() * scale;
            var indent = (p.ParagraphProperties?.LeftMargin?.Value ?? (PptxText.ListOf(p.ParagraphProperties) is null ? 0 : 342900)) / 12700.0;
            var lines = 0;
            foreach (var line in PptxText.ParagraphText(p).Split('\n'))
                lines += Math.Max(1, (int)Math.Ceiling(line.Sum(Em) * runSize / Math.Max(1, inner - indent)));
            height += lines * runSize * 1.2 + runSize * 0.3;
        }
        return height;
    }

    static double Em(char c) => c switch
    {
        >= '⺀' and <= '鿿' or >= '豈' and <= '﫿' or >= '＀' and <= '￯' or >= '　' and <= '〿' => 1.0,
        ' ' => 0.28,
        >= 'A' and <= 'Z' => 0.64,
        >= '0' and <= '9' => 0.56,
        >= 'a' and <= 'z' => 0.52,
        _ when char.IsPunctuation(c) => 0.32,
        _ => 0.56,
    };

    /// <summary>True when the text needs more height than its box has (text that does not wrap is left out).</summary>
    public static bool Overflows(P.Shape shape, (long X, long Y, long W, long H)? box, double size) =>
        box is { } b && b.W > 0 && b.H > 0 && shape.TextBody?.BodyProperties?.Wrap?.Value != A.TextWrappingValues.None
        && PptxText.BodyText(shape.TextBody).Trim().Length > 0 && TextHeight(shape, b.W / 12700.0, size, 1) > b.H / 12700.0 + 2;

    /// <summary>fit=shrink: every run's size comes down by the same factor, in 5% steps to half, until the text fits its box.</summary>
    public static void Shrink(P.Shape shape, (long X, long Y, long W, long H)? box, double size)
    {
        if (box is not { } b || !Overflows(shape, box, size)) return;
        var scale = 1.0;
        while (scale > 0.5 && TextHeight(shape, b.W / 12700.0, size, scale) > b.H / 12700.0) scale -= 0.05;
        foreach (var p in PptxText.Paragraphs(shape.TextBody))
            foreach (var rPr in p.Elements<A.Run>().Select(r => r.RunProperties ??= new A.RunProperties()).Cast<OpenXmlElement>().Append(p.GetFirstChild<A.EndParagraphRunProperties>()).OfType<A.TextCharacterPropertiesType>())
                rPr.FontSize = (int)Math.Round(Math.Max(8, (rPr.FontSize?.Value / 100.0 ?? size) * scale)) * 100;
    }
}
