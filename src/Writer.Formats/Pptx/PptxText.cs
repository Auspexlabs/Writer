using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace Writer.Formats.Pptx;

/// <summary>DrawingML text: reading and writing paragraphs and runs inside shapes and table cells.</summary>
static class PptxText
{
    public static IEnumerable<A.Paragraph> Paragraphs(OpenXmlCompositeElement? body) => body?.Elements<A.Paragraph>() ?? [];

    public static string ParagraphText(A.Paragraph p)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var child in p.ChildElements)
        {
            switch (child)
            {
                case A.Run r: sb.Append(r.Text?.Text); break;
                case A.Field f: sb.Append(f.Text?.Text); break;
                case A.Break: sb.Append('\n'); break;
            }
        }
        return sb.ToString();
    }

    public static string BodyText(OpenXmlCompositeElement? body) => string.Join("\n", Paragraphs(body).Select(ParagraphText));

    /// <summary>Replaces every paragraph of a text body. Line breaks in the runs start new paragraphs.
    /// The first paragraph's properties and the first run's formatting carry over.</summary>
    public static void SetBody(OpenXmlCompositeElement body, SlidePart slide, IEnumerable<RunSpec> specs)
    {
        var first = body.Elements<A.Paragraph>().FirstOrDefault();
        var pPr = first?.ParagraphProperties?.CloneNode(true) as A.ParagraphProperties;
        var rPr = first?.Elements<A.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties;
        foreach (var p in body.Elements<A.Paragraph>().ToList()) p.Remove();
        var lines = SplitLines(specs);
        if (lines.Count == 0) lines.Add([]);
        foreach (var line in lines)
        {
            var p = new A.Paragraph();
            if (pPr is not null) p.Append(pPr.CloneNode(true));
            foreach (var spec in line) p.Append(MakeRun(spec, rPr, slide));
            body.Append(p);
        }
    }

    /// <summary>Replaces the runs of one paragraph; line breaks become a:br.</summary>
    public static void SetParagraph(A.Paragraph p, SlidePart slide, IEnumerable<RunSpec> specs)
    {
        var rPr = p.Elements<A.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties;
        foreach (var e in p.ChildElements.Where(e => e is A.Run or A.Break or A.Field).ToList()) e.Remove();
        var end = p.GetFirstChild<A.EndParagraphRunProperties>();
        foreach (var spec in specs)
        {
            var lines = spec.Text.ReplaceLineEndings("\n").Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                if (i > 0) Insert(p, new A.Break(), end);
                if (lines[i].Length > 0) Insert(p, MakeRun(spec with { Text = lines[i] }, rPr, slide), end);
            }
        }
    }

    static void Insert(A.Paragraph p, OpenXmlElement element, A.EndParagraphRunProperties? end)
    {
        if (end is null) p.Append(element);
        else p.InsertBefore(element, end);
    }

    static List<List<RunSpec>> SplitLines(IEnumerable<RunSpec> specs)
    {
        var lines = new List<List<RunSpec>> { new() };
        foreach (var spec in specs)
        {
            var parts = spec.Text.ReplaceLineEndings("\n").Split('\n');
            for (var i = 0; i < parts.Length; i++)
            {
                if (i > 0) lines.Add([]);
                if (parts[i].Length > 0) lines[^1].Add(spec with { Text = parts[i] });
            }
        }
        return lines;
    }

    public static A.Run MakeRun(RunSpec spec, A.RunProperties? baseProperties, SlidePart slide)
    {
        spec = spec.Flattened();
        var rPr = baseProperties?.CloneNode(true) as A.RunProperties ?? new A.RunProperties { Language = "en-US" };
        if (spec.Bold) rPr.Bold = true;
        if (spec.Italic) rPr.Italic = true;
        if (spec.Strike) rPr.Strike = A.TextStrikeValues.SingleStrike;
        if (spec.Code) SetFont(rPr, "Consolas");
        if (spec.Underline) rPr.Underline = A.TextUnderlineValues.Single;
        if (spec.Color is not null) SetColor(rPr, spec.Color);
        if (spec.Size is not null) rPr.FontSize = Hundredths(spec.Size);
        if (spec.Font is not null) SetFont(rPr, spec.Font);
        if (spec.Highlight is not null) SetHighlight(rPr, spec.Highlight);
        if (spec.Link is not null) SetLink(rPr, slide, spec.Link);
        return new A.Run(rPr, new A.Text(spec.Text));
    }

    public static string? Highlight(A.RunProperties? rPr) => rPr?.GetFirstChild<A.Highlight>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();

    public static void SetHighlight(A.RunProperties rPr, string value)
    {
        rPr.RemoveAllChildren<A.Highlight>();
        if (value == "none") return;
        InsertBeforeAny(rPr, new A.Highlight(new A.RgbColorModelHex { Val = value }), e => e is A.UnderlineFollowsText or A.Underline or A.UnderlineFillText
            or A.UnderlineFill or A.LatinFont or A.EastAsianFont or A.ComplexScriptFont or A.SymbolFont or A.HyperlinkOnClick or A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList);
    }

    public static string? Color(A.RunProperties? rPr) => rPr?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();

    public static void SetColor(OpenXmlCompositeElement properties, string value)
    {
        foreach (var fill in properties.ChildElements.Where(IsFill).ToList()) fill.Remove();
        if (value == "none") return;
        var fillElement = new A.SolidFill(new A.RgbColorModelHex { Val = value });
        if (properties.GetFirstChild<A.Outline>() is { } line && properties is A.RunProperties) properties.InsertAfter(fillElement, line);
        else if (properties is A.RunProperties or A.EndParagraphRunProperties) properties.PrependChild(fillElement);
        else InsertBeforeAny(properties, fillElement, e => e is A.Outline or A.EffectList or A.EffectDag or A.Scene3DType or A.Shape3DType or A.ExtensionList);
    }

    static bool IsFill(OpenXmlElement e) => e is A.SolidFill or A.NoFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill;

    public static void SetFont(A.RunProperties rPr, string typeface)
    {
        foreach (var f in rPr.ChildElements.Where(e => e is A.LatinFont or A.EastAsianFont).ToList()) f.Remove();
        var latin = new A.LatinFont { Typeface = typeface };
        InsertBeforeAny(rPr, latin, e => e is A.ComplexScriptFont or A.SymbolFont or A.HyperlinkOnClick or A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList);
        rPr.InsertAfter(new A.EastAsianFont { Typeface = typeface }, latin);
    }

    public static void SetLink(A.RunProperties rPr, SlidePart slide, string target)
    {
        rPr.RemoveAllChildren<A.HyperlinkOnClick>();
        if (target.Length == 0) return;
        Uri uri;
        try { uri = new Uri(target, UriKind.RelativeOrAbsolute); }
        catch (UriFormatException) { throw new WriterException(ErrorCode.Validation, $"'{target}' is not a valid link", "Use a URL like https://example.com."); }
        var relationship = slide.AddHyperlinkRelationship(uri, true);
        InsertBeforeAny(rPr, new A.HyperlinkOnClick { Id = relationship.Id }, e => e is A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList);
    }

    public static string? Link(A.RunProperties? rPr, SlidePart slide)
    {
        var id = rPr?.GetFirstChild<A.HyperlinkOnClick>()?.Id?.Value;
        return id is null ? null : slide.HyperlinkRelationships.FirstOrDefault(r => r.Id == id)?.Uri.OriginalString;
    }

    /// <summary>Inserts before the first child matching the predicate, or appends.</summary>
    public static void InsertBeforeAny(OpenXmlCompositeElement parent, OpenXmlElement element, Func<OpenXmlElement, bool> after)
    {
        var anchor = parent.ChildElements.FirstOrDefault(after);
        if (anchor is null) parent.Append(element);
        else parent.InsertBefore(element, anchor);
    }

    public static string? AlignOf(A.ParagraphProperties? pPr) => pPr?.Alignment?.InnerText switch
    {
        "l" => "left",
        "ctr" => "center",
        "r" => "right",
        "just" or "justLow" or "dist" => "justify",
        _ => null,
    };

    public static A.TextAlignmentTypeValues AlignValue(string align) => align switch
    {
        "center" => A.TextAlignmentTypeValues.Center,
        "right" => A.TextAlignmentTypeValues.Right,
        "justify" => A.TextAlignmentTypeValues.Justified,
        _ => A.TextAlignmentTypeValues.Left,
    };

    public static string? ListOf(A.ParagraphProperties? pPr)
    {
        if (pPr is null) return null;
        if (pPr.GetFirstChild<A.CharacterBullet>() is not null || pPr.GetFirstChild<A.PictureBullet>() is not null) return "bullet";
        if (pPr.GetFirstChild<A.AutoNumberedBullet>() is not null) return "number";
        return null;
    }

    public static void SetList(A.ParagraphProperties pPr, string kind)
    {
        foreach (var e in pPr.ChildElements.Where(e => e is A.BulletFont or A.CharacterBullet or A.AutoNumberedBullet or A.NoBullet or A.PictureBullet or A.BulletFontText).ToList()) e.Remove();
        var level = pPr.Level?.Value ?? 0;
        if (kind == "none")
        {
            pPr.LeftMargin = 0;
            pPr.Indent = 0;
            InsertBeforeAny(pPr, new A.NoBullet(), e => e is A.TabStopList or A.DefaultRunProperties or A.ExtensionList);
            return;
        }
        pPr.LeftMargin = 342900 + 457200 * level;
        pPr.Indent = -342900;
        if (kind == "bullet")
        {
            var font = new A.BulletFont { Typeface = "Arial" };
            InsertBeforeAny(pPr, font, e => e is A.TabStopList or A.DefaultRunProperties or A.ExtensionList);
            pPr.InsertAfter(new A.CharacterBullet { Char = "•" }, font);
        }
        else InsertBeforeAny(pPr, new A.AutoNumberedBullet { Type = A.TextAutoNumberSchemeValues.ArabicPeriod }, e => e is A.TabStopList or A.DefaultRunProperties or A.ExtensionList);
    }

    public static string Points(int? hundredths) => ((hundredths ?? 0) / 100.0).ToString("0.##", CultureInfo.InvariantCulture);

    public static int Hundredths(string points) => (int)Math.Round(double.Parse(points, CultureInfo.InvariantCulture) * 100);
}
