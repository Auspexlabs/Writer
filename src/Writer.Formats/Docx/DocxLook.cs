using System.Globalization;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>How a paragraph looks where the file leaves it to styles, as Word cascades it: the document defaults (docDefaults), then the
/// paragraph's style and the styles it is based on. Paragraph settings in DocxParaFormat's words (spacing, indents), the text as font,
/// fontEa, size, bold, italic and color, theme fonts (minorHAnsi, majorEastAsia…) as the theme's typefaces.</summary>
static class DocxLook
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    static readonly string[] ParagraphKeys = ["lineSpacing", "spaceBefore", "spaceAfter", "indentLeft", "indentRight", "indentFirst"];

    /// <summary>What a key is when nothing in the cascade sets it (Word's own defaults).</summary>
    static readonly Dictionary<string, string> Unset = new()
    {
        ["lineSpacing"] = "1", ["spaceBefore"] = "0pt", ["spaceAfter"] = "0pt", ["indentLeft"] = "0", ["indentRight"] = "0", ["indentFirst"] = "0",
        ["size"] = "10", ["bold"] = "false", ["italic"] = "false",
    };

    /// <summary>The look of a paragraph in this style; the default paragraph style's (the document's own look) for null.</summary>
    public static Dictionary<string, string> Of(DocxDocument doc, string? styleId)
    {
        var look = new Dictionary<string, string>();
        var defaults = doc.Main.StyleDefinitionsPart?.Styles?.DocDefaults;
        Paragraph(defaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle, look);
        Run(doc, defaults?.RunPropertiesDefault?.RunPropertiesBaseStyle, look);
        var chain = new List<W.Style>();
        for (var s = doc.Styles.Find(styleId) ?? doc.Styles.DefaultParagraphStyle(); s is not null && chain.Count < 16 && !chain.Contains(s); s = doc.Styles.Find(s.BasedOn?.Val?.Value))
            chain.Add(s);
        for (var i = chain.Count - 1; i >= 0; i--)
        {
            Paragraph(chain[i].StyleParagraphProperties, look);
            Run(doc, chain[i].StyleRunProperties, look);
        }
        return look;
    }

    /// <summary>What a paragraph's style gives beyond the document's own look (Of null), for the keys the paragraph does not set itself.</summary>
    public static IEnumerable<KeyValuePair<string, string>> Differences(DocxDocument doc, string? styleId, IReadOnlyDictionary<string, string> own)
    {
        var look = Of(doc, styleId);
        var document = Of(doc, null);
        foreach (var key in look.Keys.Union(document.Keys).Where(k => !own.ContainsKey(k)))
        {
            var value = look.GetValueOrDefault(key) ?? Unset.GetValueOrDefault(key);
            if (value is not null && value != (document.GetValueOrDefault(key) ?? Unset.GetValueOrDefault(key))) yield return new(key, value);
        }
    }

    static void Paragraph(OpenXmlElement? pp, Dictionary<string, string> look)
    {
        if (pp is null) return;
        var own = new Dictionary<string, string>();
        DocxParaFormat.Read(pp, own);
        if (pp.GetFirstChild<W.Indentation>() is { } ind)
        {
            // an explicit 0 undoes the base style's indent (a heading on 正文 with its 2ch first line)
            if (ind.Left is not null || ind.Start is not null || ind.LeftChars is not null) own.TryAdd("indentLeft", "0");
            if (ind.Right is not null || ind.End is not null || ind.RightChars is not null) own.TryAdd("indentRight", "0");
            if (ind.FirstLine is not null || ind.FirstLineChars is not null || ind.Hanging is not null || ind.HangingChars is not null) own.TryAdd("indentFirst", "0");
        }
        foreach (var key in ParagraphKeys) if (own.TryGetValue(key, out var value)) look[key] = value;
    }

    static void Run(DocxDocument doc, OpenXmlElement? rp, Dictionary<string, string> look)
    {
        if (rp is null) return;
        if (rp.GetFirstChild<W.RunFonts>() is { } fonts)
        {
            if (Font(doc, fonts.AsciiTheme?.InnerText, fonts.Ascii?.Value) is { } latin) look["font"] = latin;
            if (Font(doc, fonts.EastAsiaTheme?.InnerText, fonts.EastAsia?.Value) is { } eastAsia) look["fontEa"] = eastAsia;
        }
        if (rp.GetFirstChild<W.FontSize>()?.Val?.Value is { } half && double.TryParse(half, NumberStyles.Float, Inv, out var halfPoints))
            look["size"] = (halfPoints / 2).ToString("0.##", Inv);
        if (rp.GetFirstChild<W.Bold>() is { } bold) look["bold"] = DocxRun.On(bold) ? "true" : "false";
        if (rp.GetFirstChild<W.Italic>() is { } italic) look["italic"] = DocxRun.On(italic) ? "true" : "false";
        if (rp.GetFirstChild<W.Color>()?.Val?.Value is { } color && !color.Equals("auto", StringComparison.OrdinalIgnoreCase)) look["color"] = color.ToUpperInvariant();
    }

    /// <summary>A theme font (minorHAnsi, majorEastAsia, minorBidi…) as the theme's typeface, else the font named; null when neither says.
    /// An East Asian theme font the theme leaves empty is the one it gives the script of the document's East Asian language.</summary>
    public static string? Font(DocxDocument doc, string? theme, string? name)
    {
        var named = string.IsNullOrEmpty(name) ? null : name;
        if (theme is null) return named;
        var scheme = doc.Main.ThemePart?.Theme?.ThemeElements?.FontScheme;
        A.FontCollectionType? fonts = theme.StartsWith("major", StringComparison.Ordinal) ? scheme?.MajorFont : scheme?.MinorFont;
        var eastAsia = theme.EndsWith("EastAsia", StringComparison.Ordinal);
        var typeface = eastAsia ? fonts?.EastAsianFont?.Typeface?.Value
            : theme.EndsWith("Bidi", StringComparison.Ordinal) ? fonts?.ComplexScriptFont?.Typeface?.Value : fonts?.LatinFont?.Typeface?.Value;
        if (string.IsNullOrEmpty(typeface) && eastAsia)
            typeface = fonts?.Elements<A.SupplementalFont>().FirstOrDefault(f => f.Script?.Value == Script(doc))?.Typeface?.Value;
        return string.IsNullOrEmpty(typeface) ? named : typeface;
    }

    /// <summary>The theme's script tag for the document's East Asian language (docDefaults w:lang/@w:eastAsia), simplified Chinese when unset.</summary>
    static string Script(DocxDocument doc) =>
        (doc.Main.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.GetFirstChild<W.Languages>()?.EastAsia?.Value ?? "") switch
        {
            var l when l.StartsWith("ja", StringComparison.OrdinalIgnoreCase) => "Jpan",
            var l when l.StartsWith("ko", StringComparison.OrdinalIgnoreCase) => "Hang",
            var l when l.Equals("zh-TW", StringComparison.OrdinalIgnoreCase) || l.Equals("zh-HK", StringComparison.OrdinalIgnoreCase) || l.Equals("zh-MO", StringComparison.OrdinalIgnoreCase) => "Hant",
            _ => "Hans",
        };
}
