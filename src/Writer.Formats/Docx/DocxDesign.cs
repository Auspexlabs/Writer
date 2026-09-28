using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Pptx;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

static class DocxDesign
{
    public static void Theme(DocxDocument doc, string key)
    {
        var palette = PptxTemplate.FindPalette(key) ?? throw new WriterException(ErrorCode.Validation, "Unknown theme", "Use paper, mist, clay, sand or rose.");
        var part = doc.Main.ThemePart ?? doc.Main.AddNewPart<ThemePart>();
        var theme = part.Theme ??= new A.Theme(PptxTemplate.ThemeXml);
        theme.Name = "Writer " + key;
        theme.ThemeElements!.ColorScheme = new A.ColorScheme(PptxTemplate.ColorSchemeXml(palette));
        theme.ThemeElements.FontScheme = new A.FontScheme(PptxTemplate.FontSchemeXml(theme.Name!, palette.Heading, palette.Body));
        foreach (var id in new[] { "Normal", "Title", "Subtitle", "Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6" })
        {
            doc.Styles.ResolveStyle(id, "paragraph");
            if (doc.Styles.Find(id) is not { } style) continue;
            var heading = id != "Normal";
            var rp = style.StyleRunProperties ??= new W.StyleRunProperties();
            rp.RunFonts = new W.RunFonts { AsciiTheme = heading ? W.ThemeFontValues.MajorAscii : W.ThemeFontValues.MinorAscii, HighAnsiTheme = heading ? W.ThemeFontValues.MajorHighAnsi : W.ThemeFontValues.MinorHighAnsi, EastAsiaTheme = heading ? W.ThemeFontValues.MajorEastAsia : W.ThemeFontValues.MinorEastAsia };
            rp.Color = new W.Color { Val = heading ? palette.Accents[0] : "1D1D1F", ThemeColor = heading ? W.ThemeColorValues.Accent1 : null };
        }
        doc.Styles.Invalidate();
    }

    public static void StyleSet(DocxDocument doc, string key)
    {
        if (key is not ("modern" or "formal" or "compact")) throw new WriterException(ErrorCode.Validation, "Unknown style set", "Use modern, formal or compact.");
        var compact = key == "compact"; var formal = key == "formal";
        for (var level = 0; level <= 6; level++)
        {
            var id = level == 0 ? "Normal" : "Heading" + level;
            DocxStyleGallery.Define(doc, NodeJson.Compact(w => {
                w.WriteStartObject(); w.WriteString("id", doc.Styles.ResolveStyle(id, "paragraph"));
                w.WriteString("size", (level == 0 ? formal ? 12 : 11 : Math.Max(11, (compact ? 17 : 22) - level * 2)).ToString(System.Globalization.CultureInfo.InvariantCulture));
                w.WriteBoolean("bold", level > 0); w.WriteString("lineSpacing", formal ? "1.5" : compact ? "1" : "1.15");
                w.WriteString("spaceBefore", level == 0 ? "0pt" : compact ? "6pt" : "12pt"); w.WriteString("spaceAfter", compact ? "0pt" : "6pt");
                w.WriteEndObject();
            }));
        }
        doc.Styles.Invalidate();
    }

    public static void Protection(DocxDocument doc, bool enabled)
    {
        var settings = (doc.Main.DocumentSettingsPart ?? doc.Main.AddNewPart<DocumentSettingsPart>()).Settings ??= new W.Settings();
        if (settings.GetFirstChild<W.DocumentProtection>() is { } existing && (existing.Hash is not null || existing.HashValue is not null))
            throw new WriterException(ErrorCode.Validation, "This document uses password protection", "Remove password protection in Word before editing here.");
        settings.RemoveAllChildren<W.DocumentProtection>();
        if (enabled) settings.AddChild(new W.DocumentProtection { Edit = W.DocumentProtectionValues.ReadOnly, Enforcement = true });
    }
}

static partial class DocxSection
{
    public static string PageBorder(W.SectionProperties? section)
    {
        var b = section?.GetFirstChild<W.PageBorders>()?.GetFirstChild<W.TopBorder>();
        return b is null ? "" : NodeJson.Compact(w => { w.WriteStartObject(); w.WriteString("style", b.Val?.InnerText ?? "single"); w.WriteString("color", b.Color?.Value ?? "000000"); w.WriteNumber("width", (b.Size?.Value ?? 4) / 8.0); w.WriteNumber("space", b.Space?.Value ?? 24); w.WriteEndObject(); });
    }
    public static void SetPageBorder(DocxDocument doc, string json)
    {
        using var data = JsonDocument.Parse(json); var o = data.RootElement;
        var style = o.TryGetProperty("style", out var s) ? s.GetString() : "none";
        var val = style switch { "single" => W.BorderValues.Single, "double" => W.BorderValues.Double, "dotted" => W.BorderValues.Dotted, "dashed" => W.BorderValues.Dashed, "none" => W.BorderValues.Nil, _ => throw new WriterException(ErrorCode.Validation, "Invalid page border style", "Use single, double, dotted, dashed or none.") };
        var color = o.TryGetProperty("color", out var c) ? Units.ParseColor(c.GetString()!) : "000000";
        var width = o.TryGetProperty("width", out var wi) ? wi.GetDouble() : 0.5;
        var space = o.TryGetProperty("space", out var sp) ? sp.GetDouble() : 24;
        if (!double.IsFinite(width) || width is < 0.25 or > 12 || !double.IsFinite(space) || space is < 0 or > 31) throw new WriterException(ErrorCode.Validation, "Invalid border width or inset", "Width: 0.25–12 pt; inset: 0–31 pt.");
        foreach (var section in Sections(doc))
        {
            section.RemoveAllChildren<W.PageBorders>(); if (style == "none") continue;
            var borders = new W.PageBorders { OffsetFrom = W.PageBorderOffsetValues.Page };
            foreach (var border in new W.BorderType[] { new W.TopBorder(), new W.LeftBorder(), new W.BottomBorder(), new W.RightBorder() }) { border.Val = val; border.Color = color; border.Size = (uint)Math.Round(width * 8); border.Space = (uint)space; borders.Append(border); }
            section.AddChild(borders);
        }
    }
}
