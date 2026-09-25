using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>The document's style gallery: what styles.xml defines for paragraphs and runs, each with the look it gives on its
/// own (inherited along basedOn, not the document defaults), read as JSON; and defining or changing a style from the same
/// fields, which is how the editor's 修改样式 and 新建样式 reach the file.</summary>
static class DocxStyleGallery
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    static readonly string[] LookKeys = ["font", "size", "bold", "italic", "color", "align", "spaceBefore", "spaceAfter", "lineSpacing", "indentFirst"];

    public static string Read(DocxDocument doc)
    {
        var styles = doc.Main.StyleDefinitionsPart?.Styles?.Elements<W.Style>().Where(s => s.StyleId?.Value is { Length: > 0 } && s.SemiHidden is null && s.Type?.Value is var t && (t is null || t == W.StyleValues.Paragraph || t == W.StyleValues.Character)).ToList() ?? [];
        return NodeJson.Compact(w =>
        {
            w.WriteStartArray();
            foreach (var style in styles)
            {
                var id = style.StyleId!.Value!;
                w.WriteStartObject();
                w.WriteString("id", id);
                w.WriteString("name", style.StyleName?.Val?.Value ?? id);
                w.WriteString("type", style.Type?.Value == W.StyleValues.Character ? "character" : "paragraph");
                if (style.BasedOn?.Val?.Value is { } basedOn) w.WriteString("basedOn", basedOn);
                if (style.Type?.Value != W.StyleValues.Character && HeadingLevel(doc, style) is { } level) w.WriteNumber("heading", level);
                w.WriteStartObject("look");
                foreach (var (key, value) in Look(doc, style)) w.WriteString(key, value);
                w.WriteEndObject();
                w.WriteEndObject();
            }
            w.WriteEndArray();
        });
    }

    static int? HeadingLevel(DocxDocument doc, W.Style style)
    {
        var p = new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = style.StyleId!.Value }));
        var level = doc.Styles.HeadingLevel(p);
        return level > 0 ? level : null;
    }

    /// <summary>The style's own look, base styles first so the style's own settings win.</summary>
    static Dictionary<string, string> Look(DocxDocument doc, W.Style style)
    {
        var chain = new List<W.Style>();
        for (var s = style; s is not null && chain.Count < 16 && !chain.Contains(s); s = doc.Styles.Find(s.BasedOn?.Val?.Value)) chain.Add(s);
        chain.Reverse();
        var look = new Dictionary<string, string>();
        foreach (var s in chain)
        {
            if (s.StyleRunProperties is { } rp)
            {
                if (rp.RunFonts is { } fonts && (fonts.Ascii?.Value ?? fonts.EastAsia?.Value ?? fonts.HighAnsi?.Value) is { } font) look["font"] = font;
                if (rp.FontSize?.Val?.Value is { } half && double.TryParse(half, NumberStyles.Float, Inv, out var halfPoints)) look["size"] = (halfPoints / 2).ToString("0.##", Inv);
                if (rp.Bold is { } b) look["bold"] = DocxRun.On(b) ? "true" : "false";
                if (rp.Italic is { } i) look["italic"] = DocxRun.On(i) ? "true" : "false";
                if (rp.Color?.Val?.Value is { } color && !color.Equals("auto", StringComparison.OrdinalIgnoreCase)) look["color"] = color.ToUpperInvariant();
            }
            if (s.StyleParagraphProperties is { } pp)
            {
                if (DocxParagraph.AlignOf(pp.Justification) is { } align) look["align"] = align;
                var format = new Dictionary<string, string>();
                DocxParaFormat.Read(pp, format);
                foreach (var key in new[] { "spaceBefore", "spaceAfter", "lineSpacing", "indentFirst" }) if (format.TryGetValue(key, out var v)) look[key] = v;
            }
        }
        return look;
    }

    /// <summary>Defines a style from {id|name, type, basedOn, look fields…}: an existing one (by id, then by name) is changed,
    /// a new one added with the template's shape; none clears a setting; fields left out stay.</summary>
    public static void Define(DocxDocument doc, string json)
    {
        using var parsed = JsonDocument.Parse(json);
        var o = parsed.RootElement;
        if (o.ValueKind != JsonValueKind.Object) throw new WriterException(ErrorCode.Validation, "style takes a JSON object", "Example: {\"id\":\"Note\",\"name\":\"Note\",\"italic\":true}.");
        string? Text(string key) => o.TryGetProperty(key, out var v) ? v.ValueKind switch { JsonValueKind.String => v.GetString(), JsonValueKind.Null => "none", JsonValueKind.True => "true", JsonValueKind.False => "false", _ => v.GetRawText() } : null;
        var name = Text("name");
        var id = Text("id") ?? (name is null ? null : new string(name.Where(char.IsLetterOrDigit).ToArray()));
        if (string.IsNullOrEmpty(id)) throw new WriterException(ErrorCode.Validation, "style needs an id or a name", "Example: {\"id\":\"Note\",\"name\":\"Note\",\"italic\":true}.");
        var type = Text("type") == "character" ? "character" : "paragraph";
        var styles = doc.Main.StyleDefinitionsPart?.Styles ?? throw new WriterException(ErrorCode.Validation, "The document has no styles part", "Save it from Word once.");
        var style = doc.Styles.Find(id) ?? styles.Elements<W.Style>().FirstOrDefault(s => string.Equals(s.StyleName?.Val?.Value, name ?? id, StringComparison.OrdinalIgnoreCase));
        if (style is null)
        {
            style = new W.Style(new W.StyleName { Val = name ?? id }, new W.BasedOn { Val = type == "character" ? "DefaultParagraphFont" : "Normal" }, new W.PrimaryStyle()) { Type = type == "character" ? W.StyleValues.Character : W.StyleValues.Paragraph, StyleId = id };
            if (type == "paragraph") style.InsertAfter(new W.NextParagraphStyle { Val = "Normal" }, style.BasedOn);
            styles.Append(style);
            doc.Styles.Invalidate();
        }
        else if (name is not null) style.StyleName = new W.StyleName { Val = name };
        if (Text("basedOn") is { } basedOn) style.BasedOn = basedOn == "none" ? null : new W.BasedOn { Val = doc.Styles.ResolveStyle(basedOn, style.Type?.Value == W.StyleValues.Character ? "character" : "paragraph") };
        var rp = style.StyleRunProperties ??= new W.StyleRunProperties();
        if (Text("font") is { } font) { if (font == "none") rp.RunFonts = null; else rp.RunFonts = new W.RunFonts { Ascii = font, HighAnsi = font, EastAsia = font, ComplexScript = font }; }
        if (Text("size") is { } size)
        {
            if (size == "none") { rp.FontSize = null; rp.FontSizeComplexScript = null; }
            else { var half = ((int)Math.Round(double.Parse(size.TrimEnd('p', 't'), NumberStyles.Float, Inv) * 2)).ToString(Inv); rp.FontSize = new W.FontSize { Val = half }; rp.FontSizeComplexScript = new W.FontSizeComplexScript { Val = half }; }
        }
        if (Text("bold") is { } bold) rp.Bold = bold == "none" ? null : new W.Bold { Val = bold == "true" ? null : false };
        if (Text("italic") is { } italic) rp.Italic = italic == "none" ? null : new W.Italic { Val = italic == "true" ? null : false };
        if (Text("color") is { } color) rp.Color = color == "none" ? null : new W.Color { Val = Units.ParseColor(color) };
        if (!rp.HasChildren) rp.Remove();
        if (style.Type?.Value == W.StyleValues.Character) return;
        var pp = style.StyleParagraphProperties ??= new W.StyleParagraphProperties();
        if (Text("align") is { } align) pp.Justification = align == "none" ? null : DocxParagraph.JustificationOf(align);
        var format = new Dictionary<string, string>();
        foreach (var key in new[] { "spaceBefore", "spaceAfter", "lineSpacing", "indentFirst" }) if (Text(key) is { } v) format[key] = v;
        if (format.Count > 0)
        {
            // the paragraph writer works on a pPr: seed one with the style's spacing and indent, apply, and put the children back
            var scratch = new W.ParagraphProperties();
            foreach (var child in pp.ChildElements.Where(c => c is W.SpacingBetweenLines or W.Indentation)) scratch.Append(child.CloneNode(true));
            foreach (var (key, v) in format) DocxParaFormat.Write(scratch, key, v);
            pp.RemoveAllChildren<W.SpacingBetweenLines>();
            pp.RemoveAllChildren<W.Indentation>();
            foreach (var child in scratch.ChildElements.ToList()) { child.Remove(); pp.AddChild(child); }
        }
        if (!pp.HasChildren) pp.Remove();
    }
}
