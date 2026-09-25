using System.Globalization;
using System.Text;
using Writer.Core;

namespace Writer.Formats.Html;

/// <summary>A Word document drawn as Word lays it out: each section a sheet of its paper size with its margins, columns, header and
/// footer, the section's footnotes under its text and the endnotes after the last; text in the document's own fonts, headings as their
/// styles make them (no browser sizes), paragraphs with the spacing and indents their styles give.</summary>
public static partial class HtmlWriter
{
    const string DocxCss = """
        body.docx{max-width:none;background:#f0f0f0}
        .docx .page{position:relative;isolation:isolate;box-sizing:border-box;margin:0 auto 24px auto;padding:0}
        .docx td{padding:0 .19cm;border-color:#000;line-height:1.2}.docx .page td p{padding-top:0;padding-bottom:0}
        .docx .page>header,.docx .page>footer{position:absolute}
        .docx p,.docx h1,.docx h2,.docx h3,.docx h4,.docx h5,.docx h6{margin:0;font-size:1em;font-weight:inherit}
        .docx .notes{margin-top:12pt;font-size:.8em}.docx .notes::before{content:"";display:block;width:33%;border-top:1px solid #999;margin-bottom:4pt}
        """;

    /// <summary>The notes met so far on this page (footnotes) or in the document (endnotes), with their numbers; set while a document is drawn.</summary>
    [ThreadStatic] static List<(bool Endnote, string Label, string Text)>? _notes;

    /// <summary>The page CSS for the document's own look: its fonts, size and colour, and every paragraph's spacing from its default style.</summary>
    static string DocxLookCss(IReadOnlyDictionary<string, string> look)
    {
        var text = look.Where(x => x.Key is "font" or "fontEa" or "size" or "color").ToDictionary();
        text.TryAdd("font", "Times New Roman"); // Word's own when nothing names one
        text.TryAdd("size", "10");
        return $"\n.docx .page{{{string.Join(';', TextCss(text))}}}\n.docx .page p,.docx .page h1,.docx .page h2,.docx .page h3,.docx .page h4,.docx .page h5,.docx .page h6{{{string.Join(';', ParaCss(look))}}}";
    }

    /// <summary>The body as sheets, one per section: a paragraph that ends a section (sectionBreak) carries that section's page; the last
    /// section's is the document's.</summary>
    static void Pages(Node body, IReadOnlyDictionary<string, string> root, StringBuilder sb)
    {
        (_notes, _footnoteCount, _endnoteCount) = ([], 0, 0);
        try
        {
            var blocks = body.Children;
            var (start, first) = (0, true);
            for (var i = 0; i < blocks.Count; i++)
                if (blocks[i].Kind is "paragraph" or "heading" && blocks[i].GetProps().ContainsKey("sectionBreak"))
                {
                    Page(blocks.Skip(start).Take(i + 1 - start).ToList(), WithComputed(blocks[i]), root, first, sb);
                    (start, first) = (i + 1, false);
                }
            if (first || start < blocks.Count) Page(blocks.Skip(start).ToList(), root, root, first, sb); // no empty sheet after a last break
            Notes(endnotes: true, sb);
        }
        finally { _notes = null; }
    }

    static void Page(IReadOnlyList<Node> blocks, IReadOnlyDictionary<string, string> page, IReadOnlyDictionary<string, string> root, bool first, StringBuilder sb)
    {
        var size = (page.GetValueOrDefault("pageSize") ?? "21cm x 29.7cm").Split(" x ");
        var margin = page.GetValueOrDefault("pageMargin") ?? "2.54cm 3.18cm 2.54cm 3.18cm";
        var sides = margin.Split(' ');
        sb.Append("<section class=\"page\" style=\"width:").Append(size[0]).Append(";min-height:").Append(size[^1]).Append(";padding:").Append(margin).Append("\">\n");
        var titlePage = first && root.GetValueOrDefault("titlePg") == "true";
        if (root.GetValueOrDefault(titlePage ? "firstHeader" : "header") is { Length: > 0 } header)
            sb.Append("<header style=\"left:").Append(sides[3]).Append(";right:").Append(sides[1]).Append(";top:").Append(page.GetValueOrDefault("headerDistance") ?? "1.27cm").Append("\">")
                .Append(PageNumbers(header)).Append("</header>\n");
        var columns = page.GetValueOrDefault("columns") is { } c && c != "1" ? c : null;
        sb.Append(columns is null ? "<div>\n" : $"<div style=\"column-count:{columns};column-gap:{page.GetValueOrDefault("columnGap") ?? "1.27cm"}\">\n");
        Blocks(blocks, sb);
        sb.Append("</div>\n");
        Notes(endnotes: false, sb);
        if (root.GetValueOrDefault(titlePage ? "firstFooter" : "footer") is { Length: > 0 } footer)
            sb.Append("<footer style=\"left:").Append(sides[3]).Append(";right:").Append(sides[1]).Append(";bottom:").Append(page.GetValueOrDefault("footerDistance") ?? "1.27cm").Append("\">")
                .Append(PageNumbers(footer)).Append("</footer>\n");
        sb.Append("</section>\n");
    }

    /// <summary>A header's page fields as the one page this view draws: page 1 of 1.</summary>
    static string PageNumbers(string html) => html.Replace("{page}", "1").Replace("{pages}", "1");

    /// <summary>The footnotes of the page (and the page's endnotes at the end of the document) under a short rule.</summary>
    static void Notes(bool endnotes, StringBuilder sb)
    {
        var notes = _notes?.Where(n => n.Endnote == endnotes).ToList() ?? [];
        if (notes.Count == 0) return;
        sb.Append("<aside class=\"notes\">\n");
        foreach (var (_, label, text) in notes) sb.Append("<p><sup>").Append(label).Append("</sup> ").Append(Esc(text).Replace("\n", "<br>")).Append("</p>\n");
        sb.Append("</aside>\n");
        _notes!.RemoveAll(n => n.Endnote == endnotes);
    }

    /// <summary>A note's reference mark, numbered in reading order as Word does (footnotes 1, 2, 3, endnotes i, ii, iii), its text kept for the page's foot.</summary>
    static string NoteMark(IReadOnlyDictionary<string, string> note)
    {
        var endnote = note.GetValueOrDefault("kind") == "endnote";
        var label = endnote ? Roman(++_endnoteCount) : (++_footnoteCount).ToString(CultureInfo.InvariantCulture);
        _notes!.Add((endnote, label, note.GetValueOrDefault("text") ?? ""));
        return "<sup class=\"note\">" + label + "</sup>";
    }

    [ThreadStatic] static int _footnoteCount, _endnoteCount;

    static string Roman(int n)
    {
        var sb = new StringBuilder();
        foreach (var (value, digits) in new[] { (100, "c"), (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i") })
            for (; n >= value; n -= value) sb.Append(digits);
        return sb.ToString();
    }

    /// <summary>A paragraph's text look as CSS: its fonts (the East Asian one after the Latin, as Word picks per character), size, weight, slant, colour.</summary>
    static IEnumerable<string> TextCss(IReadOnlyDictionary<string, string> p)
    {
        var fonts = new[] { p.GetValueOrDefault("font"), p.GetValueOrDefault("fontEa") }.Where(f => !string.IsNullOrEmpty(f)).Distinct().Select(f => "'" + Esc(f!) + "'").ToList();
        if (fonts.Count > 0) yield return "font-family:" + string.Join(',', fonts);
        if (p.TryGetValue("size", out var size)) yield return "font-size:" + size + "pt";
        if (p.TryGetValue("bold", out var bold)) yield return "font-weight:" + (bold == "true" ? "bold" : "normal");
        if (p.TryGetValue("italic", out var italic)) yield return "font-style:" + (italic == "true" ? "italic" : "normal");
        if (p.TryGetValue("color", out var color) && color != "none") yield return "color:#" + color;
    }

    /// <summary>A paragraph's 段落 settings as CSS: line spacing (a multiple of Word's single line, about 1.2 of the font size), space before
    /// and after as padding (Word adds the two where CSS margins would overlap), indents (a negative first line hangs), borders.</summary>
    static IEnumerable<string> ParaCss(IReadOnlyDictionary<string, string> p)
    {
        if (p.TryGetValue("lineSpacing", out var line) && line != "none")
        {
            if (line.StartsWith("min ", StringComparison.Ordinal)) yield return $"line-height:max(1.2em,{Len(line[4..])})";
            else if (double.TryParse(line, NumberStyles.Float, CultureInfo.InvariantCulture, out var multiple)) yield return "line-height:" + (multiple * 1.2).ToString("0.###", CultureInfo.InvariantCulture);
            else yield return "line-height:" + Len(line);
        }
        if (p.TryGetValue("spaceBefore", out var before) && before != "none") yield return "padding-top:" + Len(before);
        if (p.TryGetValue("spaceAfter", out var after) && after != "none") yield return "padding-bottom:" + Len(after);
        if (p.TryGetValue("indentLeft", out var left) && left != "none") yield return "margin-left:" + Len(left);
        if (p.TryGetValue("indentRight", out var right) && right != "none") yield return "margin-right:" + Len(right);
        if (p.TryGetValue("indentFirst", out var firstLine) && firstLine != "none") yield return "text-indent:" + Len(firstLine);
        if (p.TryGetValue("border", out var border) && border != "none")
            foreach (var side in border == "box" ? ["top", "bottom", "left", "right"] : border.Split(' ', StringSplitOptions.RemoveEmptyEntries)) yield return $"border-{side}:1px solid";
    }

    /// <summary>A paragraph's pictures where Word puts them: in the line, floated to a side with the text around it (square, tight,
    /// through), on a line of their own (topAndBottom), or behind or in front of the text at their offset from the paragraph.</summary>
    static void Pictures(Node paragraph, StringBuilder sb)
    {
        foreach (var picture in paragraph.Children.Where(c => c.Kind == "image"))
        {
            var p = picture.GetProps();
            var css = new List<string>();
            if (Px(p.GetValueOrDefault("width")) is > 0 and var width) css.Add("width:" + PxText(width));
            if (Px(p.GetValueOrDefault("height")) is > 0 and var height) css.Add("height:" + PxText(height));
            var align = p.GetValueOrDefault("xAlign");
            var wrap = p.GetValueOrDefault("wrap");
            switch (wrap)
            {
                case "square" or "tight" or "through":
                    css.Add(align == "center" ? "display:block;margin:0 auto" : align is "right" or "outside" ? "float:right;margin:0 0 4px 12px" : "float:left;margin:0 12px 4px 0");
                    break;
                case "topAndBottom":
                    css.Add("display:block;margin:" + (align == "center" ? "0 auto" : align == "right" ? "0 0 0 auto" : "0"));
                    break;
                case "behind" or "inFront":
                    css.Add("position:absolute;z-index:" + (wrap == "behind" ? "-1" : "1"));
                    css.Add(align == "center" ? "left:50%;transform:translateX(-50%)" : align == "right" ? "right:0" : "left:" + PxText(Px(p.GetValueOrDefault("x"))));
                    css.Add("top:" + PxText(Px(p.GetValueOrDefault("y"))));
                    break;
            }
            sb.Append("<img src=\"").Append(ImageSource(picture, p)).Append("\" style=\"").Append(string.Join(';', css)).Append('"');
            if (p.TryGetValue("alt", out var alt)) sb.Append(" alt=\"").Append(Esc(alt)).Append('"');
            sb.Append('>');
        }
    }

    /// <summary>position:relative for a paragraph holding a picture behind or in front of its text, which is placed from it.</summary>
    static string? Anchored(Node paragraph) =>
        paragraph.Children.Any(c => c.Kind == "image" && c.GetProps().GetValueOrDefault("wrap") is "behind" or "inFront") ? "position:relative" : null;

    /// <summary>A JSON array of lengths (a table's column widths) as CSS lengths; none when it is not one.</summary>
    static List<string> Lengths(string json)
    {
        try
        {
            using var parsed = System.Text.Json.JsonDocument.Parse(json);
            return parsed.RootElement.ValueKind == System.Text.Json.JsonValueKind.Array ? parsed.RootElement.EnumerateArray().Select(e => e.ToString()).ToList() : [];
        }
        catch (System.Text.Json.JsonException) { return []; }
    }

    /// <summary>A Word length as CSS: characters (2ch, as CJK Word counts an indent) are ems; cm, pt and 0 as they are.</summary>
    static string Len(string value) => value.EndsWith("ch", StringComparison.Ordinal) ? value[..^2] + "em" : value;
}
