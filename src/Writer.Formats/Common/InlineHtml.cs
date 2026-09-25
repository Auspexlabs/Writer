using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Writer.Core;

namespace Writer.Formats.Common;

/// <summary>Inline HTML in and out of runs: what a browser's contentEditable produces, and what a UI needs to show a paragraph.
/// Block tags become line breaks; unknown tags are ignored but their text kept; script and style content is dropped.</summary>
public static partial class InlineHtml
{
    /// <summary>A page break (a form feed in Word's text), as Word writes one in HTML.</summary>
    public const string PageBreak = "<br style=\"page-break-before:always\">";

    /// <summary>With <paramref name="pageBreaks"/> (Word text), a br styled page-break-before:always or break-before:page is a
    /// page break (a form feed); otherwise every br is a line break.</summary>
    public static List<RunSpec> Parse(string html, bool pageBreaks = false)
    {
        var runs = new List<RunSpec>();
        var stack = new List<(string Tag, RunSpec Style)>();
        var style = new RunSpec("");
        var i = 0;
        while (i < html.Length)
        {
            var lt = html.IndexOf('<', i);
            if (lt < 0)
            {
                Add(runs, style with { Text = WebUtility.HtmlDecode(html[i..]) });
                break;
            }
            if (lt > i) Add(runs, style with { Text = WebUtility.HtmlDecode(html[i..lt]) });
            var gt = TagEnd(html, lt);
            if (gt < 0)
            {
                Add(runs, style with { Text = WebUtility.HtmlDecode(html[lt..]) });
                break;
            }
            var raw = html[(lt + 1)..gt];
            i = gt + 1;
            if (raw.StartsWith("!--", StringComparison.Ordinal))
            {
                var end = html.IndexOf("-->", lt, StringComparison.Ordinal);
                i = end < 0 ? html.Length : end + 3;
                continue;
            }
            var closing = raw.StartsWith('/');
            var name = TagName(closing ? raw[1..] : raw);
            if (name.Length == 0) continue;
            if (closing)
            {
                for (var k = stack.Count - 1; k >= 0; k--)
                {
                    if (stack[k].Tag != name) continue;
                    style = stack[k].Style;
                    stack.RemoveRange(k, stack.Count - k);
                    break;
                }
                continue;
            }
            if (name is "script" or "style" or "template")
            {
                var end = html.IndexOf("</" + name, i, StringComparison.OrdinalIgnoreCase);
                i = end < 0 ? html.Length : html.IndexOf('>', end) + 1;
                if (i == 0) i = html.Length;
                continue;
            }
            if (name == "br")
            {
                var page = pageBreaks && Attributes(raw).TryGetValue("style", out var css) && PageBreakPattern().IsMatch(css);
                Add(runs, style with { Text = page ? "\f" : "\n" });
                continue;
            }
            if (name is "p" or "div" or "li" or "tr" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "blockquote" or "pre" or "ul" or "ol" or "table" or "hr")
            {
                if (runs.Count > 0 && !runs[^1].Text.EndsWith('\n')) Add(runs, style with { Text = "\n" });
                if (name == "hr") continue;
            }
            if (raw.EndsWith('/')) continue;
            var next = Apply(style, name, Attributes(raw));
            stack.Add((name, style));
            style = next;
        }
        return runs;
    }

    /// <summary>The runs as HTML: text escaped, newlines as br, links only to safe schemes.</summary>
    public static string Render(IEnumerable<RunSpec> runs)
    {
        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            var open = new StringBuilder();
            var close = new StringBuilder();
            if (run.Change is "inserted" or "deleted")
            {
                var tag = run.Change == "inserted" ? "ins" : "del";
                open.Append('<').Append(tag);
                if (run.Author is not null) open.Append(" data-author=\"").Append(Esc(run.Author)).Append('"');
                if (run.Date is not null) open.Append(" data-date=\"").Append(Esc(run.Date)).Append('"');
                open.Append('>');
                close.Insert(0, "</" + tag + ">");
            }
            if (run.Link is not null)
            {
                open.Append("<a href=\"").Append(Esc(SafeUrl(run.Link))).Append("\">");
                close.Insert(0, "</a>");
            }
            if (run.Style is not null)
            {
                open.Append("<span data-style=\"").Append(Esc(run.Style)).Append("\">");
                close.Insert(0, "</span>");
            }
            var css = new List<string>();
            if (run.Color is not null) css.Add("color:#" + run.Color);
            if (run.Size is not null) css.Add("font-size:" + run.Size + "pt");
            if (run.Font is not null) css.Add("font-family:'" + Esc(run.Font) + "'");
            if (run.Highlight is not null) css.Add("background-color:#" + run.Highlight);
            if (run.Spacing is not null) css.Add("letter-spacing:" + run.Spacing);
            if (run.Caps is not null) css.Add(run.Caps == "small" ? "font-variant:small-caps" : "text-transform:uppercase");
            if (run.Shadow) css.Add(ShadowCss);
            if (run.Outline) css.Add(OutlineCss);
            if (css.Count > 0)
            {
                open.Append("<span style=\"").Append(string.Join(';', css)).Append("\">");
                close.Insert(0, "</span>");
            }
            Wrap(run.Bold, "b", open, close);
            Wrap(run.Italic, "i", open, close);
            if (run.Underline && run.UnderlineStyle is { } line)
            {
                open.Append("<u style=\"text-decoration-style:").Append(line).Append("\">");
                close.Insert(0, "</u>");
            }
            else Wrap(run.Underline, "u", open, close);
            Wrap(run.Strike, "s", open, close);
            Wrap(run.Code, "code", open, close);
            Wrap(run.VertAlign == "superscript", "sup", open, close);
            Wrap(run.VertAlign == "subscript", "sub", open, close);
            sb.Append(open).Append(Esc(run.Text).Replace("\n", "<br>").Replace("\f", PageBreak)).Append(close);
        }
        return sb.ToString();
    }

    static void Wrap(bool on, string tag, StringBuilder open, StringBuilder close)
    {
        if (!on) return;
        open.Append('<').Append(tag).Append('>');
        close.Insert(0, "</" + tag + ">");
    }

    public static string Esc(string s) => WebUtility.HtmlEncode(s);

    public static string SafeUrl(string url)
    {
        var t = url.Trim();
        var ok = t.StartsWith("http:", StringComparison.OrdinalIgnoreCase) || t.StartsWith("https:", StringComparison.OrdinalIgnoreCase)
            || t.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) || t.StartsWith('#') || !t.Contains(':');
        return ok ? t : "#";
    }

    static RunSpec Apply(RunSpec style, string tag, Dictionary<string, string> attrs)
    {
        var s = tag switch
        {
            "b" or "strong" => style with { Bold = true },
            "i" or "em" => style with { Italic = true },
            "u" => style with { Underline = true },
            "s" or "strike" => style with { Strike = true },
            "ins" or "del" => style with { Change = tag == "ins" ? "inserted" : "deleted", Author = NonEmpty(attrs, "data-author"), Date = NonEmpty(attrs, "data-date") },
            "code" or "tt" or "kbd" or "samp" => style with { Code = true },
            "a" when attrs.TryGetValue("href", out var href) && href.Length > 0 => style with { Link = href },
            "mark" => style with { Highlight = "FFFF00" },
            "sup" => style with { VertAlign = "superscript" },
            "sub" => style with { VertAlign = "subscript" },
            "span" when NonEmpty(attrs, "data-style") is { } characterStyle => style with { Style = characterStyle },
            _ => style,
        };
        if (attrs.TryGetValue("color", out var fontColor) && ParseColor(fontColor) is { } fc) s = s with { Color = fc };
        if (attrs.TryGetValue("face", out var face)) s = s with { Font = face.Split(',')[0].Trim().Trim('"', '\'') };
        if (!attrs.TryGetValue("style", out var css)) return s;
        foreach (var declaration in css.Split(';'))
        {
            var colon = declaration.IndexOf(':');
            if (colon < 0) continue;
            var prop = declaration[..colon].Trim().ToLowerInvariant();
            var value = declaration[(colon + 1)..].Trim();
            switch (prop)
            {
                case "color" when ParseColor(value) is { } c: s = s with { Color = c }; break;
                case "background-color" or "background" when ParseColor(value) is { } h: s = s with { Highlight = h }; break;
                case "font-size" when ParseSize(value) is { } size: s = s with { Size = size }; break;
                case "font-family": s = s with { Font = value.Split(',')[0].Trim().Trim('"', '\'') }; break;
                case "font-weight" when value is "bold" or "bolder" || (int.TryParse(value, out var weight) && weight >= 600): s = s with { Bold = true }; break;
                case "font-weight" when value is "normal": s = s with { Bold = false }; break;
                case "font-style" when value is "italic" or "oblique": s = s with { Italic = true }; break;
                case "text-decoration" or "text-decoration-line":
                    if (value.Contains("underline")) s = s with { Underline = true, UnderlineStyle = LineStyle(value) ?? s.UnderlineStyle };
                    if (value.Contains("line-through")) s = s with { Strike = true };
                    if (value == "none") s = s with { Underline = false, Strike = false, UnderlineStyle = null };
                    break;
                case "text-decoration-style" when s.Underline: s = s with { UnderlineStyle = LineStyle(value) }; break;
                case "text-transform": s = s with { Caps = value == "uppercase" ? "all" : null }; break;
                case "font-variant" or "font-variant-caps": s = s with { Caps = value.Contains("small-caps") ? "small" : null }; break;
                case "vertical-align": s = s with { VertAlign = value is "super" or "text-top" ? "superscript" : value is "sub" or "text-bottom" ? "subscript" : null }; break;
                case "letter-spacing": s = s with { Spacing = value is "normal" or "0" ? null : ParseSpacing(value) }; break;
                case "text-shadow": s = s with { Shadow = value != "none" }; break;
                case "-webkit-text-stroke" or "-webkit-text-stroke-width": s = s with { Outline = value != "0" && !value.StartsWith("0px", StringComparison.Ordinal) && value != "none" }; break;
            }
        }
        return s;
    }

    /// <summary>The underline's line style a text-decoration names, if any but solid.</summary>
    static string? LineStyle(string value) => new[] { "double", "dotted", "dashed", "wavy" }.FirstOrDefault(value.Contains);

    static string? NonEmpty(Dictionary<string, string> attrs, string name) => attrs.TryGetValue(name, out var v) && v.Length > 0 ? v : null;

    /// <summary>For raw tags met elsewhere (inline HTML in markdown): the change an ins or del tag opens, with its data-author
    /// and data-date; an empty spec for the closing tag; null for any other tag.</summary>
    public static RunSpec? RevisionTag(string tag)
    {
        var raw = tag.Trim().TrimStart('<').TrimEnd('>').TrimEnd('/');
        var closing = raw.StartsWith('/');
        var name = TagName(closing ? raw[1..] : raw);
        if (name is not ("ins" or "del")) return null;
        return closing ? new RunSpec("") : Apply(new RunSpec(""), name, Attributes(raw));
    }

    /// <summary>RRGGBB from #hex, rgb(), rgba() or a color name; null when unknown or transparent.</summary>
    public static string? ParseColor(string value)
    {
        var v = value.Trim();
        var rgb = RgbPattern().Match(v);
        if (rgb.Success)
        {
            if (rgb.Groups[4].Success && double.TryParse(rgb.Groups[4].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var alpha) && alpha == 0) return null;
            return string.Concat(new[] { rgb.Groups[1].Value, rgb.Groups[2].Value, rgb.Groups[3].Value }.Select(x => Math.Clamp(int.Parse(x, CultureInfo.InvariantCulture), 0, 255).ToString("X2", CultureInfo.InvariantCulture)));
        }
        if (v.Length == 4 && v[0] == '#') v = "#" + string.Concat(v[1..].Select(c => new string(c, 2)));
        if (v is "transparent" or "inherit" or "initial" or "currentcolor") return null;
        try
        {
            var parsed = Units.ParseColor(v);
            return parsed == "none" ? null : parsed;
        }
        catch (WriterException)
        {
            return null;
        }
    }

    /// <summary>Word's text effects as CSS: a soft shadow, and outlined letters (the stroke drawn, the fill transparent).</summary>
    public const string ShadowCss = "text-shadow:1px 1px 0 #B3B3B3", OutlineCss = "-webkit-text-stroke:0.5px currentColor;-webkit-text-fill-color:transparent";

    /// <summary>Character spacing in points (2pt, -0.5pt) from a letter-spacing value in pt, px or em.</summary>
    static string? ParseSpacing(string value)
    {
        var m = SizePattern().Match(value.Trim().TrimStart('+'));
        if (!m.Success) return null;
        var n = double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture) * (value.Trim().StartsWith('-') ? -1 : 1);
        var pt = m.Groups[2].Value.ToLowerInvariant() switch { "pt" => n, "px" => n * 0.75, _ => n * 12 };
        return pt == 0 ? null : Math.Round(pt, 2).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
    }

    static string? ParseSize(string value)
    {
        var m = SizePattern().Match(value.Trim());
        if (!m.Success) return null;
        var n = double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
        var pt = m.Groups[2].Value.ToLowerInvariant() switch { "pt" => n, "px" => n * 0.75, "em" or "rem" => n * 12, _ => -1 };
        return pt <= 0 ? null : Math.Round(pt, 1).ToString("0.##", CultureInfo.InvariantCulture);
    }

    static int TagEnd(string html, int lt)
    {
        char? quote = null;
        for (var i = lt + 1; i < html.Length; i++)
        {
            var c = html[i];
            if (quote is not null) { if (c == quote) quote = null; }
            else if (c is '"' or '\'') quote = c;
            else if (c == '>') return i;
        }
        return -1;
    }

    static string TagName(string raw)
    {
        var end = 0;
        while (end < raw.Length && (char.IsLetterOrDigit(raw[end]) || raw[end] == '-')) end++;
        return raw[..end].ToLowerInvariant();
    }

    static Dictionary<string, string> Attributes(string raw)
    {
        var attrs = new Dictionary<string, string>();
        foreach (Match m in AttributePattern().Matches(raw))
            attrs[m.Groups[1].Value.ToLowerInvariant()] = WebUtility.HtmlDecode(m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Success ? m.Groups[3].Value : m.Groups[4].Value);
        return attrs;
    }

    static void Add(List<RunSpec> runs, RunSpec run)
    {
        if (run.Text.Length == 0) return;
        if (runs.Count > 0 && runs[^1] with { Text = "" } == run with { Text = "" })
            runs[^1] = runs[^1] with { Text = runs[^1].Text + run.Text };
        else runs.Add(run);
    }

    [GeneratedRegex(@"\s([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*=\s*(?:""([^""]*)""|'([^']*)'|([^\s""'>]+))")]
    private static partial Regex AttributePattern();

    [GeneratedRegex(@"^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+)\s*)?\)$", RegexOptions.IgnoreCase)]
    private static partial Regex RgbPattern();

    [GeneratedRegex(@"^-?([\d.]+)\s*(pt|px|em|rem)$", RegexOptions.IgnoreCase)]
    private static partial Regex SizePattern();

    [GeneratedRegex(@"(?:page-break-before\s*:\s*always|break-before\s*:\s*page)", RegexOptions.IgnoreCase)]
    private static partial Regex PageBreakPattern();
}
