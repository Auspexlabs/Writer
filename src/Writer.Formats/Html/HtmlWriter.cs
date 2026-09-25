using System.Globalization;
using System.Net;
using System.Text;
using Writer.Core;

namespace Writer.Formats.Html;

/// <summary>The unified tree as a self-contained HTML page. Flow layout for documents, absolute layout for slides.</summary>
public static partial class HtmlWriter
{
    const string Css = """
        body{font-family:-apple-system,"Segoe UI",Helvetica,Arial,"PingFang SC","Microsoft YaHei",sans-serif;max-width:820px;margin:2rem auto;padding:0 1rem;line-height:1.5;color:#222}
        body.deck{max-width:none;background:#e9e9e9}
        table{border-collapse:collapse;margin:1em 0}td,th{border:1px solid #bbb;padding:4px 8px;vertical-align:top;text-align:left}th{background:#f3f3f3}
        pre{background:#f5f5f5;padding:.75em;overflow:auto}code{font-family:Consolas,Menlo,monospace}img{max-width:100%}
        .slide{position:relative;overflow:hidden;background:#fff;box-shadow:0 2px 12px rgba(0,0,0,.25);margin:0 auto 24px auto;font-size:14pt;line-height:1.2}
        .shape{position:absolute;box-sizing:border-box;padding:6px 8px;overflow:hidden}
        .shape p{margin:0 0 .25em 0}.shape ul,.shape ol{margin:0;padding-left:1.2em}
        .slide table{position:absolute;margin:0;background:#fff;font-size:12pt}.slide img{position:absolute;max-width:none}
        .page{background:#fff;box-shadow:0 2px 12px rgba(0,0,0,.15);padding:2rem;margin:0 0 24px 0}
        .toc ul{list-style:none;padding-left:0}.toc li{display:flex;justify-content:space-between}.toc-title{font-weight:bold}
        """;

    public static string Render(Document doc)
    {
        var sb = new StringBuilder();
        var rootProps = doc.Root.GetProps();
        var word = doc.Format == "docx";
        var look = word ? WithComputed(doc.Root) : rootProps;
        var title = rootProps.GetValueOrDefault("title") ?? FirstHeading(doc.Root) ?? "Document";
        sb.Append("<!doctype html>\n<html>\n<head>\n<meta charset=\"utf-8\">\n<title>").Append(Esc(title)).Append("</title>\n<style>\n").Append(Css).Append(doc.Format == "xlsx" ? SheetCss : "")
            .Append(word ? DocxCss + DocxLookCss(look) : "").Append("\n</style>\n</head>\n");
        sb.Append(doc.Format switch { "pptx" => "<body class=\"deck\">\n", "xlsx" => "<body class=\"book\">\n", "docx" => "<body class=\"docx\">\n", _ => "<body>\n" });
        var book = new Book(doc);
        foreach (var top in doc.Root.Children)
        {
            if (top.Kind == "body" && word) Pages(top, look, sb);
            else if (top.Kind == "body") Blocks(top.Children, sb);
            else if (top.Kind == "slide") Slide(top, rootProps, sb);
            else if (top.Kind == "sheet") Sheet(top, book, rootProps, sb);
            else Block(top, sb);
        }
        sb.Append("</body>\n</html>\n");
        return sb.ToString();
    }

    static string? FirstHeading(Node root) =>
        PathResolver.Query(root, "//heading").FirstOrDefault()?.Text is { Length: > 0 } t ? t : null;

    static double Px(string? emu) => emu is not null && long.TryParse(emu, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var v) ? Math.Round(v / (double)Units.EmuPerPx, 1) : 0;

    static string PxText(double px) => px.ToString("0.#", CultureInfo.InvariantCulture) + "px";

    /// <summary>A slide as a fixed-size box; decor from the layout and master comes first, then shapes, pictures and tables at their real positions.</summary>
    static void Slide(Node slide, IReadOnlyDictionary<string, string> rootProps, StringBuilder sb)
    {
        var props = WithComputed(slide);
        var width = Px(rootProps.GetValueOrDefault("width") ?? "12192000");
        var height = Px(rootProps.GetValueOrDefault("height") ?? "6858000");
        sb.Append("<section class=\"slide\" style=\"width:").Append(PxText(width)).Append(";height:").Append(PxText(height));
        if (props.TryGetValue("background", out var bg) && bg != "none") sb.Append(";background:#").Append(bg);
        sb.Append("\">\n");
        foreach (var child in slide.Children)
        {
            var p = WithComputed(child);
            var box = Box(p);
            switch (child.Kind == "decor" ? p.GetValueOrDefault("type") : child.Kind)
            {
                case "shape":
                    sb.Append("<div class=\"shape\" style=\"").Append(box).Append(ShapeStyle(p)).Append("\">\n");
                    var paragraphs = child.Children.Where(c => Registry.Find(c.Kind)?.Inline != true).ToList();
                    if (paragraphs.Count > 0) Blocks(paragraphs, sb);
                    else if (p.TryGetValue("html", out var html) && html.Length > 0) sb.Append("<p>").Append(html).Append("</p>\n");
                    else if (p.TryGetValue("text", out var text) && text.Length > 0) sb.Append("<p>").Append(Esc(text).Replace("\n", "<br>")).Append("</p>\n");
                    sb.Append("</div>\n");
                    break;
                case "image":
                    sb.Append("<img src=\"").Append(ImageSource(child, p)).Append("\" style=\"").Append(box).Append('"');
                    if (p.TryGetValue("alt", out var alt)) sb.Append(" alt=\"").Append(Esc(alt)).Append('"');
                    sb.Append(">\n");
                    break;
                case "table":
                    Table(child, sb, box);
                    break;
            }
        }
        sb.Append("</section>\n");
    }

    /// <summary>A node's props plus what it inherits (theme colours, layout and master styles), for drawing it as it appears.</summary>
    static IReadOnlyDictionary<string, string> WithComputed(Node node)
    {
        var props = node.GetProps();
        if (node.GetComputed(props) is not { Count: > 0 } computed) return props;
        var all = new Dictionary<string, string>(props);
        foreach (var (name, value) in computed) all.TryAdd(name, value);
        return all;
    }

    static string Box(IReadOnlyDictionary<string, string> p)
    {
        if (!p.ContainsKey("x") || !p.ContainsKey("y")) return "position:static;";
        var css = $"left:{PxText(Px(p["x"]))};top:{PxText(Px(p["y"]))};";
        if (p.TryGetValue("w", out var w)) css += $"width:{PxText(Px(w))};";
        if (p.TryGetValue("h", out var h)) css += $"height:{PxText(Px(h))};";
        return css;
    }

    static string ShapeStyle(IReadOnlyDictionary<string, string> p)
    {
        var css = new StringBuilder();
        if (p.TryGetValue("fill", out var fill) && fill != "none") css.Append("background:#").Append(fill).Append(';');
        if (p.TryGetValue("line", out var line) && line != "none") css.Append("border:1px solid #").Append(line).Append(';');
        if (p.TryGetValue("geometry", out var geometry) && geometry == "ellipse") css.Append("border-radius:50%;");
        else if (geometry == "roundRect") css.Append("border-radius:12px;");
        if (p.TryGetValue("font", out var font))
        {
            css.Append("font-family:'").Append(Esc(font)).Append('\'');
            if (p.TryGetValue("fontEa", out var ea)) css.Append(",'").Append(Esc(ea)).Append('\'');
            css.Append(';');
        }
        if (p.GetValueOrDefault("bold") == "true") css.Append("font-weight:bold;");
        if (p.TryGetValue("size", out var size)) css.Append("font-size:").Append(size).Append("pt;");
        else if (p.TryGetValue("placeholder", out var placeholder))
            css.Append(placeholder switch { "title" => "font-size:32pt;font-weight:bold;", "subtitle" => "font-size:20pt;", _ => "font-size:18pt;" });
        if (p.TryGetValue("color", out var color) && color != "none") css.Append("color:#").Append(color).Append(';');
        return css.ToString();
    }

    public static void Blocks(IReadOnlyList<Node> nodes, StringBuilder sb)
    {
        var i = 0;
        while (i < nodes.Count)
        {
            if (nodes[i].Kind == "paragraph" && nodes[i].GetProps().ContainsKey("list")) i = ListGroup(nodes, i, sb);
            else
            {
                Block(nodes[i], sb);
                i++;
            }
        }
    }

    static int ListGroup(IReadOnlyList<Node> nodes, int start, StringBuilder sb)
    {
        var open = new List<(string Tag, bool ItemOpen)>();
        var i = start;
        for (; i < nodes.Count && nodes[i].Kind == "paragraph" && nodes[i].GetProps().TryGetValue("list", out _); i++)
        {
            var props = nodes[i].GetProps();
            var tag = props["list"] == "bullet" ? "ul" : "ol";
            var level = int.Parse(props.GetValueOrDefault("level") ?? "0", CultureInfo.InvariantCulture);
            while (open.Count > level + 1) Close(open, sb);
            if (open.Count == level + 1 && open[^1].Tag != tag) Close(open, sb);
            while (open.Count < level + 1)
            {
                sb.Append('<').Append(tag).Append(">\n");
                open.Add((tag, false));
            }
            if (open[^1].ItemOpen) sb.Append("</li>\n");
            sb.Append("<li>");
            Inlines(nodes[i], sb);
            open[^1] = (open[^1].Tag, true);
        }
        while (open.Count > 0) Close(open, sb);
        return i;
    }

    static void Close(List<(string Tag, bool ItemOpen)> open, StringBuilder sb)
    {
        var (tag, itemOpen) = open[^1];
        if (itemOpen) sb.Append("</li>\n");
        sb.Append("</").Append(tag).Append(">\n");
        open.RemoveAt(open.Count - 1);
    }

    static void Block(Node node, StringBuilder sb)
    {
        var props = node.Format == "docx" && node.Kind is "paragraph" or "heading" ? WithComputed(node) : node.GetProps();
        switch (node.Kind)
        {
            case "heading":
                var level = Math.Clamp(int.Parse(props.GetValueOrDefault("level") ?? "1", CultureInfo.InvariantCulture), 1, 6);
                sb.Append("<h").Append(level).Append(Style(props, Anchored(node))).Append('>');
                Pictures(node, sb);
                Inlines(node, sb);
                sb.Append("</h").Append(level).Append(">\n");
                break;
            case "paragraph":
                sb.Append("<p").Append(Style(props, Anchored(node))).Append('>');
                Pictures(node, sb);
                Inlines(node, sb);
                sb.Append("</p>\n");
                break;
            case "code":
                sb.Append("<pre><code");
                if (props.TryGetValue("lang", out var lang)) sb.Append(" class=\"language-").Append(Esc(lang)).Append('"');
                sb.Append('>').Append(Esc(props.GetValueOrDefault("text") ?? "")).Append("</code></pre>\n");
                break;
            case "pagebreak":
                sb.Append("<div class=\"pagebreak\" style=\"page-break-after:always\"></div>\n");
                break;
            case "toc":
                sb.Append("<nav class=\"toc\">\n");
                if (props.TryGetValue("title", out var tocTitle)) sb.Append("<p class=\"toc-title\">").Append(Esc(tocTitle)).Append("</p>\n");
                sb.Append("<ul>\n");
                foreach (var line in (props.GetValueOrDefault("text") ?? "").Split('\n', StringSplitOptions.RemoveEmptyEntries))
                {
                    var tab = line.LastIndexOf('\t');
                    sb.Append("<li><span>").Append(Esc(tab < 0 ? line : line[..tab])).Append("</span>");
                    if (tab >= 0) sb.Append("<span class=\"toc-page\">").Append(Esc(line[(tab + 1)..])).Append("</span>");
                    sb.Append("</li>\n");
                }
                sb.Append("</ul>\n</nav>\n");
                break;
            case "table":
                Table(node, sb, null);
                break;
            case "image":
                sb.Append("<img src=\"").Append(ImageSource(node, props)).Append('"');
                if (props.TryGetValue("alt", out var alt)) sb.Append(" alt=\"").Append(Esc(alt)).Append('"');
                if (props.TryGetValue("width", out var width) && long.TryParse(width, out var emu)) sb.Append(" style=\"width:").Append(Math.Round(emu / (double)Units.EmuPerPx)).Append("px\"");
                sb.Append(">\n");
                break;
            case "page":
                sb.Append("<section class=\"page\">\n");
                Blocks(node.Children.Where(c => Registry.Find(c.Kind)?.Inline != true).ToList(), sb);
                sb.Append("</section>\n");
                break;
            case "chart":
                sb.Append("<figure class=\"chart\">").Append(Esc(props.GetValueOrDefault("title") ?? "Chart")).Append("</figure>\n");
                break;
            case "topic":
                sb.Append("<ul class=\"mindmap\">\n");
                Topic(node, sb);
                sb.Append("</ul>\n");
                break;
            default:
                var children = node.Children.Where(c => Registry.Find(c.Kind)?.Inline != true).ToList();
                if (children.Count > 0) Blocks(children, sb);
                else if (props.TryGetValue("text", out var text) && text.Length > 0)
                    sb.Append("<p>").Append(Esc(text).Replace("\n", "<br>")).Append("</p>\n");
                break;
        }
    }

    /// <summary>A mind map topic as a list item holding its children as a nested list.</summary>
    static void Topic(Node topic, StringBuilder sb)
    {
        var text = Esc(topic.Text ?? "").Replace("\n", "<br>");
        sb.Append("<li>").Append(topic.GetProps().TryGetValue("link", out var link) ? $"<a href=\"{SafeUrl(link)}\">{text}</a>" : text);
        var children = topic.Children;
        if (children.Count > 0)
        {
            sb.Append("<ul>\n");
            foreach (var child in children) Topic(child, sb);
            sb.Append("</ul>\n");
        }
        sb.Append("</li>\n");
    }

    static void Table(Node table, StringBuilder sb, string? box)
    {
        var tableProps = table.Kind == "table" ? table.GetProps() : new Dictionary<string, string>();
        var css = new StringBuilder(box ?? "");
        if (tableProps.TryGetValue("width", out var tableWidth))
            css.Append("width:").Append(tableWidth.EndsWith('%') ? tableWidth : PxText(Px(Units.ParseLength(tableWidth).ToString(CultureInfo.InvariantCulture)))).Append(';');
        if (tableProps.TryGetValue("align", out var tableAlign) && tableAlign != "left")
            css.Append(tableAlign == "center" ? "margin-left:auto;margin-right:auto;" : "margin-left:auto;");
        var borderColor = tableProps.TryGetValue("borderColor", out var bc) ? "#" + bc : table.Format == "docx" ? "#000" : "#bbb"; // Word's automatic line is black
        var borders = tableProps.GetValueOrDefault("borders");
        if (borders is "all" or "outside") css.Append("border:1px solid ").Append(borderColor).Append(';');
        else if (borders is "inside") css.Append("border-style:hidden;"); // collapsed borders: hides the cells' outer edges only
        var cellBorder = borders switch
        {
            "none" or "outside" => "border:none",
            "horizontal" => "border-left:none;border-right:none;border-color:" + borderColor,
            _ => tableProps.ContainsKey("borderColor") ? "border-color:" + borderColor : null,
        };
        var widths = table.Format == "docx" && tableProps.TryGetValue("widths", out var w) ? Lengths(w) : [];
        if (widths.Count > 0) css.Append("table-layout:fixed;");
        sb.Append(css.Length == 0 ? "<table>\n" : $"<table style=\"{css}\">\n");
        if (widths.Count > 0) sb.Append("<colgroup>").Append(string.Concat(widths.Select(x => $"<col style=\"width:{x}\">"))).Append("</colgroup>\n");
        var rows = table.Children.Where(c => c.Kind == "row").ToList();
        for (var r = 0; r < rows.Count; r++)
        {
            var tag = r == 0 && table.Format == "md" ? "th" : "td";
            sb.Append("<tr>");
            foreach (var cell in rows[r].Children)
            {
                var props = cell.GetProps();
                sb.Append('<').Append(tag);
                if (props.TryGetValue("colspan", out var colspan)) sb.Append(" colspan=\"").Append(colspan).Append('"');
                if (props.TryGetValue("rowspan", out var rowspan)) sb.Append(" rowspan=\"").Append(rowspan).Append('"');
                sb.Append(CellStyle(props, cellBorder, borderColor)).Append('>');
                var blocks = cell.Children.Where(c => Registry.Find(c.Kind)?.Inline != true).ToList();
                if (blocks.Count == 1 && blocks[0].Kind == "paragraph" && !blocks[0].GetProps().ContainsKey("list")) Inlines(blocks[0], sb);
                else if (blocks.Count > 0)
                {
                    var inner = new StringBuilder();
                    Blocks(blocks, inner);
                    sb.Append(inner.ToString().Trim());
                }
                else Inlines(cell, sb);
                sb.Append("</").Append(tag).Append('>');
            }
            sb.Append("</tr>\n");
        }
        sb.Append("</table>\n");
    }

    static string ImageSource(Node node, IReadOnlyDictionary<string, string> props)
    {
        if (node.GetBinary() is { } binary)
            return "data:" + binary.ContentType + ";base64," + Convert.ToBase64String(binary.Data);
        return SafeUrl(props.GetValueOrDefault("src") ?? "");
    }

    /// <summary>Only web, mail, anchor and relative targets reach the page; script and data schemes become "#".</summary>
    static string SafeUrl(string url)
    {
        var t = url.TrimStart();
        var ok = t.StartsWith("http:", StringComparison.OrdinalIgnoreCase)
            || t.StartsWith("https:", StringComparison.OrdinalIgnoreCase)
            || t.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
            || t.StartsWith('#')
            || !t.Contains(':');
        return ok ? Esc(url) : "#";
    }

    static void Inlines(Node node, StringBuilder sb)
    {
        var runs = node.Children.Where(c => c.Kind == "run").ToList();
        if (runs.Count == 0)
        {
            sb.Append(Esc(node.Text ?? "").Replace("\n", "<br>"));
            return;
        }
        // docx notes, while a document's pages are drawn: each mark goes in at its character offset of the (undeleted) text
        var marks = _notes is null ? [] : node.Children.Where(c => c.Kind == "footnote").Select(c => c.GetProps())
            .OrderBy(n => int.TryParse(n.GetValueOrDefault("at"), out var at) ? at : int.MaxValue).ToList();
        var (pos, next) = (0, 0);
        int At(int i) => int.TryParse(marks[i].GetValueOrDefault("at"), out var at) ? at : int.MaxValue;
        foreach (var run in runs)
        {
            var p = run.Format == "docx" ? WithComputed(run) : run.GetProps();
            var open = new StringBuilder();
            var close = new StringBuilder();
            if (p.TryGetValue("change", out var change))
            {
                var tag = change == "deleted" ? "del" : "ins";
                open.Append('<').Append(tag).Append('>');
                close.Insert(0, "</" + tag + ">");
            }
            if (p.TryGetValue("link", out var link))
            {
                open.Append("<a href=\"").Append(SafeUrl(link)).Append("\">");
                close.Insert(0, "</a>");
            }
            var css = new List<string>();
            if (p.TryGetValue("color", out var color) && color != "none") css.Add("color:#" + color);
            if (p.TryGetValue("size", out var size)) css.Add("font-size:" + size + "pt");
            if (p.TryGetValue("font", out var font)) css.Add("font-family:'" + Esc(font) + "'" + (p.TryGetValue("fontEa", out var fontEa) ? ",'" + Esc(fontEa) + "'" : ""));
            if (p.ContainsKey("underline")) css.Add("text-decoration:underline");
            if (css.Count > 0)
            {
                open.Append("<span style=\"").Append(string.Join(';', css)).Append("\">");
                close.Insert(0, "</span>");
            }
            Wrap(p, "bold", "strong", open, close);
            Wrap(p, "italic", "em", open, close);
            Wrap(p, "strike", "s", open, close);
            Wrap(p, "code", "code", open, close);
            if (p.GetValueOrDefault("vertAlign") is { } vertical && vertical is "superscript" or "subscript") { var tag = vertical == "superscript" ? "sup" : "sub"; open.Append('<').Append(tag).Append('>'); close.Insert(0, "</" + tag + ">"); }
            var text = p.GetValueOrDefault("text") ?? "";
            sb.Append(open);
            var from = 0;
            if (p.GetValueOrDefault("change") != "deleted")
            {
                for (; next < marks.Count && At(next) <= pos + text.Length; next++)
                {
                    var cut = Math.Clamp(At(next) - pos, from, text.Length);
                    sb.Append(RunText(text[from..cut])).Append(NoteMark(marks[next]));
                    from = cut;
                }
                pos += text.Length;
            }
            sb.Append(RunText(text[from..])).Append(close);
        }
        for (; next < marks.Count; next++) sb.Append(NoteMark(marks[next]));
    }

    static string RunText(string text) => Esc(text).Replace("\n", "<br>").Replace("\f", Common.InlineHtml.PageBreak);

    static void Wrap(IReadOnlyDictionary<string, string> props, string prop, string tag, StringBuilder open, StringBuilder close)
    {
        if (!props.TryGetValue(prop, out var v) || v != "true") return;
        open.Append('<').Append(tag).Append('>');
        close.Insert(0, "</" + tag + ">");
    }

    static string Style(IReadOnlyDictionary<string, string> props, string? extra = null)
    {
        var css = new List<string>();
        if (extra is not null) css.Add(extra);
        if (props.TryGetValue("align", out var align)) css.Add("text-align:" + align);
        if (props.TryGetValue("fill", out var fill) && fill != "none") css.AddRange(Fill(fill));
        css.AddRange(ParaCss(props)); // docx: the paragraph's own 段落 settings and what its style gives beyond the document's look
        css.AddRange(TextCss(props));
        return css.Count == 0 ? "" : " style=\"" + string.Join(';', css) + "\"";
    }

    /// <summary>Cell CSS: alignment and fill like any block, plus the table's line rule, the cell's own borders, vertical alignment and width.</summary>
    static string CellStyle(IReadOnlyDictionary<string, string> props, string? tableBorder, string lineColor)
    {
        var css = new List<string>();
        if (props.TryGetValue("align", out var align)) css.Add("text-align:" + align);
        if (props.TryGetValue("fill", out var fill) && fill != "none") css.AddRange(Fill(fill));
        if (tableBorder is not null) css.Add(tableBorder);
        if (props.TryGetValue("borders", out var borders)) css.Add(borders == "none" ? "border:none" : "border:1px solid " + lineColor);
        if (props.TryGetValue("valign", out var valign)) css.Add("vertical-align:" + valign);
        if (props.TryGetValue("width", out var width) && long.TryParse(width, NumberStyles.Integer, CultureInfo.InvariantCulture, out _)) css.Add("width:" + PxText(Px(width)));
        return css.Count == 0 ? "" : " style=\"" + string.Join(';', css) + "\"";
    }

    /// <summary>A fill and, on a dark one, white for the automatic text: as in Word, automatic text follows the colour behind it.</summary>
    static IEnumerable<string> Fill(string rrggbb)
    {
        yield return "background:#" + rrggbb;
        var rgb = int.TryParse(rrggbb, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var v) ? v : 0xFFFFFF;
        if ((((rgb >> 16) & 255) * 299 + ((rgb >> 8) & 255) * 587 + (rgb & 255) * 114) / 1000 < 128) yield return "color:#FFFFFF";
    }

    static string Esc(string s) => WebUtility.HtmlEncode(s);
}
