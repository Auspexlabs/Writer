using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace Writer.Formats.Compat;

/// <summary>HTML files → blocks (or sheets, one per table, for a spreadsheet target: many systems export "Excel" files
/// that are HTML tables). A tolerant tag walker, no DOM: block tags open and close blocks, inline tags are passed
/// through as the inline html the engine understands, everything else is unwrapped.</summary>
public static class HtmlReader
{
    static readonly Regex Token = new(@"<!--.*?-->|<!\[CDATA\[.*?\]\]>|<[!?/]?[a-zA-Z][^>]*>|<[^>]*>|[^<]+", RegexOptions.Singleline | RegexOptions.Compiled);
    static readonly Regex Attr = new(@"([a-zA-Z_:][-a-zA-Z0-9_:.]*)\s*(?:=\s*(?:""([^""]*)""|'([^']*)'|([^\s""'=<>`]+)))?", RegexOptions.Compiled);
    static readonly HashSet<string> Void = ["br", "img", "hr", "meta", "link", "input", "col", "area", "base", "wbr", "source", "embed", "param", "track"];
    static readonly HashSet<string> BlockTags = ["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote", "pre", "section", "article", "header", "footer", "main", "aside", "nav", "figure", "figcaption", "dd", "dt", "center", "address"];

    public static bool LooksLikeHtml(ReadOnlySpan<char> head)
    {
        if (head.Length == 0 || head[0] != '<') return false;
        var probe = head[..Math.Min(head.Length, 4000)].ToString().ToLowerInvariant();
        return probe.StartsWith("<!doctype html") || probe.StartsWith("<html") || probe.Contains("<html") || probe.Contains("<body") || probe.Contains("<table") || probe.Contains("<div") || probe.Contains("<p>") || probe.Contains("<p ") || probe.Contains("<br");
    }

    public static List<Block> Read(string html, List<string> warnings)
    {
        var w = new Walker(warnings);
        w.Run(html);
        return w.Blocks;
    }

    /// <summary>One sheet per table (the first row's cells become the header row as text); no table → the text as one column.</summary>
    public static List<SheetModel> Sheets(string html, List<string> warnings)
    {
        var blocks = Read(html, warnings);
        var sheets = new List<SheetModel>();
        foreach (var t in blocks.Where(b => b.Kind == BlockKind.Table && b.Rows is { Count: > 0 }))
        {
            var sheet = new SheetModel { Name = sheets.Count == 0 ? "Sheet1" : $"Sheet{sheets.Count + 1}" };
            var reserved = new HashSet<(int, int)>();
            for (var r = 0; r < t.Rows!.Count; r++)
            {
                var c = 0;
                foreach (var cell in t.Rows[r])
                {
                    while (reserved.Contains((r, c))) c++;
                    var text = Writers.Plain(cell.Html).Trim();
                    if (text.Length > 0)
                    {
                        var value = TextFiles.Typed(text);
                        sheet.Cells.Add(new CellModel { Row = r + 1, Col = c + 1, Value = value, Fill = cell.Fill, Bold = cell.Html.Contains("<b>"), Format = value is DateTime ? "yyyy-mm-dd" : null });
                    }
                    else if (cell.Fill is not null) sheet.Cells.Add(new CellModel { Row = r + 1, Col = c + 1, Fill = cell.Fill });
                    if (cell.ColSpan > 1 || cell.RowSpan > 1)
                    {
                        sheet.Merges.Add($"{Xlsx.XlsxCells.Reference(c + 1, r + 1)}:{Xlsx.XlsxCells.Reference(c + Math.Max(1, cell.ColSpan), r + Math.Max(1, cell.RowSpan))}");
                        for (var dr = 0; dr < Math.Max(1, cell.RowSpan); dr++) for (var dc = 0; dc < Math.Max(1, cell.ColSpan); dc++) reserved.Add((r + dr, c + dc));
                    }
                    c += Math.Max(1, cell.ColSpan);
                }
            }
            sheets.Add(sheet);
        }
        if (sheets.Count == 0)
        {
            var sheet = new SheetModel();
            var row = 1;
            foreach (var b in blocks.Where(b => b.Kind is BlockKind.Paragraph or BlockKind.Heading))
            {
                var text = Writers.Plain(b.Html).Trim();
                if (text.Length > 0) sheet.Cells.Add(new CellModel { Row = row, Col = 1, Value = text });
                row++;
            }
            sheets.Add(sheet);
        }
        return sheets;
    }

    sealed class Walker(List<string> warnings)
    {
        public readonly List<Block> Blocks = [];
        readonly StringBuilder _inline = new();
        readonly List<string> _open = [];          // inline tags open right now, as emitted
        readonly List<string> _lists = [];         // ul / ol nesting
        Block? _block;                              // the block being filled (null = none started yet)
        int _quote, _pre;
        // tables: the innermost one being read
        readonly Stack<(List<List<TableCell>> Rows, List<TableCell>? Row, TableCell? Cell, StringBuilder Text)> _tables = new();

        public void Run(string html)
        {
            string? skipUntil = null;
            foreach (Match m in Token.Matches(html))
            {
                var tok = m.Value;
                if (tok.StartsWith("<!") || tok.StartsWith("<?")) continue;
                if (tok[0] == '<' && tok.Length > 1)
                {
                    var close = tok[1] == '/';
                    var body = tok.Substring(close ? 2 : 1, tok.Length - (close ? 3 : 2)).Trim();
                    var nameEnd = body.IndexOfAny([' ', '\t', '\r', '\n', '/']);
                    var name = (nameEnd < 0 ? body : body[..nameEnd]).ToLowerInvariant();
                    if (name.Length == 0 || !char.IsLetter(name[0])) continue;
                    if (skipUntil is not null) { if (close && name == skipUntil) skipUntil = null; continue; }
                    if (!close && name is "script" or "style" or "head" or "title" or "noscript" or "template") { skipUntil = name; continue; }
                    var attrs = nameEnd < 0 ? new Dictionary<string, string>() : Attrs(body[nameEnd..]);
                    if (close) Close(name); else Open(name, attrs, tok.EndsWith("/>") || Void.Contains(name));
                }
                else if (skipUntil is null) Text(tok);
            }
            Flush();
            while (_tables.Count > 0) EndTable();
        }

        static Dictionary<string, string> Attrs(string s)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match m in Attr.Matches(s))
                d[m.Groups[1].Value.ToLowerInvariant()] = WebUtility.HtmlDecode(m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Success ? m.Groups[3].Value : m.Groups[4].Value);
            return d;
        }

        StringBuilder Target => _tables.Count > 0 && _tables.Peek().Cell is not null ? _tables.Peek().Text : _inline;

        void Text(string raw)
        {
            var text = WebUtility.HtmlDecode(raw);
            if (_pre == 0)
            {
                text = Regex.Replace(text, @"[ \t\r\n]+", " ");
                if (text == " " && (Target.Length == 0 || Target[^1] == ' ' || Target.ToString().EndsWith("<br>"))) return;
            }
            else text = text.Replace("\r\n", "\n");
            if (_pre > 0) { var lines = text.Split('\n'); for (var i = 0; i < lines.Length; i++) { if (i > 0) Target.Append("<br>"); Target.Append(Inline.Esc(lines[i])); } }
            else Target.Append(Inline.Esc(text));
        }

        void Open(string name, Dictionary<string, string> a, bool selfClosing)
        {
            switch (name)
            {
                case "br": Target.Append("<br>"); return;
                case "hr": Flush(); return;
                case "img":
                    if (a.TryGetValue("src", out var src) && src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            var (bytes, _) = Common.ImageInfo.Load(src);
                            var img = Block.Picture(bytes, Px(a.GetValueOrDefault("width")), Px(a.GetValueOrDefault("height")));
                            if (_tables.Count > 0) { warnings.Add("picture inside a table cell dropped"); return; }
                            Flush(); Blocks.Add(img);
                        }
                        catch (Exception ex) when (ex is Core.WriterException or FormatException) { warnings.Add($"picture dropped: {ex.Message}"); }
                    }
                    else warnings.Add($"picture {a.GetValueOrDefault("src") ?? ""} dropped: only embedded (data:) pictures are read");
                    return;
                case "table":
                    if (_tables.Count == 0) Flush();
                    _tables.Push(([], null, null, new StringBuilder()));
                    return;
                case "tr" when _tables.Count > 0: EndCell(); var t = _tables.Pop(); t.Row = []; t.Rows.Add(t.Row); _tables.Push(t); return;
                case "td" or "th" when _tables.Count > 0:
                    EndCell();
                    var tb = _tables.Pop();
                    if (tb.Row is null) { tb.Row = []; tb.Rows.Add(tb.Row); }
                    var cell = new TableCell { ColSpan = Math.Max(1, Int(a.GetValueOrDefault("colspan"))), RowSpan = Math.Max(1, Int(a.GetValueOrDefault("rowspan"))), Fill = Color(a.GetValueOrDefault("bgcolor")) ?? StyleColor(a.GetValueOrDefault("style"), "background-color") ?? StyleColor(a.GetValueOrDefault("style"), "background") };
                    tb.Row.Add(cell); tb.Cell = cell; tb.Text.Clear();
                    _tables.Push(tb);
                    if (name == "th") tb.Text.Append("<b>");
                    return;
                case "ul" or "ol": Flush(); _lists.Add(name); return;
                case "blockquote": Flush(); _quote++; return;
                case "pre": Flush(); _pre++; return;
                case "b" or "strong": Push("<b>"); return;
                case "i" or "em" or "cite" or "dfn" or "var": Push("<i>"); return;
                case "u" or "ins": Push("<u>"); return;
                case "s" or "strike" or "del": Push("<s>"); return;
                case "code" or "kbd" or "samp" or "tt": Push("<span style=\"font-family:Courier New\">"); return;
                case "a":
                    if (a.TryGetValue("href", out var href) && href.Length > 0 && !href.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase)) Push($"<a href=\"{Inline.Esc(href)}\">");
                    else Push("<span>");
                    return;
                case "span" or "font" or "small" or "big" or "mark" or "sub" or "sup" or "label" or "abbr" or "time" or "q":
                    var style = new StringBuilder();
                    if (a.TryGetValue("style", out var st))
                        foreach (var k in new[] { "color", "background-color", "font-size", "font-family", "font-weight", "font-style", "text-decoration" })
                            if (StyleValue(st, k) is { } v) style.Append(k).Append(':').Append(v).Append(';');
                    if (name == "font")
                    {
                        if (Color(a.GetValueOrDefault("color")) is { } fc) style.Append("color:#").Append(fc).Append(';');
                        if (a.TryGetValue("face", out var face)) style.Append("font-family:").Append(face.Split(',')[0].Trim()).Append(';');
                        if (int.TryParse(a.GetValueOrDefault("size"), out var fs)) style.Append("font-size:").Append(new[] { 8, 10, 12, 14, 18, 24, 36 }[Math.Clamp(fs, 1, 7) - 1]).Append("pt;");
                    }
                    if (name == "mark") style.Append("background-color:#FFFF00;");
                    Push(style.Length > 0 ? $"<span style=\"{style}\">" : "<span>");
                    return;
            }
            if (BlockTags.Contains(name))
            {
                if (_tables.Count > 0) { if (_tables.Peek().Text.Length > 0) _tables.Peek().Text.Append("<br>"); return; }
                Flush();
                var b = name[0] == 'h' && name.Length == 2 && char.IsDigit(name[1]) ? Block.Heading(name[1] - '0', "") : Block.Paragraph("");
                if (name == "li" && _lists.Count > 0) { b.List = _lists[^1] == "ol" ? "number" : "bullet"; b.Level = Math.Min(8, _lists.Count - 1); }
                if (_quote > 0 && b.Kind == BlockKind.Paragraph) b.Style = "Quote";
                b.Align = Align(a.GetValueOrDefault("align")) ?? Align(StyleValue(a.GetValueOrDefault("style"), "text-align"));
                if (_pre > 0) Push("<span style=\"font-family:Courier New\">");
                _block = b;
                return;
            }
            // anything else is unwrapped: its text lands in the current block
        }

        void Close(string name)
        {
            switch (name)
            {
                case "td" or "th" when _tables.Count > 0: EndCell(); return;
                case "tr" when _tables.Count > 0: EndCell(); return;
                case "table" when _tables.Count > 0: EndTable(); return;
                case "ul" or "ol": Flush(); if (_lists.Count > 0) _lists.RemoveAt(_lists.Count - 1); return;
                case "blockquote": Flush(); _quote = Math.Max(0, _quote - 1); return;
                case "pre": Flush(); _pre = Math.Max(0, _pre - 1); return;
                case "b" or "strong": Pop("</b>"); return;
                case "i" or "em" or "cite" or "dfn" or "var": Pop("</i>"); return;
                case "u" or "ins": Pop("</u>"); return;
                case "s" or "strike" or "del": Pop("</s>"); return;
                case "a": Pop("</a>"); return;
                case "code" or "kbd" or "samp" or "tt" or "span" or "font" or "small" or "big" or "mark" or "sub" or "sup" or "label" or "abbr" or "time" or "q": Pop("</span>"); return;
            }
            if (BlockTags.Contains(name)) { if (_tables.Count == 0) Flush(); }
        }

        void Push(string tag) { _open.Add(tag); Target.Append(tag); }

        void Pop(string closing)
        {
            var open = closing == "</a>" ? "<a " : closing == "</span>" ? "<span" : closing.Replace("/", "");
            for (var i = _open.Count - 1; i >= 0; i--)
                if (_open[i].StartsWith(open)) { _open.RemoveAt(i); break; }
            Target.Append(closing);
        }

        /// <summary>Ends the block being filled; the inline tags still open continue in the next one, as in a browser.</summary>
        void Flush()
        {
            if (_tables.Count > 0) return;
            var html = _inline.ToString();
            foreach (var tag in Enumerable.Reverse(_open)) html += Closing(tag);
            html = html.Trim();
            if (_block is not null || Writers.Plain(html).Trim().Length > 0)
            {
                var b = _block ?? Block.Paragraph("");
                if (_block is null && _quote > 0) b.Style = "Quote"; // bare text inside a blockquote
                b.Html = html;
                Blocks.Add(b);
            }
            _block = null;
            _inline.Clear();
            foreach (var tag in _open) _inline.Append(tag);
        }

        static string Closing(string tag) => tag.StartsWith("<a ") ? "</a>" : tag.StartsWith("<span") ? "</span>" : tag.Insert(1, "/");

        void EndCell()
        {
            var t = _tables.Pop();
            if (t.Cell is not null)
            {
                var text = t.Text.ToString();
                if (text.StartsWith("<b>") && !text.Contains("</b>")) text += "</b>";
                t.Cell.Html = text.Trim();
                t.Cell = null;
                t.Text.Clear();
            }
            _tables.Push(t);
        }

        void EndTable()
        {
            EndCell();
            var t = _tables.Pop();
            var rows = t.Rows.Where(r => r.Count > 0).ToList();
            if (_tables.Count > 0)
            {
                // a nested table: its cells' text joins the outer cell
                var outer = _tables.Peek();
                foreach (var r in rows) { if (outer.Text.Length > 0) outer.Text.Append("<br>"); outer.Text.Append(string.Join(" | ", r.Select(c => c.Html))); }
                return;
            }
            if (rows.Count > 0) Blocks.Add(Block.Table(rows));
        }

        static int Int(string? s) => int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : 0;
        static double? Px(string? s) => double.TryParse(s?.Replace("px", "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) && v > 0 ? v / 96 * 2.54 : null;
        static string? Align(string? s) => s?.ToLowerInvariant() switch { "center" => "center", "right" or "end" => "right", "justify" => "justify", "left" or "start" => "left", _ => null };

        static string? StyleValue(string? style, string key)
        {
            if (style is null) return null;
            foreach (var part in style.Split(';'))
            {
                var colon = part.IndexOf(':');
                if (colon > 0 && part[..colon].Trim().Equals(key, StringComparison.OrdinalIgnoreCase)) return part[(colon + 1)..].Trim();
            }
            return null;
        }

        static string? StyleColor(string? style, string key) => Color(StyleValue(style, key));

        static readonly Dictionary<string, string> Named = new(StringComparer.OrdinalIgnoreCase)
        {
            ["black"] = "000000", ["white"] = "FFFFFF", ["red"] = "FF0000", ["green"] = "008000", ["blue"] = "0000FF", ["yellow"] = "FFFF00", ["gray"] = "808080", ["grey"] = "808080",
            ["silver"] = "C0C0C0", ["orange"] = "FFA500", ["purple"] = "800080", ["navy"] = "000080", ["teal"] = "008080", ["maroon"] = "800000", ["olive"] = "808000", ["lime"] = "00FF00", ["aqua"] = "00FFFF", ["fuchsia"] = "FF00FF",
        };

        /// <summary>RRGGBB from #rgb, #rrggbb, rgb(r,g,b) or a common name; null for anything else (transparent, inherit).</summary>
        static string? Color(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            s = s.Trim();
            if (Named.TryGetValue(s, out var named)) return named;
            if (s.StartsWith('#')) s = s[1..];
            if (s.Length == 3 && s.All(Uri.IsHexDigit)) return string.Concat(s.Select(c => new string(c, 2))).ToUpperInvariant();
            if (s.Length == 6 && s.All(Uri.IsHexDigit)) return s.ToUpperInvariant();
            var m = Regex.Match(s, @"^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)");
            return m.Success ? Inline.Rgb(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), int.Parse(m.Groups[3].Value)) : null;
        }
    }
}
