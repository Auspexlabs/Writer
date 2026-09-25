using System.Globalization;
using System.Text;
using System.Text.Json;
using Writer.Core;
using Writer.Formats.Xlsx;

namespace Writer.Formats.Html;

public static partial class HtmlWriter
{
    const string SheetCss = """
        body.book{max-width:none}
        .sheet{position:relative;margin:0 0 2em 0}.sheet svg{position:absolute;overflow:visible}
        .grid{border-collapse:collapse;table-layout:fixed;margin:0;line-height:1.2}
        .grid td{border:none;padding:0 3px;overflow:hidden;white-space:nowrap;vertical-align:bottom;text-align:left}
        .grid.lines td{box-shadow:inset -1px -1px #d4d4d4}
        """;

    static readonly Dictionary<string, string> BorderCss = new()
    {
        ["thin"] = "1px solid", ["hair"] = "1px solid", ["medium"] = "2px solid", ["thick"] = "3px solid", ["double"] = "3px double", ["dashed"] = "1px dashed",
        ["dotted"] = "1px dotted", ["mediumDashed"] = "2px dashed", ["dashDot"] = "1px dashed", ["mediumDashDot"] = "2px dashed", ["dashDotDot"] = "1px dotted",
        ["mediumDashDotDot"] = "2px dotted", ["slantDashDot"] = "2px dashed",
    };

    sealed class Book(Document doc)
    {
        readonly Dictionary<string, Dictionary<(int C, int R), IReadOnlyDictionary<string, string>>> _cells = new(StringComparer.OrdinalIgnoreCase);

        public Dictionary<(int C, int R), IReadOnlyDictionary<string, string>> Cells(Node sheet)
        {
            var name = sheet.Name ?? "";
            if (_cells.TryGetValue(name, out var known)) return known;
            var cells = new Dictionary<(int C, int R), IReadOnlyDictionary<string, string>>();
            foreach (var row in sheet.Children.Where(c => c.Kind == "row"))
                foreach (var cell in row.Children)
                    if (cell.Key is { } key) cells[XlsxCells.Parse(key)] = cell.GetProps();
            return _cells[name] = cells;
        }

        /// <summary>The cells of "B2:B5" or "'Other sheet'!B2:B5", row by row; null entries are empty cells. A chart's constant array
        /// ({1,2,3} or {"a","b"}) gives its items as cells.</summary>
        public List<IReadOnlyDictionary<string, string>?> Range(Node sheet, string reference)
        {
            if (reference.TrimStart().StartsWith('{'))
                return System.Text.RegularExpressions.Regex.Matches(reference.Trim().Trim('{', '}'), "\"(?:[^\"]|\"\")*\"|[^,;]+")
                    .Select(m => m.Value.Trim()).Select(v => (IReadOnlyDictionary<string, string>?)(v.StartsWith('"')
                        ? new Dictionary<string, string> { ["value"] = v[1..^1].Replace("\"\"", "\""), ["type"] = "string" }
                        : new Dictionary<string, string> { ["value"] = v, ["type"] = "number" })).ToList();
            var bang = reference.LastIndexOf('!');
            if (bang >= 0)
            {
                var name = reference[..bang].Trim('\'').Replace("''", "'");
                sheet = doc.Root.Children.FirstOrDefault(s => s.Kind == "sheet" && string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)) ?? sheet;
            }
            var list = new List<IReadOnlyDictionary<string, string>?>();
            try
            {
                var (c1, r1, c2, r2) = XlsxCells.ParseRange(reference[(bang + 1)..]);
                if ((long)(c2 - c1 + 1) * (r2 - r1 + 1) > 10000) return list;
                var cells = Cells(sheet);
                for (var r = r1; r <= r2; r++)
                    for (var c = c1; c <= c2; c++) list.Add(cells.GetValueOrDefault((c, r)));
            }
            catch (WriterException) { }
            return list;
        }

        public string Accent(int i) => doc is XlsxDocument x ? x.Styles.Colors.Accent(i) : "4472C4";
    }

    /// <summary>A worksheet as Excel draws it: column widths and row heights, merges, hidden lines and grid lines, each cell's fill,
    /// font, borders and alignment, numbers and dates through their format codes, and the sheet's charts where they sit.</summary>
    static void Sheet(Node sheet, Book book, IReadOnlyDictionary<string, string> rootProps, StringBuilder sb)
    {
        var p = sheet.GetProps();
        sb.Append("<h2>").Append(Esc(p.GetValueOrDefault("name") ?? "Sheet")).Append("</h2>\n");
        var cells = book.Cells(sheet);
        var merges = Json(p, "merges").EnumerateArray().Select(m => XlsxCells.ParseRange(m.GetString() ?? "A1")).ToList();
        var hidden = Json(p, "hidden");
        var hiddenRows = hidden.TryGetProperty("rows", out var hr) ? hr.EnumerateArray().Select(r => r.GetInt32()).ToHashSet() : [];
        var hiddenCols = hidden.TryGetProperty("cols", out var hc) ? hc.EnumerateArray().Select(c => XlsxCells.ColumnIndex(c.GetString() ?? "A")).ToHashSet() : [];
        var widths = Json(p, "widths");
        var heights = Json(p, "heights");
        int ColPx(int c) => hiddenCols.Contains(c) ? 0 : widths.TryGetProperty(XlsxCells.ColumnName(c), out var w) ? (int)Math.Truncate((256 * w.GetDouble() + 18) / 256 * 7) : 64;
        int RowPx(int r) => hiddenRows.Contains(r) ? 0 : heights.TryGetProperty(r.ToString(CultureInfo.InvariantCulture), out var h) ? (int)Math.Round(h.GetDouble() * 96 / 72) : 20;
        int maxC = 0, maxR = 0;
        foreach (var ((c, r), cp) in cells)
            if (cp.GetValueOrDefault("value") is { Length: > 0 } || cp.ContainsKey("fill") || cp.ContainsKey("border")) (maxC, maxR) = (Math.Max(maxC, c), Math.Max(maxR, r));
        foreach (var m in merges) (maxC, maxR) = (Math.Max(maxC, m.Col2), Math.Max(maxR, m.Row2));
        (maxC, maxR) = (Math.Min(maxC, 256), Math.Min(maxR, 5000)); // ponytail: a page, not the whole grid
        var charts = sheet.Children.Where(c => c.Kind == "chart").Select(c => c.GetProps()).ToList();
        double tableW = Enumerable.Range(1, maxC).Sum(ColPx);
        var boxW = charts.Select(ch => Px(ch.GetValueOrDefault("x")) + Px(ch.GetValueOrDefault("w"))).Append(tableW).Max();
        var boxH = charts.Select(ch => Px(ch.GetValueOrDefault("y")) + Px(ch.GetValueOrDefault("h"))).Append((double)Enumerable.Range(1, maxR).Sum(RowPx)).Max();
        sb.Append("<div class=\"sheet\" style=\"width:").Append(PxText(boxW)).Append(";height:").Append(PxText(boxH)).Append("\">\n");
        sb.Append("<table class=\"grid").Append(p.GetValueOrDefault("gridlines") == "false" ? "" : " lines").Append("\" style=\"width:").Append(PxText(tableW)).Append(";font-family:")
            .Append(FontStack(rootProps.GetValueOrDefault("font") ?? "Calibri")).Append(";font-size:").Append(rootProps.GetValueOrDefault("size") ?? "11").Append("pt\">\n<colgroup>");
        for (var c = 1; c <= maxC; c++)
            if (ColPx(c) > 0) sb.Append("<col style=\"width:").Append(ColPx(c)).Append("px\">");
        sb.Append("</colgroup>\n");
        var covered = new HashSet<(int, int)>();
        foreach (var m in merges)
            for (var r = m.Row1; r <= m.Row2; r++)
                for (var c = m.Col1; c <= m.Col2; c++)
                    if (r != m.Row1 || c != m.Col1) covered.Add((c, r));
        for (var r = 1; r <= maxR; r++)
        {
            if (hiddenRows.Contains(r)) continue;
            sb.Append("<tr style=\"height:").Append(RowPx(r)).Append("px\">");
            for (var c = 1; c <= maxC; c++)
            {
                if (hiddenCols.Contains(c) || covered.Contains((c, r))) continue;
                sb.Append("<td");
                if (merges.FirstOrDefault(x => x.Col1 == c && x.Row1 == r) is { Col2: > 0 } merge)
                {
                    var cs = Enumerable.Range(merge.Col1, merge.Col2 - merge.Col1 + 1).Count(x => !hiddenCols.Contains(x));
                    var rs = Enumerable.Range(merge.Row1, merge.Row2 - merge.Row1 + 1).Count(x => !hiddenRows.Contains(x));
                    if (cs > 1) sb.Append(" colspan=\"").Append(cs).Append('"');
                    if (rs > 1) sb.Append(" rowspan=\"").Append(rs).Append('"');
                }
                if (cells.TryGetValue((c, r), out var cp))
                {
                    var (text, color, align) = Shown(cp);
                    sb.Append(" style=\"").Append(CellCss(cp, color, align)).Append("\">").Append(Esc(text).Replace("\n", "<br>"));
                }
                else sb.Append('>');
                sb.Append("</td>");
            }
            sb.Append("</tr>\n");
        }
        sb.Append("</table>\n");
        foreach (var chart in charts) Chart(sheet, book, chart, sb);
        sb.Append("</div>\n");
    }

    static JsonElement Json(IReadOnlyDictionary<string, string> props, string name) => Parse(props.GetValueOrDefault(name) ?? (name is "merges" ? "[]" : "{}"));

    static JsonElement Parse(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        return parsed.RootElement.Clone();
    }

    static string FontStack(string font) => "'" + Esc(font).Replace("'", "") + "',Calibri,Carlito,'PingFang SC','Microsoft YaHei',sans-serif";

    /// <summary>What a cell shows: its value through the number format, the format's colour, and how Excel aligns it when the cell does not say.</summary>
    static (string Text, string? Color, string Align) Shown(IReadOnlyDictionary<string, string> p)
    {
        var value = p.GetValueOrDefault("value") ?? "";
        var format = p.GetValueOrDefault("format");
        switch (p.GetValueOrDefault("type"))
        {
            case "number" when double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n):
                var (text, color) = XlsxNumberFormat.Format(n, format);
                return (text, color, "right");
            case "date" when DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d):
                var (shown, dateColor) = XlsxNumberFormat.Format(XlsxCells.ToSerial(d, false), format ?? "yyyy-mm-dd");
                return (shown, dateColor, "right");
            case "bool": return (value == "true" ? "TRUE" : "FALSE", null, "center");
            case "error": return (value, null, "center");
            default: return (value, null, "left");
        }
    }

    static string CellCss(IReadOnlyDictionary<string, string> p, string? formatColor, string align)
    {
        var css = new StringBuilder();
        if (p.TryGetValue("fill", out var fill)) css.Append("background:#").Append(fill).Append(';');
        if ((formatColor ?? p.GetValueOrDefault("color")) is { } color) css.Append("color:#").Append(color).Append(';');
        if (p.GetValueOrDefault("bold") == "true") css.Append("font-weight:bold;");
        if (p.GetValueOrDefault("italic") == "true") css.Append("font-style:italic;");
        var lines = (p.GetValueOrDefault("underline") == "true" ? "underline " : "") + (p.GetValueOrDefault("strike") == "true" ? "line-through" : "");
        if (lines.Length > 0) css.Append("text-decoration:").Append(lines.Trim()).Append(';');
        if (p.TryGetValue("size", out var size)) css.Append("font-size:").Append(size).Append("pt;");
        if (p.TryGetValue("font", out var font)) css.Append("font-family:").Append(FontStack(font)).Append(';');
        align = p.GetValueOrDefault("align") ?? align;
        if (align != "left") css.Append("text-align:").Append(align).Append(';');
        var valign = p.GetValueOrDefault("valign");
        if (valign is "top" or "middle") css.Append("vertical-align:").Append(valign).Append(';');
        if (p.GetValueOrDefault("wrap") == "true") css.Append("white-space:pre-wrap;word-wrap:break-word;");
        if (int.TryParse(p.GetValueOrDefault("indent"), out var indent)) css.Append(align == "right" ? "padding-right:" : "padding-left:").Append(3 + indent * 9).Append("px;");
        var sides = new Dictionary<string, string>();
        if (p.TryGetValue("borders", out var borders))
            foreach (var side in Parse(borders).EnumerateObject()) sides[side.Name] = side.Value.GetString() ?? "none";
        else if (p.TryGetValue("border", out var border)) foreach (var side in new[] { "top", "right", "bottom", "left" }) sides[side] = border;
        var borderColor = "#" + (p.GetValueOrDefault("borderColor") ?? "000000");
        foreach (var (side, style) in sides)
            if (BorderCss.TryGetValue(style, out var line)) css.Append("border-").Append(side).Append(':').Append(line).Append(' ').Append(borderColor).Append(';');
        return css.ToString();
    }

    /// <summary>A chart as Excel draws its kind (column, bar, line, area, pie, doughnut, scatter; stacked or not) from the cells it names,
    /// in the theme's accent colours, placed where it sits on the sheet.</summary>
    static void Chart(Node sheet, Book book, IReadOnlyDictionary<string, string> p, StringBuilder sb)
    {
        var type = p.GetValueOrDefault("type") ?? "column";
        double w = Px(p.GetValueOrDefault("w")), h = Px(p.GetValueOrDefault("h"));
        if (w <= 0 || h <= 0) return;
        var stacked = p.GetValueOrDefault("stacked") == "true";
        var categories = p.TryGetValue("categories", out var cat) ? book.Range(sheet, cat).Select(c => c is null ? "" : Shown(c).Text).ToList() : [];
        var series = new List<(string Name, List<double> Values, List<double>? X)>();
        string? axisFormat = null; // the value axis speaks the first value cell's format, as Excel links it to the source
        using (var json = JsonDocument.Parse(p.GetValueOrDefault("series") ?? "[]"))
            foreach (var s in json.RootElement.EnumerateArray())
            {
                static double Number(IReadOnlyDictionary<string, string>? c) => double.TryParse(c?.GetValueOrDefault("value"), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0;
                var cells = book.Range(sheet, s.TryGetProperty("values", out var v) ? v.GetString() ?? "" : "");
                axisFormat ??= cells.FirstOrDefault(c => c?.ContainsKey("format") == true)?["format"];
                var values = cells.Select(Number).ToList();
                var x = s.TryGetProperty("x", out var xs) ? book.Range(sheet, xs.GetString() ?? "").Select(Number).ToList() : null;
                var name = s.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "";
                if (name.Length > 1 && name[0] == '"' && name[^1] == '"') name = name[1..^1].Replace("\"\"", "\"");
                var named = name.Length > 0 && book.Range(sheet, name) is [{ } nameCell] && Shown(nameCell).Text is { Length: > 0 } shown ? shown : name;
                series.Add((named.Length > 0 ? named : "Series " + (series.Count + 1), values, x));
            }
        var svg = new StringBuilder();
        string Color(int i) => "#" + book.Accent(i);
        string Num(double v) => v.ToString("0.#", CultureInfo.InvariantCulture);
        void Text(double x, double y, string text, string anchor = "middle", int size = 10, string fill = "#595959") =>
            svg.Append("<text x=\"").Append(Num(x)).Append("\" y=\"").Append(Num(y)).Append("\" font-size=\"").Append(size).Append("\" fill=\"").Append(fill)
                .Append("\" text-anchor=\"").Append(anchor).Append("\">").Append(Esc(text)).Append("</text>");
        double top = 8, left = 8, right = w - 8, bottom = h - 8;
        if (p.GetValueOrDefault("title") is { Length: > 0 } title) { Text(w / 2, 22, title, size: 14, fill: "#404040"); top = 34; }
        var pie = type is "pie" or "doughnut";
        var legend = pie ? categories.Select((c, i) => (c, Color(i))).ToList() : series.Select((s, i) => (s.Name, Color(i))).ToList();
        if (p.GetValueOrDefault("legend") != "none" && legend.Count > 0)
        {
            var x = left + 4.0;
            bottom -= 18;
            foreach (var (name, fill) in legend.Take(12))
            {
                svg.Append("<rect x=\"").Append(Num(x)).Append("\" y=\"").Append(Num(h - 19)).Append("\" width=\"8\" height=\"8\" fill=\"").Append(fill).Append("\"/>");
                Text(x + 11, h - 11, name, "start");
                x += 20 + Math.Min(120, name.Length * 6);
            }
        }
        if (pie)
        {
            var values = series.FirstOrDefault().Values ?? [];
            var total = values.Where(v => v > 0).Sum();
            double cx = (left + right) / 2, cy = (top + bottom) / 2, radius = Math.Max(10, Math.Min(right - left, bottom - top) / 2 - 4), a0 = -Math.PI / 2;
            for (var i = 0; i < values.Count && total > 0; i++)
            {
                var a1 = a0 + Math.Max(0, values[i]) / total * 2 * Math.PI;
                string At(double a) => Num(cx + radius * Math.Cos(a)) + " " + Num(cy + radius * Math.Sin(a));
                var d = a1 - a0 >= 2 * Math.PI - 1e-9
                    ? $"M {Num(cx - radius)} {Num(cy)} a {Num(radius)} {Num(radius)} 0 1 0 {Num(2 * radius)} 0 a {Num(radius)} {Num(radius)} 0 1 0 {Num(-2 * radius)} 0"
                    : $"M {Num(cx)} {Num(cy)} L {At(a0)} A {Num(radius)} {Num(radius)} 0 {(a1 - a0 > Math.PI ? 1 : 0)} 1 {At(a1)} Z";
                svg.Append("<path d=\"").Append(d).Append("\" fill=\"").Append(Color(i)).Append("\" stroke=\"#fff\"/>");
                a0 = a1;
            }
            if (type == "doughnut") svg.Append("<circle cx=\"").Append(Num(cx)).Append("\" cy=\"").Append(Num(cy)).Append("\" r=\"").Append(Num(radius / 2)).Append("\" fill=\"#fff\"/>");
        }
        else if (series.Count > 0)
        {
            var count = Math.Max(categories.Count, series.Max(s => s.Values.Count));
            var horizontal = type == "bar";
            var sums = Enumerable.Range(0, count).Select(i => (Up: series.Sum(s => Math.Max(0, s.Values.ElementAtOrDefault(i))), Down: series.Sum(s => Math.Min(0, s.Values.ElementAtOrDefault(i))))).ToList();
            var all = stacked && type != "scatter" ? sums.SelectMany(s => new[] { s.Up, s.Down }) : series.SelectMany(s => s.Values);
            double lo = Math.Min(0, all.DefaultIfEmpty(0).Min()), hi = Math.Max(0, all.DefaultIfEmpty(1).Max());
            if (hi == lo) hi = lo + 1;
            var step = Math.Pow(10, Math.Floor(Math.Log10((hi - lo) / 5)));
            step *= (hi - lo) / step > 25 ? 5 : (hi - lo) / step > 10 ? 2 : 1;
            (lo, hi) = (Math.Floor(lo / step) * step, Math.Ceiling(hi / step) * step);
            left += 44;
            bottom -= 18;
            double V(double v) => horizontal ? left + (v - lo) / (hi - lo) * (right - left) : bottom - (v - lo) / (hi - lo) * (bottom - top);
            for (var v = lo; v <= hi + step / 2; v += step)
            {
                var at = V(v);
                svg.Append(horizontal ? $"<line x1=\"{Num(at)}\" y1=\"{Num(top)}\" x2=\"{Num(at)}\" y2=\"{Num(bottom)}\"" : $"<line x1=\"{Num(left)}\" y1=\"{Num(at)}\" x2=\"{Num(right)}\" y2=\"{Num(at)}\"")
                    .Append(" stroke=\"#d9d9d9\"/>");
                var label = XlsxNumberFormat.Format(Math.Round(v, 10), axisFormat).Text;
                if (horizontal) Text(at, bottom + 13, label);
                else Text(left - 4, at + 3, label, "end");
            }
            var band = (horizontal ? bottom - top : right - left) / Math.Max(1, count);
            double Band(int i) => horizontal ? top + band * (i + 0.5) : left + band * (i + 0.5);
            if (type == "scatter")
            {
                var xs = series.SelectMany(s => s.X ?? Enumerable.Range(1, s.Values.Count).Select(i => (double)i).ToList()).DefaultIfEmpty(0).ToList();
                double x0 = xs.Min(), x1 = xs.Max() == x0 ? x0 + 1 : xs.Max();
                for (var si = 0; si < series.Count; si++)
                    for (var i = 0; i < series[si].Values.Count; i++)
                    {
                        var xv = series[si].X?.ElementAtOrDefault(i) ?? i + 1;
                        svg.Append("<circle cx=\"").Append(Num(left + (xv - x0) / (x1 - x0) * (right - left))).Append("\" cy=\"").Append(Num(V(series[si].Values[i])))
                            .Append("\" r=\"3.5\" fill=\"").Append(Color(si)).Append("\"/>");
                    }
            }
            else
            {
                for (var i = 0; i < count && i < categories.Count; i++)
                    if (horizontal) Text(left - 4, Band(i) + 3, categories[i], "end");
                    else Text(Band(i), bottom + 13, categories[i]);
                var baseUp = new double[count];
                var baseDown = new double[count];
                for (var si = 0; si < series.Count; si++)
                {
                    var values = series[si].Values;
                    if (type is "line" or "area")
                    {
                        var points = new List<(double X, double Y, double Base)>();
                        for (var i = 0; i < values.Count && i < count; i++)
                        {
                            var from = stacked ? baseUp[i] : 0;
                            var to = from + values[i];
                            if (stacked) baseUp[i] = to;
                            points.Add((Band(i), V(to), V(from)));
                        }
                        if (points.Count == 0) continue;
                        if (type == "area")
                            svg.Append("<path d=\"M ").Append(string.Join(" L ", points.Select(q => Num(q.X) + " " + Num(q.Y)).Concat(points.AsEnumerable().Reverse().Select(q => Num(q.X) + " " + Num(q.Base)))))
                                .Append(" Z\" fill=\"").Append(Color(si)).Append("\" fill-opacity=\"0.85\"/>");
                        else svg.Append("<polyline points=\"").Append(string.Join(" ", points.Select(q => Num(q.X) + "," + Num(q.Y)))).Append("\" fill=\"none\" stroke=\"").Append(Color(si)).Append("\" stroke-width=\"2.25\"/>");
                        continue;
                    }
                    var thick = band * 0.62 / (stacked ? 1 : series.Count);
                    for (var i = 0; i < values.Count && i < count; i++)
                    {
                        var v = values[i];
                        double from = stacked ? (v >= 0 ? baseUp[i] : baseDown[i]) : 0, to = from + v;
                        if (stacked) { if (v >= 0) baseUp[i] = to; else baseDown[i] = to; }
                        var offset = Band(i) - band * 0.31 + (stacked ? 0 : si * thick);
                        double a = V(Math.Min(from, to)), b = V(Math.Max(from, to));
                        svg.Append(horizontal
                            ? $"<rect x=\"{Num(a)}\" y=\"{Num(offset)}\" width=\"{Num(Math.Max(0.5, b - a))}\" height=\"{Num(Math.Max(1, thick - 1))}\""
                            : $"<rect x=\"{Num(offset)}\" y=\"{Num(b)}\" width=\"{Num(Math.Max(1, thick - 1))}\" height=\"{Num(Math.Max(0.5, a - b))}\"").Append(" fill=\"").Append(Color(si)).Append("\"/>");
                    }
                }
            }
        }
        sb.Append("<svg class=\"chart\" style=\"left:").Append(PxText(Px(p.GetValueOrDefault("x")))).Append(";top:").Append(PxText(Px(p.GetValueOrDefault("y"))))
            .Append("\" width=\"").Append(Num(w)).Append("\" height=\"").Append(Num(h)).Append("\" viewBox=\"0 0 ").Append(Num(w)).Append(' ').Append(Num(h))
            .Append("\" font-family=\"Calibri,Carlito,sans-serif\"><rect width=\"100%\" height=\"100%\" fill=\"#fff\" stroke=\"#d9d9d9\"/>").Append(svg).Append("</svg>\n");
    }
}
