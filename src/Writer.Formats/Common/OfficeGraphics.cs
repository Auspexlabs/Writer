using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Writer.Formats.Common;

/// <summary>
/// What an editor needs to draw a chart or a SmartArt graphic as Office drew it, read from the parts' XML as it stands and never
/// written. A chart gives its kind, title, axis titles, categories and the values its series cache (what Office last drew, so a
/// chart linked to a workbook that is not there still shows); SmartArt gives the shapes Office laid out for it (the drawing part it
/// keeps beside the data), each with its place, geometry, colours and text, or only the texts when there is no drawing. Colours
/// are resolved against the document's theme, with the modifiers Office applies (lumMod, lumOff, tint, shade).
/// </summary>
public static class OfficeGraphics
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    static readonly XNamespace Dsp = "http://schemas.microsoft.com/office/drawing/2008/diagram";
    static readonly XNamespace Dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    static readonly XNamespace Cx = "http://schemas.microsoft.com/office/drawing/2014/chartex";

    public const string ChartUri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    public const string ChartExUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";
    public const string DiagramUri = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

    /// <summary>The Office theme's colours, for a document without a theme part.</summary>
    static readonly Dictionary<string, string> OfficeScheme = new()
    {
        ["dk1"] = "000000", ["lt1"] = "FFFFFF", ["dk2"] = "44546A", ["lt2"] = "E7E6E6",
        ["accent1"] = "4472C4", ["accent2"] = "ED7D31", ["accent3"] = "A5A5A5", ["accent4"] = "FFC000", ["accent5"] = "5B9BD5", ["accent6"] = "70AD47",
        ["hlink"] = "0563C1", ["folHlink"] = "954F72",
    };

    /// <summary>A theme's colour scheme by name (dk1, lt1, accent1…), the Office theme's where the part is missing or silent.</summary>
    public static IReadOnlyDictionary<string, string> Scheme(ThemePart? theme)
    {
        var scheme = new Dictionary<string, string>(OfficeScheme);
        if (theme is null) return scheme;
        try
        {
            using var stream = theme.GetStream(FileMode.Open, FileAccess.Read);
            var clr = XDocument.Load(stream).Descendants(A + "clrScheme").FirstOrDefault();
            foreach (var e in clr?.Elements() ?? [])
            {
                var c = e.Element(A + "srgbClr")?.Attribute("val")?.Value ?? e.Element(A + "sysClr")?.Attribute("lastClr")?.Value;
                if (c is { Length: 6 }) scheme[e.Name.LocalName] = c.ToUpperInvariant();
            }
        }
        catch (Exception e) when (e is System.Xml.XmlException or IOException) { }
        return scheme;
    }

    static XDocument? Load(OpenXmlPart part)
    {
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XDocument.Load(stream);
        }
        catch (Exception e) when (e is System.Xml.XmlException or IOException) { return null; }
    }

    // ----- colours -----

    /// <summary>The colour an element holding one (a:solidFill, a:fontRef, a:gs…) stands for, as RRGGBB; null when it holds none.</summary>
    public static string? ColorIn(XElement? holder, IReadOnlyDictionary<string, string> scheme, string? phClr = null)
    {
        if (holder is null) return null;
        foreach (var e in holder.Elements())
        {
            string? baseColor = e.Name.LocalName switch
            {
                "srgbClr" => e.Attribute("val")?.Value,
                "sysClr" => e.Attribute("lastClr")?.Value ?? (e.Attribute("val")?.Value == "window" ? "FFFFFF" : "000000"),
                "schemeClr" => e.Attribute("val")?.Value switch
                {
                    "phClr" => phClr,
                    "bg1" => scheme.GetValueOrDefault("lt1"), "tx1" => scheme.GetValueOrDefault("dk1"),
                    "bg2" => scheme.GetValueOrDefault("lt2"), "tx2" => scheme.GetValueOrDefault("dk2"),
                    { } name => scheme.GetValueOrDefault(name),
                    null => null,
                },
                "prstClr" => e.Attribute("val")?.Value switch { "black" => "000000", "white" => "FFFFFF", "red" => "FF0000", "blue" => "0000FF", "green" => "008000", "yellow" => "FFFF00", "gray" => "808080", _ => null },
                "scrgbClr" => Hex(Pct(e, "r") / 100, Pct(e, "g") / 100, Pct(e, "b") / 100),
                "hslClr" => null,
                _ => null,
            };
            if (baseColor is not { Length: 6 }) continue;
            return Modify(baseColor.ToUpperInvariant(), e);
        }
        return null;
    }

    static double Pct(XElement e, string attr) => double.TryParse(e.Attribute(attr)?.Value, NumberStyles.Float, Inv, out var v) ? v / 1000.0 : 0;

    static string Hex(double r, double g, double b) =>
        string.Concat(new[] { r, g, b }.Select(x => ((int)Math.Round(Math.Clamp(x, 0, 1) * 255)).ToString("X2", Inv)));

    /// <summary>Office's colour modifiers: lumMod and lumOff on the HSL lightness, tint towards white, shade towards black.</summary>
    static string Modify(string hex, XElement color)
    {
        double r = Convert.ToInt32(hex[..2], 16) / 255.0, g = Convert.ToInt32(hex[2..4], 16) / 255.0, b = Convert.ToInt32(hex[4..], 16) / 255.0;
        foreach (var m in color.Elements())
        {
            var v = double.TryParse(m.Attribute("val")?.Value, NumberStyles.Float, Inv, out var x) ? x / 100000.0 : 1;
            switch (m.Name.LocalName)
            {
                case "lumMod" or "lumOff":
                    var (h, s, l) = ToHsl(r, g, b);
                    l = m.Name.LocalName == "lumMod" ? l * v : l + v;
                    (r, g, b) = FromHsl(h, s, Math.Clamp(l, 0, 1));
                    break;
                case "tint": (r, g, b) = (1 - (1 - r) * v, 1 - (1 - g) * v, 1 - (1 - b) * v); break;
                case "shade": (r, g, b) = (r * v, g * v, b * v); break;
            }
        }
        return Hex(r, g, b);
    }

    static (double H, double S, double L) ToHsl(double r, double g, double b)
    {
        double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b)), l = (max + min) / 2;
        if (max == min) return (0, 0, l);
        var d = max - min;
        var s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        var h = max == r ? (g - b) / d + (g < b ? 6 : 0) : max == g ? (b - r) / d + 2 : (r - g) / d + 4;
        return (h / 6, s, l);
    }

    static (double R, double G, double B) FromHsl(double h, double s, double l)
    {
        if (s == 0) return (l, l, l);
        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;
        static double Hue(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            return t < 1 / 6.0 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2 / 3.0 ? p + (q - p) * (2 / 3.0 - t) * 6 : p;
        }
        return (Hue(p, q, h + 1 / 3.0), Hue(p, q, h), Hue(p, q, h - 1 / 3.0));
    }

    static string TextOf(XElement? e) => e is null ? "" : string.Join("\n", e.Descendants(A + "p").Select(p => string.Concat(p.Descendants(A + "t").Select(t => t.Value)))).Trim();

    // ----- charts -----

    static readonly string[] PlotKinds =
        ["barChart", "bar3DChart", "lineChart", "line3DChart", "pieChart", "pie3DChart", "doughnutChart", "ofPieChart", "areaChart", "area3DChart",
         "scatterChart", "radarChart", "bubbleChart", "stockChart", "surfaceChart", "surface3DChart"];

    static string KindOf(string plot) => plot switch
    {
        "barChart" or "bar3DChart" => "bar",
        "lineChart" or "line3DChart" or "stockChart" => "line",
        "pieChart" or "pie3DChart" or "ofPieChart" => "pie",
        "doughnutChart" => "doughnut",
        "areaChart" or "area3DChart" => "area",
        "scatterChart" or "bubbleChart" => "scatter",
        "radarChart" => "radar",
        _ => "other",
    };

    /// <summary>A chart part (c:chartSpace) as JSON: { kind, dir, grouping, title, xTitle, yTitle, min, max, legend, hole, cats,
    /// series: [{ name, kind, values, x, color, points }] }; null when the part cannot be read.</summary>
    public static string? ChartJson(OpenXmlPart part, IReadOnlyDictionary<string, string> scheme)
    {
        var chart = Load(part)?.Root?.Element(C + "chart");
        var plot = chart?.Element(C + "plotArea");
        if (plot is null) return null;
        var groups = plot.Elements().Where(e => e.Name.Namespace == C && PlotKinds.Contains(e.Name.LocalName)).ToList();
        if (groups.Count == 0) return null;
        using var ms = new MemoryStream();
        using (var w = new Utf8JsonWriter(ms))
        {
            w.WriteStartObject();
            var first = groups[0];
            w.WriteString("kind", KindOf(first.Name.LocalName));
            if (first.Element(C + "barDir")?.Attribute("val")?.Value is { } dir) w.WriteString("dir", dir);
            if (first.Element(C + "grouping")?.Attribute("val")?.Value is { } grouping) w.WriteString("grouping", grouping);
            if (first.Element(C + "holeSize")?.Attribute("val")?.Value is { } hole && int.TryParse(hole.TrimEnd('%'), NumberStyles.Integer, Inv, out var h)) w.WriteNumber("hole", h);
            else if (first.Name.LocalName == "doughnutChart") w.WriteNumber("hole", 50);
            if (first.Element(C + "radarStyle")?.Attribute("val")?.Value is { } radar) w.WriteString("radarStyle", radar);
            if (first.Element(C + "scatterStyle")?.Attribute("val")?.Value is { } scatter) w.WriteString("scatterStyle", scatter);
            var deleted = chart!.Element(C + "autoTitleDeleted")?.Attribute("val")?.Value is "1" or "true";
            var title = TextOf(chart.Element(C + "title")?.Element(C + "tx"));
            if (title.Length == 0 && chart.Element(C + "title") is not null && !deleted && groups.SelectMany(g => g.Elements(C + "ser")).Count() == 1)
                title = SeriesName(groups[0].Element(C + "ser")!); // a one-series chart's own title is the series name
            if (title.Length > 0) w.WriteString("title", title);
            // axes: the category (or first value) axis across, the value axis up; a bar chart turns them
            var axes = plot.Elements().Where(e => e.Name.LocalName is "catAx" or "valAx" or "dateAx" or "serAx").ToList();
            var catAx = axes.FirstOrDefault(a => a.Name.LocalName is "catAx" or "dateAx") ?? axes.FirstOrDefault(a => a.Name.LocalName == "valAx" && a.Element(C + "axPos")?.Attribute("val")?.Value is "b" or "t");
            var valAx = axes.FirstOrDefault(a => a.Name.LocalName == "valAx" && a != catAx);
            if (TextOf(catAx?.Element(C + "title")?.Element(C + "tx")) is { Length: > 0 } xt) w.WriteString("xTitle", xt);
            if (TextOf(valAx?.Element(C + "title")?.Element(C + "tx")) is { Length: > 0 } yt) w.WriteString("yTitle", yt);
            var scaling = valAx?.Element(C + "scaling");
            if (Num(scaling?.Element(C + "min")) is { } min) w.WriteNumber("min", min);
            if (Num(scaling?.Element(C + "max")) is { } max) w.WriteNumber("max", max);
            if (valAx?.Element(C + "numFmt")?.Attribute("formatCode")?.Value is { } fmt && fmt != "General") w.WriteString("format", fmt);
            if (valAx?.Element(C + "majorGridlines") is not null) w.WriteBoolean("grid", true);
            var legend = chart.Element(C + "legend");
            w.WriteString("legend", legend is null ? "" : legend.Element(C + "legendPos")?.Attribute("val")?.Value ?? "r");
            if (plot.Descendants(C + "dLbls").Any(d => d.Element(C + "showVal")?.Attribute("val")?.Value is "1" or "true")) w.WriteBoolean("labels", true);
            string[]? cats = null;
            // a group drawn against another value axis than the first group's: on the right, to its own scale (a combo chart's line)
            var firstVal = first.Elements(C + "axId").Select(a => a.Attribute("val")?.Value).LastOrDefault();
            var secondary = new HashSet<string>(groups.Skip(1).Select(g => g.Elements(C + "axId").Select(a => a.Attribute("val")?.Value).LastOrDefault() ?? "")
                .Where(id => id.Length > 0 && id != firstVal && axes.Any(a => a.Name.LocalName == "valAx" && a.Element(C + "axId")?.Attribute("val")?.Value == id)));
            w.WritePropertyName("series");
            w.WriteStartArray();
            var index = 0;
            foreach (var g in groups)
                foreach (var ser in g.Elements(C + "ser"))
                {
                    w.WriteStartObject();
                    w.WriteString("name", SeriesName(ser));
                    if (g != first || groups.Count > 1) w.WriteString("kind", KindOf(g.Name.LocalName));
                    if (g.Element(C + "axId") is not null && secondary.Contains(g.Elements(C + "axId").Select(a => a.Attribute("val")?.Value).LastOrDefault() ?? "")) w.WriteBoolean("y2", true);
                    var fill = ser.Element(C + "spPr")?.Element(A + "noFill") is not null && ser.Element(C + "spPr")?.Element(A + "ln")?.Element(A + "solidFill") is null ? "none"
                        : ColorIn(ser.Element(C + "spPr")?.Element(A + "solidFill"), scheme)
                        ?? ColorIn(ser.Element(C + "spPr")?.Element(A + "ln")?.Element(A + "solidFill"), scheme)
                        ?? ColorIn(ser.Element(C + "marker")?.Element(C + "spPr")?.Element(A + "solidFill"), scheme)
                        ?? scheme.GetValueOrDefault("accent" + (index % 6 + 1));
                    if (fill is not null) w.WriteString("color", fill);
                    var values = Values(ser.Element(C + "val") ?? ser.Element(C + "yVal"));
                    WriteNumbers(w, "values", values);
                    if (ser.Element(C + "xVal") is { } xVal) WriteNumbers(w, "x", Values(xVal));
                    if (ser.Element(C + "bubbleSize") is { } size) WriteNumbers(w, "size", Values(size));
                    cats ??= Labels(ser.Element(C + "cat"));
                    var points = ser.Elements(C + "dPt").Select(d => (Idx: (int?)Num(d.Element(C + "idx")) ?? -1, Color: ColorIn(d.Element(C + "spPr")?.Element(A + "solidFill"), scheme))).Where(p => p.Idx >= 0 && p.Color is not null).ToList();
                    if (points.Count > 0)
                    {
                        w.WritePropertyName("points");
                        w.WriteStartObject();
                        foreach (var (idx, color) in points) w.WriteString(idx.ToString(Inv), color);
                        w.WriteEndObject();
                    }
                    w.WriteEndObject();
                    index++;
                }
            w.WriteEndArray();
            w.WritePropertyName("cats");
            w.WriteStartArray();
            foreach (var c in cats ?? []) w.WriteStringValue(c);
            w.WriteEndArray();
            w.WriteEndObject();
        }
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    static string SeriesName(XElement ser)
    {
        var tx = ser.Element(C + "tx");
        if (tx is null) return "";
        return tx.Element(C + "v")?.Value ?? tx.Descendants(C + "pt").FirstOrDefault()?.Element(C + "v")?.Value ?? TextOf(tx.Element(C + "rich"));
    }

    static double? Num(XElement? e) => double.TryParse(e?.Attribute("val")?.Value ?? e?.Value, NumberStyles.Float, Inv, out var v) ? v : null;

    /// <summary>A data source's cached points by index (numRef/numCache, numLit, strRef/strCache, strLit, multi-level labels' first level).</summary>
    static List<(int Idx, string V)> Points(XElement? source)
    {
        if (source is null) return [];
        var cache = source.Descendants().FirstOrDefault(e => e.Name.Namespace == C && e.Name.LocalName is "numCache" or "numLit" or "strCache" or "strLit")
            ?? source.Descendants(C + "multiLvlStrCache").FirstOrDefault()?.Element(C + "lvl");
        if (cache is null) return [];
        return cache.Elements(C + "pt").Select(p => (int.TryParse(p.Attribute("idx")?.Value, NumberStyles.Integer, Inv, out var i) ? i : -1, p.Element(C + "v")?.Value ?? ""))
            .Where(p => p.Item1 >= 0).ToList();
    }

    static int CountOf(XElement? source) =>
        (int?)Num(source?.Descendants(C + "ptCount").FirstOrDefault()) ?? 0;

    static double?[] Values(XElement? source)
    {
        var points = Points(source);
        var n = Math.Max(CountOf(source), points.Count == 0 ? 0 : points.Max(p => p.Idx) + 1);
        var values = new double?[Math.Min(n, 4096)];
        foreach (var (i, v) in points)
            if (i < values.Length && double.TryParse(v, NumberStyles.Float, Inv, out var d)) values[i] = d;
        return values;
    }

    static string[] Labels(XElement? source)
    {
        var points = Points(source);
        var n = Math.Max(CountOf(source), points.Count == 0 ? 0 : points.Max(p => p.Idx) + 1);
        var labels = new string[Math.Min(n, 4096)];
        Array.Fill(labels, "");
        foreach (var (i, v) in points) if (i < labels.Length) labels[i] = v;
        return labels;
    }

    static void WriteNumbers(Utf8JsonWriter w, string name, double?[] values)
    {
        w.WritePropertyName(name);
        w.WriteStartArray();
        foreach (var v in values)
            if (v is { } d && double.IsFinite(d)) w.WriteNumberValue(d); else w.WriteNullValue();
        w.WriteEndArray();
    }

    /// <summary>An Office 2016 chart (cx:chartSpace: waterfall, funnel, treemap, sunburst, histogram, box and whisker, map) as the same
    /// JSON: kind is its layout (waterfall, funnel, treemap, sunburst, histogram, boxWhisker, pareto, regionMap), one series per cx:series
    /// with the values of its data, the categories its data's first string dimension.</summary>
    public static string? ChartExJson(OpenXmlPart part, IReadOnlyDictionary<string, string> scheme)
    {
        var root = Load(part)?.Root;
        if (root is null) return null;
        var data = root.Element(Cx + "chartData")?.Elements(Cx + "data").ToDictionary(d => d.Attribute("id")?.Value ?? "", d => d) ?? [];
        var series = root.Descendants(Cx + "series").ToList();
        if (series.Count == 0) return null;
        using var ms = new MemoryStream();
        using (var w = new Utf8JsonWriter(ms))
        {
            w.WriteStartObject();
            var layout = series[0].Attribute("layoutId")?.Value ?? "other";
            w.WriteString("kind", layout switch { "clusteredColumn" => "histogram", "paretoLine" => "pareto", _ => layout });
            if (TextOf(root.Element(Cx + "chart")?.Element(Cx + "title")) is { Length: > 0 } title) w.WriteString("title", title);
            w.WriteString("legend", root.Element(Cx + "chart")?.Element(Cx + "legend") is null ? "" : "b");
            string[]? cats = null;
            w.WritePropertyName("series");
            w.WriteStartArray();
            var index = 0;
            foreach (var s in series.Where(s => s.Attribute("layoutId")?.Value != "paretoLine"))
            {
                w.WriteStartObject();
                w.WriteString("name", s.Element(Cx + "tx")?.Descendants(Cx + "v").FirstOrDefault()?.Value ?? "");
                var color = ColorIn(s.Element(Cx + "spPr")?.Element(A + "solidFill"), scheme) ?? scheme.GetValueOrDefault("accent" + (index % 6 + 1));
                if (color is not null) w.WriteString("color", color);
                var d = data.GetValueOrDefault(s.Element(Cx + "dataId")?.Attribute("val")?.Value ?? "");
                var nums = d?.Elements(Cx + "numDim").FirstOrDefault()?.Element(Cx + "lvl")?.Elements(Cx + "pt")
                    .Select(p => (Idx: int.TryParse(p.Attribute("idx")?.Value, NumberStyles.Integer, Inv, out var i) ? i : -1, V: double.TryParse(p.Value, NumberStyles.Float, Inv, out var v) ? v : (double?)null))
                    .Where(p => p.Idx >= 0).ToList() ?? [];
                var values = new double?[nums.Count == 0 ? 0 : Math.Min(4096, nums.Max(p => p.Idx) + 1)];
                foreach (var (i, v) in nums) if (i < values.Length) values[i] = v;
                WriteNumbers(w, "values", values);
                cats ??= d?.Elements(Cx + "strDim").FirstOrDefault()?.Elements(Cx + "lvl").FirstOrDefault()?.Elements(Cx + "pt").Select(p => p.Value).ToArray();
                // a waterfall's subtotals: the points drawn from zero, not floating
                var totals = s.Element(Cx + "layoutPr")?.Element(Cx + "subtotals")?.Elements(Cx + "idx").Select(i => i.Attribute("val")?.Value).OfType<string>().ToList();
                if (totals is { Count: > 0 })
                {
                    w.WritePropertyName("totals");
                    w.WriteStartArray();
                    foreach (var t in totals) if (int.TryParse(t, NumberStyles.Integer, Inv, out var ti)) w.WriteNumberValue(ti);
                    w.WriteEndArray();
                }
                w.WriteEndObject();
                index++;
            }
            w.WriteEndArray();
            w.WritePropertyName("cats");
            w.WriteStartArray();
            foreach (var c in cats ?? []) w.WriteStringValue(c);
            w.WriteEndArray();
            w.WriteEndObject();
        }
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    /// <summary>The chart title of a chart part, for a label.</summary>
    public static string? ChartTitle(string? json)
    {
        if (json is null) return null;
        using var d = JsonDocument.Parse(json);
        return d.RootElement.TryGetProperty("title", out var t) ? t.GetString() : null;
    }

    // ----- SmartArt -----

    /// <summary>The shapes Office drew for a SmartArt graphic (its dsp:drawing part) as JSON: { shapes: [{ x, y, w, h, rot, geom, fill,
    /// line, lw, text, size, color, bold, anchor }] } in EMU from the graphic's corner; else, from the data part, { texts: [...] } — the
    /// words of its nodes in order. Null when neither part can be read.</summary>
    public static string? SmartArtJson(OpenXmlPart? drawing, OpenXmlPart? data, IReadOnlyDictionary<string, string> scheme)
    {
        var tree = drawing is null ? null : Load(drawing)?.Root?.Element(Dsp + "spTree");
        using var ms = new MemoryStream();
        using (var w = new Utf8JsonWriter(ms))
        {
            w.WriteStartObject();
            if (tree is not null)
            {
                var ext=tree.Element(Dsp+"grpSpPr")?.Element(A+"xfrm")?.Element(A+"chExt");
                if(long.TryParse(ext?.Attribute("cx")?.Value,out var width))w.WriteNumber("width",width);
                if(long.TryParse(ext?.Attribute("cy")?.Value,out var height))w.WriteNumber("height",height);
                w.WritePropertyName("shapes");
                w.WriteStartArray();
                foreach (var sp in tree.Descendants(Dsp + "sp")) WriteShape(w, sp, scheme);
                w.WriteEndArray();
            }
            else
            {
                var model = data is null ? null : Load(data)?.Root;
                if (model is null) return null;
                w.WritePropertyName("texts");
                w.WriteStartArray();
                foreach (var pt in model.Descendants(Dgm + "pt").Where(p => (p.Attribute("type")?.Value ?? "node") == "node"))
                    if (TextOf(pt.Element(Dgm + "t")) is { Length: > 0 } t) w.WriteStringValue(t);
                w.WriteEndArray();
            }
            if(data is not null){var model=OfficeSmartArt.Read(data);w.WriteString("layout",model.Layout);w.WritePropertyName("nodes");w.WriteStartArray();foreach(var n in model.Nodes){w.WriteStartObject();w.WriteString("id",n.Id);w.WriteString("text",n.Text);if(n.Parent is not null)w.WriteString("parent",n.Parent);w.WriteEndObject();}w.WriteEndArray();}
            w.WriteEndObject();
        }
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    static void WriteShape(Utf8JsonWriter w, XElement sp, IReadOnlyDictionary<string, string> scheme)
    {
        var spPr = sp.Element(Dsp + "spPr");
        var xfrm = spPr?.Element(A + "xfrm");
        var off = xfrm?.Element(A + "off");
        var ext = xfrm?.Element(A + "ext");
        long L(XElement? e, string a) => long.TryParse(e?.Attribute(a)?.Value, NumberStyles.Integer, Inv, out var v) ? v : 0;
        var style = sp.Element(Dsp + "style");
        w.WriteStartObject();
        w.WriteNumber("x", L(off, "x"));
        w.WriteNumber("y", L(off, "y"));
        w.WriteNumber("w", L(ext, "cx"));
        w.WriteNumber("h", L(ext, "cy"));
        if (L(xfrm, "rot") is var rot and not 0) w.WriteNumber("rot", rot / 60000.0);
        if (xfrm?.Attribute("flipH")?.Value is "1" or "true") w.WriteBoolean("flipH", true);
        if (xfrm?.Attribute("flipV")?.Value is "1" or "true") w.WriteBoolean("flipV", true);
        w.WriteString("geom", spPr?.Element(A + "prstGeom")?.Attribute("prst")?.Value ?? (spPr?.Element(A + "custGeom") is null ? "rect" : "custom"));
        var styleFill = ColorIn(style?.Element(A + "fillRef"), scheme);
        var fill = spPr?.Element(A + "noFill") is not null ? "none"
            : ColorIn(spPr?.Element(A + "solidFill"), scheme, styleFill)
            ?? ColorIn(spPr?.Element(A + "gradFill")?.Descendants(A + "gs").FirstOrDefault(), scheme, styleFill)
            ?? (spPr?.Element(A + "gradFill") is null && spPr?.Element(A + "blipFill") is null ? styleFill : null);
        if (fill is not null) w.WriteString("fill", fill);
        var ln = spPr?.Element(A + "ln");
        var styleLine = ColorIn(style?.Element(A + "lnRef"), scheme);
        var line = ln?.Element(A + "noFill") is not null ? "none" : ColorIn(ln?.Element(A + "solidFill"), scheme, styleLine) ?? styleLine;
        if (line is not null) w.WriteString("line", line);
        if (L(ln, "w") is var lw and > 0) w.WriteNumber("lw", lw);
        if(ln?.Element(A+"tailEnd")?.Attribute("type")?.Value is {} endArrow && endArrow!="none")w.WriteString("endArrow",endArrow);
        var body = sp.Element(Dsp + "txBody");
        var text = TextOf(body);
        if (text.Length > 0)
        {
            w.WriteString("text", text);
            var run = body!.Descendants(A + "rPr").FirstOrDefault() ?? body.Descendants(A + "endParaRPr").FirstOrDefault();
            var size = run?.Attribute("sz")?.Value ?? body.Descendants(A + "defRPr").FirstOrDefault()?.Attribute("sz")?.Value;
            if (int.TryParse(size, NumberStyles.Integer, Inv, out var sz)) w.WriteNumber("size", sz / 100.0);
            var color = ColorIn(run?.Element(A + "solidFill"), scheme) ?? ColorIn(style?.Element(A + "fontRef"), scheme);
            if (color is not null) w.WriteString("color", color);
            if (run?.Attribute("b")?.Value is "1" or "true") w.WriteBoolean("bold", true);
            var anchor = body.Element(A + "bodyPr")?.Attribute("anchor")?.Value;
            if (anchor is not null and not "ctr") w.WriteString("anchor", anchor);
            if (body.Descendants(A + "pPr").FirstOrDefault()?.Attribute("algn")?.Value is { } algn and not "ctr") w.WriteString("align", algn);
            // the text's own box inside the shape (an arrow's text sits in its body, not over the point)
            if (sp.Element(Dsp + "txXfrm") is { } tx)
            {
                w.WritePropertyName("tx");
                w.WriteStartObject();
                w.WriteNumber("x", L(tx.Element(A + "off"), "x"));
                w.WriteNumber("y", L(tx.Element(A + "off"), "y"));
                w.WriteNumber("w", L(tx.Element(A + "ext"), "cx"));
                w.WriteNumber("h", L(tx.Element(A + "ext"), "cy"));
                w.WriteEndObject();
            }
        }
        w.WriteEndObject();
    }

    /// <summary>The relationship id of a SmartArt graphic's drawing part, from its data part (dsp:dataModelExt relId).</summary>
    public static string? SmartArtDrawingId(OpenXmlPart data) =>
        Load(data)?.Root?.Descendants(Dsp + "dataModelExt").FirstOrDefault()?.Attribute("relId")?.Value;
}
