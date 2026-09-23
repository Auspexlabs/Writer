using System.Globalization;
using System.Security;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Writer.Formats.Xlsx;

sealed record ChartSeries(string? Name, string Values, string? X);

/// <summary>What the engine models of a chart. Everything else in the chart part is left alone.</summary>
sealed record ChartSpec(string Type, string? Title, string? Categories, IReadOnlyList<ChartSeries> Series, string Legend, bool Stacked);

/// <summary>A chart on a worksheet: its anchor in the drawing part plus its chart part.</summary>
sealed class XlsxChart(XlsxDocument doc, XlsxSheet sheet, OpenXmlCompositeElement anchor, ChartPart part) : Node
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    public override string Kind => "chart";
    public override object Anchor => anchor;
    public override string? Name => XlsxChartXml.Read(Chart, sheet.SheetName).Title;

    C.Chart? Chart => part.ChartSpace?.GetFirstChild<C.Chart>();

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var spec = XlsxChartXml.Read(Chart, sheet.SheetName);
        var props = new Dictionary<string, string> { ["type"] = spec.Type };
        if (spec.Title is { Length: > 0 } title) props["title"] = title;
        if (spec.Categories is { Length: > 0 } categories) props["categories"] = categories;
        props["series"] = XlsxChartXml.SeriesJson(spec.Series);
        props["legend"] = spec.Legend;
        if (spec.Stacked) props["stacked"] = "true";
        var box = XlsxAnchors.Read(anchor, sheet.Grid);
        props["x"] = box.X.ToString(Inv);
        props["y"] = box.Y.ToString(Inv);
        props["w"] = box.W.ToString(Inv);
        props["h"] = box.H.ToString(Inv);
        if (anchor.GetFirstChild<Xdr.GraphicFrame>()?.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value is { } id)
            props["id"] = id.ToString(Inv);
        return props;
    }

    public override void SetProp(string name, string value)
    {
        var chart = Chart ?? throw new WriterException(ErrorCode.FormatError, "This chart part has no chart", "Remove it and add a new chart.");
        switch (name)
        {
            case "title": XlsxChartXml.SetTitle(chart, value); break;
            case "legend": XlsxChartXml.SetLegend(chart, value); break;
            case "type" or "series" or "categories" or "stacked":
                var spec = XlsxChartXml.Read(chart, sheet.SheetName);
                if (name != "type" && !XlsxChartXml.CanBuild(spec.Type))
                    throw new WriterException(ErrorCode.Validation, $"A {spec.Type} chart's series cannot be edited here",
                        "Only its title, legend and position can change. Set --prop type=column (or bar, line, pie, area, scatter, doughnut) first to redraw it.");
                spec = name switch
                {
                    "type" => spec with { Type = value },
                    "series" => spec with { Series = XlsxChartXml.ParseSeries(value) },
                    "categories" => spec with { Categories = value.Trim().Length == 0 ? null : value.Trim() },
                    _ => spec with { Stacked = value == "true" },
                };
                chart.PlotArea = XlsxChartXml.PlotArea(doc, sheet, spec);
                break;
            case "x" or "y" or "w" or "h":
                var box = XlsxAnchors.Read(anchor, sheet.Grid);
                var emu = Math.Max(0, long.Parse(value, Inv));
                box = name switch { "x" => box with { X = emu }, "y" => box with { Y = emu }, "w" => box with { W = emu }, _ => box with { H = emu } };
                anchor = XlsxAnchors.Write(anchor, box, sheet.Grid);
                break;
        }
    }

    public override string GetRaw() => part.ChartSpace?.OuterXml ?? "";

    public override void SetRaw(string raw)
    {
        try { part.ChartSpace = new C.ChartSpace(RawXml.WithNamespaces(raw, doc.Namespaces)); }
        catch (Exception ex) when (ex is not WriterException)
        {
            throw new WriterException(ErrorCode.Validation, $"Raw chart XML could not be parsed: {ex.Message}", "Start from 'get --raw' output; the root must be c:chartSpace.");
        }
    }

    public override void Remove()
    {
        var drawings = sheet.Part.DrawingsPart!;
        anchor.Remove();
        drawings.DeletePart(part);
        if (drawings.WorksheetDrawing?.HasChildren == true) return;
        sheet.Part.Worksheet!.GetFirstChild<Drawing>()?.Remove();
        sheet.Part.DeletePart(drawings);
    }
}

/// <summary>Creates charts and finds the ones a sheet already has.</summary>
static class XlsxCharts
{
    const long DefaultWidth = 15 * Units.EmuPerCm;
    const long DefaultHeight = 15 * Units.EmuPerCm / 2;

    public static IEnumerable<XlsxChart> Of(XlsxDocument doc, XlsxSheet sheet)
    {
        var drawings = sheet.Part.DrawingsPart;
        if (drawings?.WorksheetDrawing is not { } drawing) yield break;
        foreach (var anchor in drawing.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            var rid = anchor.GetFirstChild<Xdr.GraphicFrame>()?.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>()?.Id?.Value;
            if (rid is not null && drawings.TryGetPartById(rid, out var p) && p is ChartPart chart)
                yield return new XlsxChart(doc, sheet, anchor, chart);
        }
    }

    public static XlsxChart Add(XlsxDocument doc, XlsxSheet sheet, IReadOnlyDictionary<string, string> props)
    {
        var spec = new ChartSpec(
            props.GetValueOrDefault("type") ?? "column",
            props.GetValueOrDefault("title") is { Length: > 0 } title ? title : null,
            props.GetValueOrDefault("categories") is { Length: > 0 } categories ? categories.Trim() : null,
            props.TryGetValue("series", out var series) ? XlsxChartXml.ParseSeries(series) : [],
            props.GetValueOrDefault("legend") ?? "bottom",
            props.GetValueOrDefault("stacked") == "true");
        var space = XlsxChartXml.ChartSpace(doc, sheet, spec);
        var drawings = Drawings(sheet);
        var drawing = drawings.WorksheetDrawing!;
        var chartPart = drawings.AddNewPart<ChartPart>();
        chartPart.ChartSpace = space;
        var id = drawing.Descendants<Xdr.NonVisualDrawingProperties>().Select(p => p.Id?.Value ?? 0u).DefaultIfEmpty(1u).Max() + 1;
        var box = DefaultBox(sheet, props);
        var anchor = XlsxChartXml.Fragment(new Xdr.TwoCellAnchor(RawXml.WithNamespaces(
            $"<xdr:twoCellAnchor>{XlsxAnchors.Markers(box, sheet.Grid)}<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{id}\" name=\"Chart {id}\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
            + "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
            + $"<c:chart r:id=\"{drawings.GetIdOfPart(chartPart)}\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>", XlsxAnchors.Namespaces)));
        drawing.Append(anchor);
        return new XlsxChart(doc, sheet, anchor, chartPart);
    }

    /// <summary>The sheet's drawing part, made when the sheet has none yet.</summary>
    internal static DrawingsPart Drawings(XlsxSheet sheet)
    {
        var part = sheet.Part;
        var drawings = part.DrawingsPart;
        if (drawings is null)
        {
            drawings = part.AddNewPart<DrawingsPart>();
            part.Worksheet!.AddChild(new Drawing { Id = part.GetIdOfPart(drawings) });
        }
        drawings.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
        return drawings;
    }

    /// <summary>Given position and size, or the default size (15 × 7.5 cm) to the right of the used range, below any chart already there.</summary>
    internal static XlsxAnchors.Box DefaultBox(XlsxSheet sheet, IReadOnlyDictionary<string, string> props, long width = DefaultWidth, long height = DefaultHeight)
    {
        var grid = sheet.Grid;
        var used = sheet.GetProps().GetValueOrDefault("range");
        var x = props.TryGetValue("x", out var xs) ? long.Parse(xs, CultureInfo.InvariantCulture) : used is null ? 0 : grid.X(XlsxCells.ParseRange(used).Col2 + 1, 0);
        long y = 0;
        if (props.TryGetValue("y", out var ys)) y = long.Parse(ys, CultureInfo.InvariantCulture);
        else
        {
            var charts = XlsxCharts.Of(sheet.Doc, sheet).Select(c => XlsxAnchors.Read((OpenXmlCompositeElement)c.Anchor, grid)).ToList();
            if (charts.Count > 0) y = charts.Max(b => b.Y + b.H) + grid.Row(0);
        }
        var w = props.TryGetValue("w", out var ws) ? long.Parse(ws, CultureInfo.InvariantCulture) : width;
        var h = props.TryGetValue("h", out var hs) ? long.Parse(hs, CultureInfo.InvariantCulture) : height;
        return new(Math.Max(0, x), Math.Max(0, y), Math.Max(1, w), Math.Max(1, h));
    }

    /// <summary>After a sheet rename, chart formulas and defined names still point at the old name; Excel would show #REF!.</summary>
    public static void RenameSheet(XlsxDocument doc, string oldName, string newName)
    {
        foreach (var (_, part) in doc.Sheets)
            foreach (var chart in part.DrawingsPart?.ChartParts ?? [])
                foreach (var f in chart.ChartSpace?.Descendants<C.Formula>() ?? [])
                    f.Text = XlsxRefs.Renamed(f.Text, oldName, newName);
        foreach (var name in doc.Workbook.Workbook!.DefinedNames?.Elements<DefinedName>() ?? [])
            name.Text = XlsxRefs.Renamed(name.Text, oldName, newName);
    }
}

/// <summary>Reads and writes the chart part: chartSpace → chart → plotArea with one chart element, its series, axes, title and legend.</summary>
static class XlsxChartXml
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    static readonly string[] Buildable = ["column", "bar", "line", "pie", "area", "scatter", "doughnut"];

    public static bool CanBuild(string type) => Buildable.Contains(type);

    public static ChartSpec Read(C.Chart? chart, string sheetName)
    {
        var charts = chart?.PlotArea?.ChildElements.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal)).ToList() ?? [];
        var first = charts.FirstOrDefault();
        var type = first is null ? "column" : first.LocalName switch
        {
            "barChart" or "bar3DChart" => Val(Child(first, "barDir")) == "bar" ? "bar" : "column",
            "lineChart" or "line3DChart" => "line",
            "pieChart" or "pie3DChart" or "ofPieChart" => "pie",
            "areaChart" or "area3DChart" => "area",
            "scatterChart" => "scatter",
            "doughnutChart" => "doughnut",
            var other => other[..^5],
        };
        var series = new List<ChartSeries>();
        foreach (var ser in charts.SelectMany(el => el.ChildElements.Where(s => s.LocalName == "ser")))
        {
            var tx = Child(ser, "tx");
            var name = Formula(tx) is { } f ? XlsxRefs.Display(f, sheetName) : Child(tx, "v")?.InnerText;
            var values = Formula(Child(ser, "val") ?? Child(ser, "yVal"));
            var x = Formula(Child(ser, "xVal"));
            series.Add(new ChartSeries(name, values is null ? "" : XlsxRefs.Display(values, sheetName), x is null ? null : XlsxRefs.Display(x, sheetName)));
        }
        string? categories = null;
        if (type == "scatter" || first?.LocalName == "bubbleChart")
        {
            if (series.Count > 0 && series[0].X is { } shared && series.All(s => s.X == shared))
            {
                categories = shared;
                series = series.Select(s => s with { X = null }).ToList();
            }
        }
        else if (first?.ChildElements.FirstOrDefault(s => s.LocalName == "ser") is { } firstSer && Formula(Child(firstSer, "cat")) is { } cf)
            categories = XlsxRefs.Display(cf, sheetName);
        var title = chart?.Title is { } t
            ? t.Descendants<A.Text>().Any() ? string.Concat(t.Descendants<A.Text>().Select(x => x.Text)) : string.Concat(t.Descendants<C.NumericValue>().Select(v => v.Text))
            : null;
        var legend = chart?.Legend is null ? "none" : Val(Child(chart.Legend, "legendPos")) switch
        {
            "b" => "bottom",
            "t" => "top",
            "l" => "left",
            _ => "right",
        };
        var stacked = Val(Child(first, "grouping")) is "stacked" or "percentStacked";
        return new ChartSpec(type, title, categories, series, legend, stacked);
    }

    static OpenXmlElement? Child(OpenXmlElement? parent, string localName) => parent?.ChildElements.FirstOrDefault(e => e.LocalName == localName);

    static string? Val(OpenXmlElement? element) => element?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value;

    static string? Formula(OpenXmlElement? parent) =>
        Child(parent?.ChildElements.FirstOrDefault(r => r.LocalName is "strRef" or "numRef" or "multiLvlStrRef"), "f")?.InnerText;

    public static string SeriesJson(IReadOnlyList<ChartSeries> series) => NodeJson.Compact(w =>
    {
        w.WriteStartArray();
        foreach (var s in series)
        {
            w.WriteStartObject();
            if (s.Name is not null) w.WriteString("name", s.Name);
            w.WriteString("values", s.Values);
            if (s.X is not null) w.WriteString("x", s.X);
            w.WriteEndObject();
        }
        w.WriteEndArray();
    });

    public static List<ChartSeries> ParseSeries(string json)
    {
        const string hint = "Example: [{\"name\":\"B1\",\"values\":\"B2:B6\"}]. name is a cell or a text, values a range; scatter series may add x, an X range.";
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Array)
            throw new WriterException(ErrorCode.Validation, "series must be a JSON array of objects", hint);
        var list = new List<ChartSeries>();
        foreach (var e in parsed.RootElement.EnumerateArray())
        {
            if (e.ValueKind != JsonValueKind.Object || Text(e, "values") is not { Length: > 0 } values)
                throw new WriterException(ErrorCode.Validation, "Every series needs a values range", hint);
            list.Add(new ChartSeries(Text(e, "name"), values.Trim(), Text(e, "x")?.Trim()));
        }
        return list;
    }

    static string? Text(JsonElement e, string name) => e.TryGetProperty(name, out var p) ? p.ValueKind switch
    {
        JsonValueKind.String => p.GetString() is { Length: > 0 } s ? s : null,
        JsonValueKind.Number => p.GetRawText(),
        _ => null,
    } : null;

    public static C.ChartSpace ChartSpace(XlsxDocument doc, XlsxSheet sheet, ChartSpec spec) => new(RawXml.WithNamespaces(
        "<c:chartSpace><c:roundedCorners val=\"0\"/><c:chart>" + Title(spec.Title) + $"<c:autoTitleDeleted val=\"{(spec.Title is null ? 1 : 0)}\"/>"
        + PlotAreaXml(doc, sheet, spec) + (spec.Legend == "none" ? "" : Legend(spec.Legend))
        + "<c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"gap\"/></c:chart></c:chartSpace>", []));

    public static C.PlotArea PlotArea(XlsxDocument doc, XlsxSheet sheet, ChartSpec spec) => Fragment(new C.PlotArea(RawXml.WithNamespaces(PlotAreaXml(doc, sheet, spec), [])));

    /// <summary>A parsed fragment without the xmlns declarations parsing needed; the part's root declares them when saved.</summary>
    public static T Fragment<T>(T element) where T : OpenXmlElement
    {
        foreach (var ns in element.NamespaceDeclarations.ToList()) element.RemoveNamespaceDeclaration(ns.Key);
        return element;
    }

    static string Title(string? title) => title is null ? "" :
        $"<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>{Esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";

    static string Legend(string legend) => $"<c:legend><c:legendPos val=\"{Pos(legend)}\"/><c:overlay val=\"0\"/></c:legend>";

    static string Pos(string legend) => legend switch { "right" => "r", "left" => "l", "top" => "t", _ => "b" };

    public static void SetTitle(C.Chart chart, string title)
    {
        chart.Title = null;
        if (title.Length > 0) chart.InsertAt(Fragment(new C.Title(RawXml.WithNamespaces(Title(title), []))), 0);
        var auto = chart.AutoTitleDeleted;
        if (auto is null) chart.InsertAt(auto = new C.AutoTitleDeleted(), title.Length > 0 ? 1 : 0);
        auto.Val = title.Length == 0;
    }

    public static void SetLegend(C.Chart chart, string legend)
    {
        if (legend == "none")
        {
            chart.Legend = null;
            return;
        }
        if (chart.Legend is not { } existing)
        {
            var created = Fragment(new C.Legend(RawXml.WithNamespaces(Legend(legend), [])));
            if (chart.PlotArea is { } plot) plot.InsertAfterSelf(created);
            else chart.Append(created);
            return;
        }
        var value = legend switch
        {
            "right" => C.LegendPositionValues.Right,
            "left" => C.LegendPositionValues.Left,
            "top" => C.LegendPositionValues.Top,
            _ => C.LegendPositionValues.Bottom,
        };
        if (existing.LegendPosition is { } position) position.Val = value;
        else existing.PrependChild(new C.LegendPosition { Val = value });
    }

    static string PlotAreaXml(XlsxDocument doc, XlsxSheet sheet, ChartSpec spec)
    {
        if (!CanBuild(spec.Type))
            throw new WriterException(ErrorCode.Validation, $"'{spec.Type}' is not a chart type writer can draw", "Types: " + string.Join(" | ", Buildable) + ".");
        var stacked = spec.Stacked && spec.Type is "column" or "bar" or "line" or "area";
        var sers = string.Concat(spec.Series.Select((s, i) => Series(doc, sheet, spec, s, i)));
        var body = spec.Type switch
        {
            "column" or "bar" => $"<c:barChart><c:barDir val=\"{(spec.Type == "bar" ? "bar" : "col")}\"/><c:grouping val=\"{(stacked ? "stacked" : "clustered")}\"/><c:varyColors val=\"0\"/>{sers}"
                + $"<c:gapWidth val=\"150\"/>{(stacked ? "<c:overlap val=\"100\"/>" : "")}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>",
            "line" => $"<c:lineChart><c:grouping val=\"{(stacked ? "stacked" : "standard")}\"/><c:varyColors val=\"0\"/>{sers}<c:marker val=\"1\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:lineChart>",
            "area" => $"<c:areaChart><c:grouping val=\"{(stacked ? "stacked" : "standard")}\"/><c:varyColors val=\"0\"/>{sers}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:areaChart>",
            "pie" => $"<c:pieChart><c:varyColors val=\"1\"/>{sers}<c:firstSliceAng val=\"0\"/></c:pieChart>",
            "doughnut" => $"<c:doughnutChart><c:varyColors val=\"1\"/>{sers}<c:firstSliceAng val=\"0\"/><c:holeSize val=\"50\"/></c:doughnutChart>",
            _ => $"<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:varyColors val=\"0\"/>{sers}<c:axId val=\"1\"/><c:axId val=\"2\"/></c:scatterChart>",
        };
        var axes = spec.Type switch
        {
            "pie" or "doughnut" => "",
            "scatter" => ValueAxis(1, "b", 2, false, "midCat") + ValueAxis(2, "l", 1, true, "midCat"),
            "bar" => CategoryAxis("l") + ValueAxis(2, "b", 1, true, "between"),
            _ => CategoryAxis("b") + ValueAxis(2, "l", 1, true, "between"),
        };
        return "<c:plotArea><c:layout/>" + body + axes + "</c:plotArea>";
    }

    static string CategoryAxis(string pos) =>
        $"<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"{pos}\"/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>"
        + "<c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"2\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/>"
        + "<c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/><c:noMultiLvlLbl val=\"0\"/></c:catAx>";

    static string ValueAxis(int id, string pos, int crossAx, bool gridlines, string crossBetween) =>
        $"<c:valAx><c:axId val=\"{id}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"{pos}\"/>{(gridlines ? "<c:majorGridlines/>" : "")}"
        + $"<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{crossAx}\"/>"
        + $"<c:crosses val=\"autoZero\"/><c:crossBetween val=\"{crossBetween}\"/></c:valAx>";

    static string Series(XlsxDocument doc, XlsxSheet sheet, ChartSpec spec, ChartSeries s, int i)
    {
        var sb = new StringBuilder($"<c:ser><c:idx val=\"{i}\"/><c:order val=\"{i}\"/>");
        if (s.Name is { } name)
            sb.Append("<c:tx>").Append(XlsxRefs.IsCell(doc, name) ? StringRef(doc, sheet, name, $"series[{i}].name") : $"<c:v>{Esc(name)}</c:v>").Append("</c:tx>");
        if (spec.Type is "column" or "bar") sb.Append("<c:invertIfNegative val=\"0\"/>");
        if (spec.Type == "scatter")
        {
            sb.Append("<c:spPr><a:ln w=\"19050\"><a:noFill/></a:ln></c:spPr>");
            if ((s.X ?? spec.Categories) is { } x) sb.Append("<c:xVal>").Append(NumberRef(doc, sheet, x, s.X is null ? "categories" : $"series[{i}].x")).Append("</c:xVal>");
            sb.Append("<c:yVal>").Append(NumberRef(doc, sheet, s.Values, $"series[{i}].values")).Append("</c:yVal><c:smooth val=\"0\"/>");
        }
        else
        {
            if (spec.Categories is { } categories) sb.Append("<c:cat>").Append(StringRef(doc, sheet, categories, "categories")).Append("</c:cat>");
            sb.Append("<c:val>").Append(NumberRef(doc, sheet, s.Values, $"series[{i}].values")).Append("</c:val>");
            if (spec.Type == "line") sb.Append("<c:smooth val=\"0\"/>");
        }
        return sb.Append("</c:ser>").ToString();
    }

    static string StringRef(XlsxDocument doc, XlsxSheet sheet, string reference, string prop)
    {
        var (name, box) = XlsxRefs.Parse(doc, sheet.SheetName, reference, prop);
        var values = Cells(doc, name, box);
        var points = values.Select((v, i) => v.Length == 0 ? "" : $"<c:pt idx=\"{i}\"><c:v>{Esc(v)}</c:v></c:pt>");
        return $"<c:strRef><c:f>{Esc(XlsxRefs.Formula(name, box))}</c:f><c:strCache><c:ptCount val=\"{values.Count}\"/>{string.Concat(points)}</c:strCache></c:strRef>";
    }

    static string NumberRef(XlsxDocument doc, XlsxSheet sheet, string reference, string prop)
    {
        var (name, box) = XlsxRefs.Parse(doc, sheet.SheetName, reference, prop);
        var values = Cells(doc, name, box);
        var points = values.Select((v, i) => double.TryParse(v, NumberStyles.Float, Inv, out var d) ? $"<c:pt idx=\"{i}\"><c:v>{d.ToString("R", Inv)}</c:v></c:pt>" : "");
        return $"<c:numRef><c:f>{Esc(XlsxRefs.Formula(name, box))}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"{values.Count}\"/>{string.Concat(points)}</c:numCache></c:numRef>";
    }

    /// <summary>Displayed values of a range, row by row.</summary>
    static List<string> Cells(XlsxDocument doc, string sheetName, (int Col1, int Row1, int Col2, int Row2) box)
    {
        var data = doc.Sheets.First(s => string.Equals(s.Sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase)).Part.Worksheet?.GetFirstChild<SheetData>();
        var values = new List<string>();
        for (var r = box.Row1; r <= box.Row2; r++)
        {
            var row = data is null ? null : XlsxCells.FindRow(data, r);
            for (var c = box.Col1; c <= box.Col2; c++)
            {
                var cell = row is null ? null : XlsxCells.FindCell(row, c);
                values.Add(cell is null ? "" : XlsxCells.Display(doc, cell));
            }
        }
        return values;
    }

    static string Esc(string s) => SecurityElement.Escape(s);
}

/// <summary>Where a drawing sits on the grid: twoCellAnchor, oneCellAnchor or absoluteAnchor ↔ x, y, w, h in EMU.</summary>
static class XlsxAnchors
{
    public static readonly KeyValuePair<string, string>[] Namespaces = [new("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")];

    public sealed record Box(long X, long Y, long W, long H);

    public static Box Read(OpenXmlCompositeElement anchor, XlsxLayout.Grid grid)
    {
        switch (anchor)
        {
            case Xdr.TwoCellAnchor two:
                var (x1, y1) = Point(two.FromMarker, grid);
                var (x2, y2) = Point(two.ToMarker, grid);
                return new(x1, y1, Math.Max(0, x2 - x1), Math.Max(0, y2 - y1));
            case Xdr.OneCellAnchor one:
                var (x, y) = Point(one.FromMarker, grid);
                return new(x, y, one.Extent?.Cx?.Value ?? 0, one.Extent?.Cy?.Value ?? 0);
            case Xdr.AbsoluteAnchor absolute:
                return new(absolute.Position?.X?.Value ?? 0, absolute.Position?.Y?.Value ?? 0, absolute.Extent?.Cx?.Value ?? 0, absolute.Extent?.Cy?.Value ?? 0);
            default:
                return new(0, 0, 0, 0);
        }
    }

    static (long X, long Y) Point(Xdr.MarkerType? marker, XlsxLayout.Grid grid) => marker is null
        ? (0, 0)
        : (grid.X((int)Number(marker.ColumnId), Number(marker.ColumnOffset)), grid.Y((int)Number(marker.RowId), Number(marker.RowOffset)));

    static long Number(OpenXmlLeafTextElement? element) =>
        element is not null && long.TryParse(element.Text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var n) ? n : 0;

    /// <summary>Moves the anchor to the box; a one-cell or absolute anchor becomes a two-cell anchor. Returns the anchor now in the drawing.</summary>
    public static OpenXmlCompositeElement Write(OpenXmlCompositeElement anchor, Box box, XlsxLayout.Grid grid)
    {
        var markers = Markers(box, grid);
        if (anchor is Xdr.TwoCellAnchor two)
        {
            two.FromMarker = XlsxChartXml.Fragment(new Xdr.FromMarker(RawXml.WithNamespaces(markers[..markers.IndexOf("<xdr:to>", StringComparison.Ordinal)], Namespaces)));
            two.ToMarker = XlsxChartXml.Fragment(new Xdr.ToMarker(RawXml.WithNamespaces(markers[markers.IndexOf("<xdr:to>", StringComparison.Ordinal)..], Namespaces)));
            return two;
        }
        var replacement = XlsxChartXml.Fragment(new Xdr.TwoCellAnchor(RawXml.WithNamespaces("<xdr:twoCellAnchor>" + markers + "</xdr:twoCellAnchor>", Namespaces)));
        foreach (var child in anchor.ChildElements.Where(c => c is not (Xdr.FromMarker or Xdr.Position or Xdr.Extent)).ToList())
        {
            child.Remove();
            replacement.Append(child);
        }
        anchor.Parent!.ReplaceChild(replacement, anchor);
        return replacement;
    }

    public static string Markers(Box box, XlsxLayout.Grid grid)
    {
        var (c1, cOff1) = grid.ColumnAt(box.X);
        var (r1, rOff1) = grid.RowAt(box.Y);
        var (c2, cOff2) = grid.ColumnAt(box.X + box.W);
        var (r2, rOff2) = grid.RowAt(box.Y + box.H);
        return $"<xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>{cOff1}</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>{rOff1}</xdr:rowOff></xdr:from>"
            + $"<xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>{cOff2}</xdr:colOff><xdr:row>{r2}</xdr:row><xdr:rowOff>{rOff2}</xdr:rowOff></xdr:to>";
    }
}
