using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Xlsx;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Writer.Formats.Common;

// Word and PowerPoint charts use the same native chart and embedded workbook as Excel.
static class OfficeChartData
{
    internal static void Write(ChartPart part, JsonElement o)
    {
        string Text(string key, string fallback = "") => o.TryGetProperty(key, out var v) ? v.GetString() ?? fallback : fallback;
        bool Flag(string key) => o.TryGetProperty(key, out var v) && v.ValueKind == JsonValueKind.True;
        var type = Text("type", "column");
        if (!XlsxChartXml.CanBuild(type)) throw Bad("Unsupported chart type");
        var labels = o.GetProperty("labels").EnumerateArray().Select(v => v.ToString()).ToArray();
        var data = o.TryGetProperty("series", out var many) ? many.EnumerateArray().ToArray() : [o];
        if (labels.Length == 0 || labels.Length > 1048575 || data.Length == 0 || data.Length * 2 + 1 > 16384) throw Bad("Chart needs labels and at least one series");
        var ys = data.Select(s => s.GetProperty("values").EnumerateArray().Select(v => v.ValueKind == JsonValueKind.Null ? (double?)null : v.GetDouble()).ToArray()).ToArray();
        var xs = data.Select(s => s.TryGetProperty("x", out var x) ? x.EnumerateArray().Select(v => v.GetDouble()).ToArray() : null).ToArray();
        if (ys.Any(v => v.Length != labels.Length || v.Any(n => n is { } number && !double.IsFinite(number))) || xs.Any(v => v is not null && (v.Length != labels.Length || v.Any(n => !double.IsFinite(n))))) throw Bad("Each series needs one finite value or blank per label");
        var names = data.Select((s, i) => s.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "Series " + (i + 1)).ToArray();
        var kinds = data.Select(s => s.TryGetProperty("kind", out var k) ? k.GetString() : null).ToArray();
        if (type == "combo" && kinds.Any(k => k is not (null or "column" or "line" or "area"))) throw Bad("Combination series use column, line or area");
        using var workbook = (XlsxDocument)new XlsxAdapter().Create();
        var sheet = (XlsxSheet)workbook.Root.Children[0];
        var rows = NodeJson.Compact(w => {
            w.WriteStartArray(); w.WriteStartArray(); w.WriteStringValue("Category");
            for (var j = 0; j < data.Length; j++) { w.WriteStringValue(names[j]); w.WriteStringValue(names[j] + " X"); }
            w.WriteEndArray();
            for (var i = 0; i < labels.Length; i++) {
                w.WriteStartArray(); w.WriteStringValue(labels[i]);
                for (var j = 0; j < data.Length; j++) { if (ys[j][i] is { } n) w.WriteNumberValue(n); else w.WriteNullValue(); w.WriteNumberValue(xs[j]?[i] ?? (double.TryParse(labels[i], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var x) ? x : i + 1)); }
                w.WriteEndArray();
            }
            w.WriteEndArray();
        });
        Mutations.Set(PathResolver.Single(workbook.Root, $"/sheet[1]/range[A1:{XlsxCells.Reference(data.Length * 2 + 1, labels.Length + 1)}]"), new Dictionary<string, string> { ["values"] = rows });
        // Category/series text may begin with '='; it is data, never an executable formula.
        void TextCell(int c,int r,string value) { var cell=XlsxCells.GetOrCreateCell(sheet.Data,c,r);cell.CellFormula=null;XlsxCells.SetString(workbook,cell,value); }
        for (var i = 0; i < labels.Length; i++) TextCell(1, i + 2, labels[i]);
        for (var i = 0; i < data.Length; i++) { TextCell(i * 2 + 2, 1, names[i]);TextCell(i * 2 + 3, 1, names[i]+" X"); }
        var series = data.Select((_, i) => new ChartSeries(XlsxCells.Reference(i * 2 + 2, 1), $"{XlsxCells.Reference(i * 2 + 2, 2)}:{XlsxCells.Reference(i * 2 + 2, labels.Length + 1)}", type == "scatter" ? $"{XlsxCells.Reference(i * 2 + 3, 2)}:{XlsxCells.Reference(i * 2 + 3, labels.Length + 1)}" : null, kinds[i])).ToArray();
        var legend = Text("legend", "bottom");
        if (legend is not ("none" or "left" or "right" or "top" or "bottom" or "topRight")) throw Bad("Invalid legend position");
        var spec = new ChartSpec(type, Text("title"), $"A2:A{labels.Length + 1}", series, legend, Flag("stacked") || Flag("percentStacked"), Flag("percentStacked"), Flag("dataLabels"), Text("xTitle"), Text("yTitle"));
        var fresh = XlsxChartXml.ChartSpace(workbook, sheet, spec);
        var old = part.ChartSpace?.GetFirstChild<C.Chart>();
        if (old is not null) {
            XlsxChartXml.KeepAppearance(old.PlotArea, fresh.GetFirstChild<C.Chart>()!.PlotArea!);
            old.PlotArea = (C.PlotArea)fresh.GetFirstChild<C.Chart>()!.PlotArea!.CloneNode(true);
            XlsxChartXml.SetTitle(old, spec.Title ?? ""); XlsxChartXml.SetLegend(old, legend);
        } else part.ChartSpace = fresh;
        var space = part.ChartSpace!;
        var embedded = space.GetFirstChild<C.ExternalData>()?.Id?.Value is { } rid && part.TryGetPartById(rid, out var prior) && prior is EmbeddedPackagePart package ? package : part.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        using (var stream = new MemoryStream()) { workbook.Save(stream); stream.Position = 0; embedded.FeedData(stream); }
        space.RemoveAllChildren<C.ExternalData>();
        space.AddChild(new C.ExternalData(new C.AutoUpdate { Val = false }) { Id = part.GetIdOfPart(embedded) });
        space.Save();
    }
    static WriterException Bad(string message) => new(ErrorCode.Validation, message, "Use {type,title,labels,series:[{name,values,x?,kind?}],legend,stacked,dataLabels,xTitle,yTitle}.");
}
