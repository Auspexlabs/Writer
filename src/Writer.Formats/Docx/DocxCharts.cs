using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Xlsx;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

static class DocxCharts
{
    public static void Write(DocxDocument doc, string json)
    {
        using var data = JsonDocument.Parse(json); var o = data.RootElement;
        var type = o.TryGetProperty("type", out var t) ? t.GetString() ?? "column" : "column";
        if (type is not ("column" or "bar" or "line" or "area" or "pie" or "doughnut")) throw new WriterException(ErrorCode.Validation, "Unsupported chart type", "Use column, bar, line, area, pie or doughnut.");
        var categories = o.GetProperty("labels").EnumerateArray().Select(v => v.ToString()).ToArray();
        var values = o.GetProperty("values").EnumerateArray().Select(v => v.GetDouble()).ToArray();
        if (categories.Length == 0 || categories.Length != values.Length || values.Any(v => !double.IsFinite(v))) throw new WriterException(ErrorCode.Validation, "Chart labels and values must have the same nonzero length", "Provide one numeric value per label.");
        var title = o.TryGetProperty("title", out var tt) ? tt.GetString() ?? "" : "";
        W.Drawing? existing = null;
        if (o.TryGetProperty("path", out var path) && path.GetString() is { Length: > 0 } target)
        {
            var node = PathResolver.Single(doc.Root, target);
            if (node is not DocxObject || node.GetProps().GetValueOrDefault("type") != "chart") throw new WriterException(ErrorCode.Validation, "Select a chart object", "The selected object is not a chart.");
            var anchor = (OpenXmlElement)node.Anchor;
            existing = anchor as W.Drawing ?? anchor.Descendants<W.Drawing>().FirstOrDefault();
        }
        using var workbook = (XlsxDocument)new XlsxAdapter().Create();
        var sheet = (XlsxSheet)workbook.Root.Children[0];
        var rows = NodeJson.Compact(w => { w.WriteStartArray(); w.WriteStartArray(); w.WriteStringValue("Category"); w.WriteStringValue(title.Length == 0 ? "Value" : title); w.WriteEndArray(); for (var i = 0; i < categories.Length; i++) { w.WriteStartArray(); w.WriteStringValue(categories[i]); w.WriteNumberValue(values[i]); w.WriteEndArray(); } w.WriteEndArray(); });
        Mutations.Set(PathResolver.Single(workbook.Root, $"/sheet[1]/range[A1:B{categories.Length + 1}]"), new Dictionary<string, string> { ["values"] = rows });
        var spec = new ChartSpec(type, title, $"A2:A{categories.Length + 1}", [new ChartSeries("B1", $"B2:B{categories.Length + 1}", null)], "bottom", false);
        var space = XlsxChartXml.ChartSpace(workbook, sheet, spec);
        var part = doc.Main.AddNewPart<ChartPart>(); part.ChartSpace = space;
        var embedded = part.AddNewPart<EmbeddedPackagePart>("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        using (var stream = new MemoryStream()) { workbook.Save(stream); stream.Position = 0; embedded.FeedData(stream); }
        space.AddChild(new C.ExternalData(new C.AutoUpdate { Val = false }) { Id = part.GetIdOfPart(embedded) });
        space.Save();
        var id = doc.Main.Document!.Descendants<DW.DocProperties>().Select(p => p.Id?.Value ?? 0u).DefaultIfEmpty().Max() + 1;
        var drawing = new W.Drawing(new DW.Inline(
            new DW.Extent { Cx = 5486400L, Cy = 3086100L },
            new DW.DocProperties { Id = id, Name = "Chart " + id, Description = title },
            new A.Graphic(new A.GraphicData(new C.ChartReference { Id = doc.Main.GetIdOfPart(part) }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })));
        if (existing is not null)
        {
            // Preserve its placement, wrapping and identity, replacing only the chart data.
            var graphic = existing.Descendants<A.Graphic>().First();
            var oldId = graphic.Descendants<C.ChartReference>().FirstOrDefault()?.Id?.Value;
            var replacement = drawing.Descendants<A.Graphic>().First(); replacement.Remove(); graphic.InsertAfterSelf(replacement); graphic.Remove();
            if (oldId is not null && !doc.Main.Document.Descendants<C.ChartReference>().Any(r => r.Id?.Value == oldId)) doc.Main.DeletePart(oldId);
        }
        else
        {
            var body = doc.Root.Children.Single();
            DocxBlocks.InsertAt(body, doc.Main.Document.Body!, new W.Paragraph(new W.Run(drawing)), o.TryGetProperty("index", out var at) ? at.GetInt32() : null);
        }
    }
}
