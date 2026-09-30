using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
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
        var title = o.TryGetProperty("title", out var tt) ? tt.GetString() ?? "" : "";
        W.Drawing? existing = null;
        if (o.TryGetProperty("path", out var path) && path.GetString() is { Length: > 0 } target)
        {
            var node = PathResolver.Single(doc.Root, target);
            if (node is not DocxObject || node.GetProps().GetValueOrDefault("type") != "chart") throw new WriterException(ErrorCode.Validation, "Select a chart object", "The selected object is not a chart.");
            var anchor = (OpenXmlElement)node.Anchor;
            existing = anchor as W.Drawing ?? anchor.Descendants<W.Drawing>().FirstOrDefault();
        }
        if (existing?.Descendants<C.ChartReference>().FirstOrDefault()?.Id?.Value is { } chartId && doc.Main.TryGetPartById(chartId, out var oldPart) && oldPart is ChartPart oldChart)
        {
            OfficeChartData.Write(oldChart, o);
            return;
        }
        var part = doc.Main.AddNewPart<ChartPart>();
        OfficeChartData.Write(part, o);
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
