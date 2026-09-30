using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

static class PptxCharts
{
    internal static void SmartArt(PptxDocument doc, SlidePart slide, string json)
    {
        using var parsed=JsonDocument.Parse(json);var spec=parsed.RootElement;
        if(spec.TryGetProperty("path",out var path)){
            var node=PathResolver.Single(doc.Root,path.GetString()!);if(node is not PptxObject||node.GetProps().GetValueOrDefault("type")!="smartart")throw new WriterException(ErrorCode.Validation,"Select SmartArt","Choose a SmartArt object.");
            var frame=(P.GraphicFrame)node.Anchor;if(frame.Ancestors<P.Slide>().FirstOrDefault()!=slide.Slide)throw new WriterException(ErrorCode.Validation,"SmartArt belongs to another slide","Select its slide first.");
            var old=frame.Graphic!.GraphicData!;var next=OfficeSmartArt.Write(slide,spec,old);if(!ReferenceEquals(old,next))frame.Graphic.GraphicData=next;return;
        }
        var data=OfficeSmartArt.Write(slide,spec);var tree=PptxDocument.ShapeTree(slide);var id=tree.Descendants<P.NonVisualDrawingProperties>().Select(p=>p.Id?.Value??0u).DefaultIfEmpty().Max()+1;
        tree.Append(new P.GraphicFrame(new P.NonVisualGraphicFrameProperties(new P.NonVisualDrawingProperties{Id=id,Name="SmartArt "+id},new P.NonVisualGraphicFrameDrawingProperties(),new P.ApplicationNonVisualDrawingProperties()),new P.Transform(new A.Offset{X=doc.SlideSize.Width/10,Y=doc.SlideSize.Height/6},new A.Extents{Cx=doc.SlideSize.Width*4/5,Cy=doc.SlideSize.Height*2/3}),new A.Graphic(data)));
    }
    internal static void Write(PptxDocument doc, SlidePart slide, string json)
    {
        using var parsed = JsonDocument.Parse(json); var o = parsed.RootElement;
        if (o.TryGetProperty("path", out var path)) {
            var node = PathResolver.Single(doc.Root, path.GetString()!);
            if (node is not PptxObject || node.GetProps().GetValueOrDefault("type") != "chart") throw new WriterException(ErrorCode.Validation, "Select a chart", "Choose a chart on this slide.");
            var frame = (P.GraphicFrame)node.Anchor;
            if (frame.Ancestors<P.Slide>().FirstOrDefault() != slide.Slide) throw new WriterException(ErrorCode.Validation, "Chart belongs to another slide", "Select its slide first.");
            if (frame.Descendants<C.ChartReference>().FirstOrDefault()?.Id?.Value is not { } id || slide.GetPartById(id) is not ChartPart existing) throw new WriterException(ErrorCode.Validation, "Unsupported chart part", "This chart format cannot be edited here.");
            OfficeChartData.Write(existing, o); return;
        }
        var part = slide.AddNewPart<ChartPart>(); OfficeChartData.Write(part, o);
        var tree = PptxDocument.ShapeTree(slide);
        var nextId = tree.Descendants<P.NonVisualDrawingProperties>().Select(p => p.Id?.Value ?? 0).DefaultIfEmpty(0u).Max() + 1;
        tree.Append(new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(new P.NonVisualDrawingProperties { Id = nextId, Name = "Chart " + nextId }, new P.NonVisualGraphicFrameDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(new A.Offset { X = doc.SlideSize.Width / 10, Y = doc.SlideSize.Height / 6 }, new A.Extents { Cx = doc.SlideSize.Width * 4 / 5, Cy = doc.SlideSize.Height * 2 / 3 }),
            new A.Graphic(new A.GraphicData(new C.ChartReference { Id = slide.GetIdOfPart(part) }) { Uri = OfficeGraphics.ChartUri })));
    }
}
