using System.Text.Json;
using DocumentFormat.OpenXml;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;
static class DocxSmartArt
{
    internal static void Write(DocxDocument doc,string json)
    {
        using var parsed=JsonDocument.Parse(json);var spec=parsed.RootElement;
        if(spec.TryGetProperty("path",out var path)){
            var node=PathResolver.Single(doc.Root,path.GetString()!);if(node is not DocxObject||node.GetProps().GetValueOrDefault("type")!="smartart")throw new WriterException(ErrorCode.Validation,"Select SmartArt","Choose a SmartArt object.");
            var anchor=(OpenXmlElement)node.Anchor;var graphic=anchor.Descendants<A.Graphic>().First();var old=graphic.GraphicData!;var next=OfficeSmartArt.Write(doc.Main,spec,old);if(!ReferenceEquals(old,next))graphic.GraphicData=next;return;
        }
        var data=OfficeSmartArt.Write(doc.Main,spec);var id=doc.Main.Document!.Descendants<DW.DocProperties>().Select(p=>p.Id?.Value??0u).DefaultIfEmpty().Max()+1;
        var drawing=new W.Drawing(new DW.Inline(new DW.Extent{Cx=5486400,Cy=3086100},new DW.DocProperties{Id=id,Name="SmartArt "+id},new A.Graphic(data)));
        DocxBlocks.InsertAt(doc.Root.Children.Single(),doc.Main.Document.Body!,new W.Paragraph(new W.Run(drawing)),spec.TryGetProperty("index",out var at)?at.GetInt32():null);
    }
}
