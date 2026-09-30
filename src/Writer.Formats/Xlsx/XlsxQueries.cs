using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;

namespace Writer.Formats.Xlsx;

// Keep refreshable Writer query definitions in a separate custom XML part. Excel can
// read/edit the result cells without losing these definitions or its own Power Query parts.
static class XlsxQueries
{
    static readonly XNamespace Ns = "urn:writer:spreadsheet-queries:v1";
    static CustomXmlPart? Find(XlsxDocument doc) => doc.Workbook.CustomXmlParts.FirstOrDefault(p => { try { using var s=p.GetStream(); return XDocument.Load(s).Root?.Name == Ns+"queries"; } catch(System.Xml.XmlException) { return false; } });
    internal static string Read(XlsxDocument doc)
    {
        var part=Find(doc);if(part is null)return "[]";using var stream=part.GetStream();return XDocument.Load(stream).Root!.Value;
    }
    internal static void Write(XlsxDocument doc,string json)
    {
        var list=JsonNode.Parse(json) as JsonArray ?? throw Bad("Queries must be an array");var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach(var q in list) { var name=q?["name"]?.GetValue<string>();if(string.IsNullOrWhiteSpace(name)||!names.Add(name)||q?["source"] is not JsonObject||q?["steps"] is not JsonArray)throw Bad("Each query needs a unique name, source and steps"); }
        var part=Find(doc);if(list.Count==0){if(part is not null)doc.Workbook.DeletePart(part);return;}
        part??=doc.Workbook.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        using var stream=new MemoryStream(Encoding.UTF8.GetBytes(new XElement(Ns+"queries",list.ToJsonString()).ToString(SaveOptions.DisableFormatting)));part.FeedData(stream);
    }
    static WriterException Bad(string message)=>new(ErrorCode.Validation,message,"Use [{name,source,steps,targetSheet,target,range?}].");
}
