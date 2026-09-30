using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;

namespace Writer.Formats.Xlsx;

// Matching rows use native row visibility; this small extension keeps Writer's
// criteria and the distinction from rows hidden manually for reopen/reapply.
static class XlsxAdvancedFilter
{
    static readonly XNamespace Ns = "urn:writer:advanced-filter:v1";
    static CustomXmlPart? Find(XlsxDocument doc) => doc.Workbook.CustomXmlParts.FirstOrDefault(p => {
        try { using var s=p.GetStream(); return XDocument.Load(s).Root?.Name==Ns+"filters"; }
        catch(System.Xml.XmlException) { return false; }
    });
    static XElement Load(CustomXmlPart? part) { if(part is null)return new XElement(Ns+"filters");using var s=part.GetStream();return XDocument.Load(s).Root!; }
    public static string Read(XlsxSheet sheet) => Load(Find(sheet.Doc)).Elements(Ns+"sheet").FirstOrDefault(e=>(string?)e.Attribute("id")==sheet.Doc.Workbook.GetIdOfPart(sheet.Part))?.Value??"null";
    public static void Write(XlsxSheet sheet,string json) {
        var spec=JsonNode.Parse(json);
        if(spec is not null) {
            if(spec is not JsonObject||string.IsNullOrWhiteSpace(spec["range"]?.GetValue<string>()))throw Bad();
            foreach(var key in new[]{"range","criteria"})if(spec[key]?.GetValue<string>() is {Length:>0} value) {
                var ends=value.Replace("$","").Split(':');
                if(ends.Length>2||ends.Any(s=>!System.Text.RegularExpressions.Regex.IsMatch(s,"^[A-Z]{1,3}[1-9][0-9]*$")))throw Bad();
                foreach(var end in ends){var (col,row)=XlsxCells.Parse(end);if(row<1||row>1048576||col<1||col>16384)throw Bad();}
            }
            foreach(var key in new[]{"hidden","manualHidden"})if(spec[key] is {} rows) {
                if(rows is not JsonArray hidden||hidden.Any(v=>v is not JsonValue j||!j.TryGetValue<int>(out var n)||n<0||n>=1048576))throw Bad();
            }
            if(spec["unique"] is {} unique&&(unique is not JsonValue flag||!flag.TryGetValue<bool>(out _)))throw Bad();
        }
        var part=Find(sheet.Doc);var root=Load(part);var id=sheet.Doc.Workbook.GetIdOfPart(sheet.Part);
        root.Elements(Ns+"sheet").Where(e=>(string?)e.Attribute("id")==id).Remove();
        if(spec is not null)root.Add(new XElement(Ns+"sheet",new XAttribute("id",id),spec.ToJsonString()));
        if(!root.HasElements){if(part is not null)sheet.Doc.Workbook.DeletePart(part);return;}
        part??=sheet.Doc.Workbook.AddCustomXmlPart(CustomXmlPartType.CustomXml);using var stream=new MemoryStream(Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting)));part.FeedData(stream);
    }
    static WriterException Bad()=>new(ErrorCode.Validation,"Invalid advanced filter configuration","Use {range,criteria,unique,hidden:[zeroBasedRows]} or null.");
}
