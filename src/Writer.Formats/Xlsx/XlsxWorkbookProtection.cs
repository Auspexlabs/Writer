using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

static class XlsxWorkbookProtection
{
    static readonly Dictionary<string,string> Attributes = new()
    {
        ["password"]="workbookPassword", ["algorithmName"]="workbookAlgorithmName",
        ["hashValue"]="workbookHashValue", ["saltValue"]="workbookSaltValue", ["spinCount"]="workbookSpinCount"
    };
    internal static string Read(Workbook book)
    {
        if (book.GetFirstChild<WorkbookProtection>() is not { } p) return "{}";
        var result=new JsonObject { ["structure"]=p.LockStructure?.Value??false };
        var verifier=new JsonObject();
        foreach(var (key,attribute) in Attributes)
        {
            var value=p.GetAttributes().FirstOrDefault(a=>a.LocalName==attribute&&a.NamespaceUri=="").Value;
            if(!string.IsNullOrEmpty(value)) verifier[key]=key=="spinCount"?JsonValue.Create(int.Parse(value,System.Globalization.CultureInfo.InvariantCulture)):JsonValue.Create(value);
        }
        if(verifier.Count>0) result["verifier"]=verifier;
        return result.ToJsonString();
    }
    internal static void Write(Workbook book,string json)
    {
        var value=JsonNode.Parse(json) as JsonObject??throw Invalid();
        foreach(var (key,node) in value)
            if(key!="verifier"&&(key!="structure"||node is not JsonValue j||!j.TryGetValue<bool>(out _)))throw Invalid();
        var same=JsonNode.DeepEquals(value["verifier"],JsonNode.Parse(Read(book))?["verifier"]);
        if(value["verifier"] is {} verifier&&!same)
        {
            if(verifier is not JsonObject v)throw Invalid();
            XlsxPassword.Validate(v);
        }
        var p=book.GetFirstChild<WorkbookProtection>()??new WorkbookProtection();
        if(value["structure"] is {} locked)p.LockStructure=locked.GetValue<bool>();
        if(value.ContainsKey("verifier")&&!same)
        {
            foreach(var attribute in Attributes.Values)p.RemoveAttribute(attribute,"");
            if(value["verifier"] is JsonObject v)
                foreach(var (key,node) in v)p.SetAttribute(new OpenXmlAttribute(Attributes[key],"",node!.ToString()));
        }
        // Imported revision/window restrictions and unknown extensions are preserved.
        if(p.Parent is null)book.AddChild(p,true);
    }
    static WriterException Invalid()=>new(ErrorCode.Validation,"Invalid workbook protection","Use structure:boolean and an optional password verifier.");
}
