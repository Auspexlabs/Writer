using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
namespace Writer.Formats.Xlsx;
static class XlsxScenarios
{
    internal static string Read(XlsxSheet sheet) => new JsonArray((sheet.Ws.GetFirstChild<Scenarios>()?.Elements<Scenario>()??[]).Select(s=>(JsonNode)new JsonObject {
        ["name"]=s.Name?.Value??"",["comment"]=s.Comment?.Value??"",["user"]=s.User?.Value??"",["locked"]=s.Locked?.Value??false,["hidden"]=s.Hidden?.Value??false,
        ["cells"]=new JsonObject(s.Elements<InputCells>().Where(c=>c.Deleted?.Value!=true&&c.CellReference?.Value is not null).Select(c=>new KeyValuePair<string,JsonNode?>(c.CellReference!.Value!,JsonValue.Create(c.Val?.Value??""))))
    }).ToArray()).ToJsonString();
    internal static void Write(XlsxSheet sheet,string json)
    {
        var list=JsonNode.Parse(json) as JsonArray??throw Bad("Scenarios must be an array");var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);var container=new Scenarios();
        foreach(var item in list){var name=item?["name"]?.GetValue<string>();if(string.IsNullOrWhiteSpace(name)||name.Length>255||!names.Add(name)||item?["cells"] is not JsonObject cells||cells.Count is <1 or >32)throw Bad("A scenario needs a unique name and 1–32 input cells");
            var scenario=new Scenario{Name=name,Count=(uint)cells.Count,Locked=item["locked"]?.GetValue<bool>()??false,Hidden=item["hidden"]?.GetValue<bool>()??false};if(item["comment"] is { } comment)scenario.Comment=comment.GetValue<string>();if(item["user"] is { } user)scenario.User=user.GetValue<string>();
            foreach(var (reference,value) in cells){var (c,r)=XlsxCells.Parse(reference);if(c>16384||r>1048576)throw Bad("Scenario input outside worksheet");scenario.Append(new InputCells{CellReference=XlsxCells.Reference(c,r),Val=value?.ToString()??""});}container.Append(scenario);
        }
        sheet.Ws.RemoveAllChildren<Scenarios>();if(list.Count>0)sheet.Ws.AddChild(container,true);
    }
    static WriterException Bad(string message)=>new(ErrorCode.Validation,message,"Use [{name,cells:{A1:10,B2:20},comment?}].");
}
