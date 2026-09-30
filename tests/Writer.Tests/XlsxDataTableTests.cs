using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxDataTableTests
{
 [Theory][InlineData(false)][InlineData(true)] public void Native_data_table_formulas_recalculate_save_and_reopen(bool two)
 {
  using var doc=new XlsxAdapter().Create();void Set(string a,string key,string value)=>PathResolver.Single(doc.Root,"/sheet[1]/cell["+a+"]").SetProp(key,value);
  Set("F1","value","10");Set("F2","value","7");Set("A2","value","3");Set("A3","value","5");Set("H1","formula","SUM(B2:C3)");
  if(two){Set("A1","formula","F1*F2");Set("B1","value","2");Set("C1","value","4");}else{Set("B1","formula","F1*2");Set("C1","formula","F1^2");}
  var table=new JsonObject{["range"]="A1:C3",["colInput"]="F1"};if(two){table["rowInput"]="F1";table["colInput"]="F2";}
  var sheet=PathResolver.Single(doc.Root,"/sheet[1]");sheet.SetProp("dataTables",new JsonArray((JsonNode)table).ToJsonString());Assert.Empty(new OpenXmlValidator().Validate(((XlsxDocument)doc).Package));
  using var stream=new MemoryStream();doc.Save(stream);stream.Position=0;using var opened=new XlsxAdapter().Open(stream);string V(string a)=>PathResolver.Single(opened.Root,"/sheet[1]/cell["+a+"]").GetProps()["value"];
  Assert.Equal("6",V("B2"));Assert.Equal(two?"20":"25",V("C3"));Assert.Equal(two?"48":"50",V("H1"));Assert.Equal("10",V("F1"));Assert.Equal("A1:C3",JsonNode.Parse(PathResolver.Single(opened.Root,"/sheet[1]").GetProps()["dataTables"])![0]!["range"]!.GetValue<string>());
  PathResolver.Single(opened.Root,"/sheet[1]/cell[A3]").SetProp("value","6");using var again=new MemoryStream();opened.Save(again);again.Position=0;using var last=new XlsxAdapter().Open(again);Assert.Equal(two?"54":"63",PathResolver.Single(last.Root,"/sheet[1]/cell[H1]").GetProps()["value"]);
 }
}
