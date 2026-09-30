using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxScenarioTests
{
 [Fact] public void Scenarios_roundtrip_as_native_worksheet_elements_without_changing_current_values()
 {
  using var doc=new XlsxAdapter().Create();var sheet=PathResolver.Single(doc.Root,"/sheet[1]");PathResolver.Single(doc.Root,"/sheet[1]/cell[A1]").SetProp("value","7");
  sheet.SetProp("scenarios","""[{"name":"Best","comment":"Growth case","cells":{"A1":"10","B2":"=literal"}},{"name":"Worst","cells":{"A1":"2"}}]""");
  Assert.Empty(new OpenXmlValidator().Validate(((XlsxDocument)doc).Package));using var stream=new MemoryStream();doc.Save(stream);stream.Position=0;using var reopened=new XlsxAdapter().Open(stream);
  var scenarios=JsonNode.Parse(PathResolver.Single(reopened.Root,"/sheet[1]").GetProps()["scenarios"])!.AsArray();Assert.Equal(2,scenarios.Count);Assert.Equal("=literal",scenarios[0]!["cells"]!["B2"]!.GetValue<string>());Assert.Equal("7",PathResolver.Single(reopened.Root,"/sheet[1]/cell[A1]").GetProps()["value"]);
  Assert.Throws<WriterException>(()=>sheet.SetProp("scenarios","""[{"name":"X","cells":{}}]"""));Assert.Equal(2,JsonNode.Parse(sheet.GetProps()["scenarios"])!.AsArray().Count);
 }
}
