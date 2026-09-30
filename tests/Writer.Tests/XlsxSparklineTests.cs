using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxSparklineTests
{
 [Fact]
 public void Sparkline_groups_keep_native_ranges_colors_and_options_after_save_and_remove()
 {
  using var doc=new XlsxAdapter().Create();var sheet=doc.Root.Children[0];
  const string spec="""[{"type":"column","negative":true,"colorSeries":"#4472C4","colorNegative":"#C00000","markers":true,"empty":"gap","sparklines":[{"source":"Sheet1!A1:D1","cell":"F1"},{"source":"Sheet1!A2:D2","cell":"F2"}]}]""";
  Mutations.Set(sheet,new Dictionary<string,string>{{"sparklines",spec}});
  using var stream=new MemoryStream();doc.Save(stream);stream.Position=0;using var back=new XlsxAdapter().Open(stream);var s=back.Root.Children[0];var groups=JsonNode.Parse(s.GetProps()["sparklines"])!;
  Assert.Equal("Sheet1!A2:D2",groups[0]!["sparklines"]![1]!["source"]!.GetValue<string>());Assert.Equal("#C00000",groups[0]!["colorNegative"]!.GetValue<string>());Assert.True(groups[0]!["negative"]!.GetValue<bool>());
  var errors=string.Join("\n",new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((XlsxDocument)back).Package).Select(e=>e.Path?.XPath+": "+e.Description));Assert.True(errors.Length==0,errors);
  Mutations.Set(s,new Dictionary<string,string>{{"sparklines",groups.ToJsonString()}});Assert.Equal("column",JsonNode.Parse(s.GetProps()["sparklines"])![0]!["type"]!.GetValue<string>());
  Mutations.Set(s,new Dictionary<string,string>{{"sparklines","[]"}});Assert.Equal("[]",s.GetProps()["sparklines"]);
 }
}
