using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxOutlineTests
{
    [Fact] public void Nested_row_and_column_groups_round_trip_collapse_and_expand()
    {
        using var d=new XlsxAdapter().Create();var sheet=d.Root.Children[0];
        sheet.SetProp("outline","""[{"axis":"r","start":1,"end":8,"level":1,"collapsed":true},{"axis":"r","start":2,"end":4,"level":2,"collapsed":false},{"axis":"c","start":1,"end":3,"level":1,"collapsed":true}]""");
        using var ms=new MemoryStream();d.Save(ms);
        using(var package=SpreadsheetDocument.Open(new MemoryStream(ms.ToArray()),false))Assert.Empty(new OpenXmlValidator().Validate(package).Select(e=>e.Description));
        using var reopened=new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));var s=reopened.Root.Children[0];var outline=JsonNode.Parse(s.GetProps()["outline"])!.AsArray();
        Assert.Equal(3,outline.Count);Assert.True(outline[0]!["collapsed"]!.GetValue<bool>());Assert.False(outline[1]!["collapsed"]!.GetValue<bool>());Assert.Contains("B",s.GetProps()["hidden"]);
        foreach(var group in outline)group!["collapsed"]=false;s.SetProp("outline",outline.ToJsonString());Assert.False(s.GetProps().ContainsKey("hidden"));
    }
}
