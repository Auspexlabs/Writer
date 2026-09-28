using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxTablesTests
{
    static void Set(Document d,string a,string key,string v)=>PathResolver.Single(d.Root,"/sheet[1]/cell["+a+"]").SetProp(key,v);
    [Fact] public void Native_tables_keep_parts_calculate_structured_references_and_validate()
    {
        using var d=new XlsxAdapter().Create();var sheet=d.Root.Children[0];
        Set(d,"A1","value","Item");Set(d,"B1","value","Amount");Set(d,"C1","value","Twice");Set(d,"A2","value","one");Set(d,"A3","value","two");Set(d,"B2","value","3");Set(d,"B3","value","7");
        sheet.SetProp("tables","""[{"name":"Sales","range":"A1:C3","style":"TableStyleMedium2"}]""");
        Set(d,"C2","formula","[@Amount]*2");Set(d,"C3","formula","[@[Amount]]*2");Set(d,"E1","formula","SUM(Sales[Amount])");
        using var ms=new MemoryStream();d.Save(ms);
        using(var package=SpreadsheetDocument.Open(new MemoryStream(ms.ToArray()),false))Assert.Empty(new OpenXmlValidator().Validate(package).Select(e=>e.Description));
        using var reopen=new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));
        Assert.Equal("6",PathResolver.Single(reopen.Root,"/sheet[1]/cell[C2]").GetProps()["value"]);Assert.Equal("14",PathResolver.Single(reopen.Root,"/sheet[1]/cell[C3]").GetProps()["value"]);Assert.Equal("10",PathResolver.Single(reopen.Root,"/sheet[1]/cell[E1]").GetProps()["value"]);
        var tables=JsonNode.Parse(reopen.Root.Children[0].GetProps()["tables"])!;Assert.Equal("Amount",tables[0]!["columns"]![1]!["name"]!.GetValue<string>());
        tables[0]!["range"]="A1:C4";Set(reopen,"B4","value","5");reopen.Root.Children[0].SetProp("tables",tables.ToJsonString());
        using var saved=new MemoryStream();reopen.Save(saved);using var again=new XlsxAdapter().Open(new MemoryStream(saved.ToArray()));Assert.Equal("15",PathResolver.Single(again.Root,"/sheet[1]/cell[E1]").GetProps()["value"]);
        Set(again,"B1","value","Value");using var renamed=new MemoryStream();again.Save(renamed);using var named=new XlsxAdapter().Open(new MemoryStream(renamed.ToArray()));Assert.Equal("15",PathResolver.Single(named.Root,"/sheet[1]/cell[E1]").GetProps()["value"]);Assert.Equal("SUM(Sales[Value])",PathResolver.Single(named.Root,"/sheet[1]/cell[E1]").GetProps()["formula"]);
        Assert.Throws<WriterException>(()=>again.Root.Children[0].SetProp("tables","""[{"name":"Sales","range":"A1:C4"},{"name":"Other","range":"B2:D4"}]"""));
        again.Root.Children[0].SetProp("tables","[]");Assert.Equal("[]",again.Root.Children[0].GetProps()["tables"]);Assert.Contains("$B$2:$B$4",PathResolver.Single(again.Root,"/sheet[1]/cell[E1]").GetProps()["formula"]);
    }
}
