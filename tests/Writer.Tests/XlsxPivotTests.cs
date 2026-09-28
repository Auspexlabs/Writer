using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxPivotTests
{
    static string Value(Document d,string sheet,string a)=>PathResolver.Single(d.Root,$"/sheet[{sheet}]/cell[{a}]").GetProps()["value"];
    [Fact] public void Native_pivot_caches_aggregate_refresh_and_survive_reopen()
    {
        using var d=new XlsxAdapter().Create();
        Mutations.Set(PathResolver.Single(d.Root,"/sheet[1]/range[A1:C4]"),new Dictionary<string,string>{{"values","""[["Region","Channel","Value"],["East","Web",3],["East","Store",7],["West","Web",5]]"""}});
        var target=Mutations.Add(d.Root,"sheet",new Dictionary<string,string>{{"name","Summary"}},null);
        target.SetProp("pivots","""[{"name":"SalesPivot","sourceSheet":"Sheet1","sourceRange":"A1:C4","target":"A1","rows":[0],"cols":[1],"values":[{"field":2,"fn":"sum","name":"Sales"}]}]""");
        Assert.Equal("10",Value(d,"2","D2"));Assert.Equal("15",Value(d,"2","D4"));Assert.Equal("8",Value(d,"2","B4"));
        using var ms=new MemoryStream();d.Save(ms);
        using(var package=SpreadsheetDocument.Open(new MemoryStream(ms.ToArray()),false)){
            Assert.Empty(new OpenXmlValidator().Validate(package).Select(e=>$"{e.Path?.XPath}: {e.Description}"));
            var cache=package.WorkbookPart!.PivotTableCacheDefinitionParts.Single();Assert.Equal(3U,cache.PivotCacheDefinition!.RecordCount!.Value);Assert.Equal(3U,cache.PivotTableCacheRecordsPart!.PivotCacheRecords!.Count!.Value);
        }
        using var again=new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));var pivot=JsonNode.Parse(again.Root.Children[1].GetProps()["pivots"])!.AsArray();Assert.Equal("SalesPivot",pivot[0]!["name"]!.GetValue<string>());
        PathResolver.Single(again.Root,"/sheet[1]/cell[C2]").SetProp("value","30");pivot[0]!["refresh"]=true;again.Root.Children[1].SetProp("pivots",pivot.ToJsonString());Assert.Equal("42",Value(again,"2","D4"));
        again.Root.Children[1].SetProp("pivots","[]");Assert.Equal("",Value(again,"2","D4"));
    }
    [Fact] public void Pivot_rejects_overwriting_source_or_other_cells()
    {
        using var d=new XlsxAdapter().Create();Mutations.Set(PathResolver.Single(d.Root,"/sheet[1]/range[A1:B3]"),new Dictionary<string,string>{{"values","""[["Group","Value"],["A",2],["A",3]]"""}});
        Assert.Throws<WriterException>(()=>d.Root.Children[0].SetProp("pivots","""[{"name":"Pivot","sourceSheet":"Sheet1","sourceRange":"A1:B3","target":"A1","rows":[0],"values":[{"field":1}]}]"""));Assert.Equal("Group",Value(d,"1","A1"));
    }
}
