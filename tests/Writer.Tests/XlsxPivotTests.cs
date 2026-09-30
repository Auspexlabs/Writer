using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxPivotTests
{
    [Fact] public void Numeric_group_members_use_value_equality_for_imported_numeric_text()
    {
        using var d=new XlsxAdapter().Create();
        PathResolver.Single(d.Root,"/sheet[1]/range[A1:B4]").SetProp("values","""[["Bucket","Value"],[1,3],[1,5],[2,7]]""");
        var native=(XlsxDocument)d;
        native.Workbook.WorksheetParts.Single().Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single(c=>c.CellReference!.Value=="A2").CellValue!.Text="1.00";
        var target=Mutations.Add(d.Root,"sheet",new Dictionary<string,string>{{"name","Summary"}},null);
        target.SetProp("pivots","""[{"name":"Grouped","sourceSheet":"Sheet1","sourceRange":"A1:B4","rows":[0],"values":[{"field":1}],"groups":[{"field":0,"items":[{"name":"One","items":[1]}]}]}]""");
        Assert.Equal("One",Value(d,"2","A2"));Assert.Equal("8",Value(d,"2","B2"));Assert.Equal("2",Value(d,"2","A3"));Assert.Equal("15",Value(d,"2","B4"));
        using var ms=new MemoryStream();d.Save(ms);using var reopened=new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));var loaded=JsonNode.Parse(reopened.Root.Children[1].GetProps()["pivots"])!;
        Assert.Single(loaded[0]!["groups"]![0]!["items"]!.AsArray());Assert.Equal(1,loaded[0]!["groups"]![0]!["items"]![0]!["items"]![0]!.GetValue<int>());
        loaded[0]!["refresh"]=1;reopened.Root.Children[1].SetProp("pivots",loaded.ToJsonString());Assert.Equal("8",Value(reopened,"2","B2"));Assert.Equal("7",Value(reopened,"2","B3"));
    }
    [Fact] public void Grouped_items_and_calculated_fields_are_native_and_use_aggregate_operands()
    {
        using var d=new XlsxAdapter().Create();
        PathResolver.Single(d.Root,"/sheet[1]/range[A1:C4]").SetProp("values","""[["Region","Sales","Unit Cost"],["A",100,50],["A",200,80],["B",100,90]]""");
        var target=Mutations.Add(d.Root,"sheet",new Dictionary<string,string>{{"name","Summary"}},null);
        var spec=JsonNode.Parse("""[{"name":"Margin","sourceSheet":"Sheet1","sourceRange":"A1:C4","rows":[0],"values":[{"name":"Profit","formula":"Sales-'Unit Cost'","fn":"sum"},{"name":"Margin ratio","formula":"(Sales-'Unit Cost')/Sales","fn":"sum"}],"groups":[{"field":0,"items":[{"name":"Combined","items":["A","B"]}]}]}]""")!;
        target.SetProp("pivots",spec.ToJsonString());Assert.Equal("Combined",Value(d,"2","A2"));Assert.Equal("180",Value(d,"2","B2"));Assert.Equal("0.45",Value(d,"2","C2"));Assert.Equal("0.45",Value(d,"2","C3"));
        using var ms=new MemoryStream();d.Save(ms);
        using(var package=SpreadsheetDocument.Open(new MemoryStream(ms.ToArray()),false)){
            Assert.Empty(new OpenXmlValidator().Validate(package).Select(e=>$"{e.Path?.XPath}: {e.Description}"));
            var fields=package.WorkbookPart!.PivotTableCacheDefinitionParts.Single().PivotCacheDefinition!.CacheFields!;
            Assert.Contains(fields.ChildElements,f=>f.GetAttributes().Any(a=>a.LocalName=="formula"&&a.Value=="Sales-'Unit Cost'"));
            Assert.Contains("discretePr",fields.OuterXml);Assert.Contains("databaseField=\"0\"",fields.OuterXml);
        }
        using var reopened=new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));
        var loaded=JsonNode.Parse(reopened.Root.Children[1].GetProps()["pivots"])!;
        Assert.Equal("Sales-'Unit Cost'",loaded[0]!["values"]![0]!["formula"]!.GetValue<string>());
        Assert.Equal(0,loaded[0]!["groups"]![0]!["field"]!.GetValue<int>());Assert.Equal("Combined",loaded[0]!["groups"]![0]!["items"]![0]!["name"]!.GetValue<string>());
        PathResolver.Single(reopened.Root,"/sheet[1]/cell[B2]").SetProp("value","200");loaded[0]!["refresh"]=1;
        reopened.Root.Children[1].SetProp("pivots",loaded.ToJsonString());Assert.Equal("280",Value(reopened,"2","B2"));Assert.Equal("0.56",Value(reopened,"2","C3"));
        loaded[0]!.AsObject().Remove("groups");loaded[0]!["refresh"]=2;reopened.Root.Children[1].SetProp("pivots",loaded.ToJsonString());Assert.Equal("270",Value(reopened,"2","B2"));Assert.Equal("0.675",Value(reopened,"2","C2"));Assert.Equal("0.56",Value(reopened,"2","C4"));
        loaded[0]!["values"]![0]!["formula"]="Missing+1";Assert.Throws<WriterException>(()=>reopened.Root.Children[1].SetProp("pivots",loaded.ToJsonString()));Assert.Equal("270",Value(reopened,"2","B2"));
        Assert.Equal("Sales-'Unit Cost'",JsonNode.Parse(reopened.Root.Children[1].GetProps()["pivots"])![0]!["values"]![0]!["formula"]!.GetValue<string>());
    }
    [Fact] public void Multi_axis_multi_value_pivots_keep_layout_and_refresh_after_reopen()
    {
        using var d=new XlsxAdapter().Create();
        PathResolver.Single(d.Root,"/sheet[1]/range[A1:F4]").SetProp("values","""[["Region","City","Channel","Year","Sales","Qty"],["East","A","Web",2025,3,1],["East","A","Store",2025,7,2],["West","B","Web",2026,5,3]]""");
        var target=Mutations.Add(d.Root,"sheet",new Dictionary<string,string>{{"name","Summary"}},null);
        target.SetProp("pivots","""[{"name":"Multi","sourceSheet":"Sheet1","sourceRange":"A1:F4","rows":[0,1],"cols":[2,3],"values":[{"field":4,"fn":"sum"},{"field":4,"fn":"average"},{"field":5,"fn":"sum"}]}]""");
        Assert.Equal("10",Value(d,"2","L4"));Assert.Equal("5",Value(d,"2","M4"));Assert.Equal("15",Value(d,"2","L6"));Assert.Equal("6",Value(d,"2","N6"));
        using var ms=new MemoryStream();d.Save(ms);
        using(var package=SpreadsheetDocument.Open(new MemoryStream(ms.ToArray()),false))Assert.Empty(new OpenXmlValidator().Validate(package).Select(e=>e.Description));
        using var reopened=new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));var pivots=JsonNode.Parse(reopened.Root.Children[1].GetProps()["pivots"])!.AsArray();
        Assert.Equal("[0,1]",pivots[0]!["rows"]!.ToJsonString());Assert.Equal("[2,3]",pivots[0]!["cols"]!.ToJsonString());Assert.Equal(3,pivots[0]!["values"]!.AsArray().Count);
        PathResolver.Single(reopened.Root,"/sheet[1]/cell[E2]").SetProp("value","30");pivots[0]!["refresh"]=1;reopened.Root.Children[1].SetProp("pivots",pivots.ToJsonString());Assert.Equal("42",Value(reopened,"2","L6"));
    }
    [Fact] public void Display_calculations_and_axis_filters_are_native_and_recalculate_after_reopen()
    {
        using var doc=new XlsxAdapter().Create();
        PathResolver.Single(doc.Root,"/sheet[1]/range[A1:C4]").SetProp("values","""[["Region","Channel","Value"],["East","Web",3],["East","Store",7],["West","Web",5]]""");
        var target=Mutations.Add(doc.Root,"sheet",new Dictionary<string,string>{{"name","Summary"}},null);
        var spec=JsonNode.Parse("""[{"name":"Sales","sourceSheet":"Sheet1","sourceRange":"A1:C4","rows":[0],"cols":[1],"values":[{"field":2,"fn":"sum","showAs":"percentOfRow"}]}]""")!;
        target.SetProp("pivots",spec.ToJsonString());Assert.Equal("0.3",Value(doc,"2","B2"));Assert.Equal("1",Value(doc,"2","D2"));
        spec[0]!["filters"]=JsonNode.Parse("""[{"field":0,"items":["East"]}]""");target.SetProp("pivots",spec.ToJsonString());Assert.Equal("0.3",Value(doc,"2","B3"));
        using var stream=new MemoryStream();doc.Save(stream);using var reopened=new XlsxAdapter().Open(new MemoryStream(stream.ToArray()));
        var loaded=JsonNode.Parse(reopened.Root.Children[1].GetProps()["pivots"])!;Assert.Equal("percentOfRow",loaded[0]!["values"]![0]!["showAs"]!.GetValue<string>());Assert.Equal("East",loaded[0]!["filters"]![0]!["items"]![0]!.GetValue<string>());
        using var package=SpreadsheetDocument.Open(new MemoryStream(stream.ToArray()),false);Assert.Empty(new OpenXmlValidator().Validate(package).Select(e=>e.Description));
        loaded[0]!.AsObject().Remove("filters");loaded[0]!["values"]![0]!["showAs"]="runTotal";loaded[0]!["values"]![0]!["baseField"]=0;
        reopened.Root.Children[1].SetProp("pivots",loaded.ToJsonString());Assert.Equal("8",Value(reopened,"2","B3"));Assert.Equal("15",Value(reopened,"2","D3"));
        loaded[0]!["values"]![0]!["showAs"]="difference";reopened.Root.Children[1].SetProp("pivots",loaded.ToJsonString());Assert.Equal("2",Value(reopened,"2","B3"));Assert.Equal("-5",Value(reopened,"2","D3"));
    }
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
