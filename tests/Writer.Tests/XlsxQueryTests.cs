using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Formats.Xlsx;
using Writer.Core;
namespace Writer.Tests;
public class XlsxQueryTests
{
    [Fact] public void Query_steps_survive_workbook_save_without_touching_unrelated_custom_xml()
    {
        using var d=(XlsxDocument)new XlsxAdapter().Create();
        var unrelated=d.Workbook.AddCustomXmlPart(CustomXmlPartType.CustomXml);using(var data=new MemoryStream(System.Text.Encoding.UTF8.GetBytes("<external xmlns='urn:other'>keep</external>")))unrelated.FeedData(data);
        const string query="""[{"name":"Clean","source":{"sheet":"Sheet1","range":"A1:C20"},"steps":[{"type":"filter","column":"City","op":"eq","value":"北京"}],"targetSheet":"Result","target":"A1"}]""";
        Mutations.Set(d.Root,new Dictionary<string,string>{{"queries",query}});using var stream=new MemoryStream();d.Save(stream);
        using var back=(XlsxDocument)new XlsxAdapter().Open(new MemoryStream(stream.ToArray()));Assert.Equal(System.Text.Json.Nodes.JsonNode.Parse(query)!.ToJsonString(),back.Root.GetProps()["queries"]);
        Assert.Empty(new OpenXmlValidator().Validate(back.Package).Select(e=>e.Description));Assert.Equal(2,back.Workbook.CustomXmlParts.Count());
        Mutations.Set(back.Root,new Dictionary<string,string>{{"queries","[]"}});Assert.Single(back.Workbook.CustomXmlParts);
    }
    [Fact] public void Solver_setup_is_per_sheet_survives_rename_and_does_not_touch_queries()
    {
        using var d=(XlsxDocument)new XlsxAdapter().Create();
        var first=d.Root.Children.First(); var second=Mutations.Add(d.Root,"sheet",new Dictionary<string,string>{{"name","Other"}},null);
        const string setup="""{"target":"C1","variables":"A1:B1","mode":"max","method":"linear","constraints":[{"left":"A1+B1","op":"<=","right":"10"}]}""";
        Mutations.Set(first,new Dictionary<string,string>{{"solver",setup}});
        Mutations.Set(first,new Dictionary<string,string>{{"name","Renamed"}});
        Mutations.Set(second,new Dictionary<string,string>{{"solver",setup.Replace("C1","D1")}});
        using var stream=new MemoryStream();d.Save(stream);using var back=(XlsxDocument)new XlsxAdapter().Open(new MemoryStream(stream.ToArray()));
        Assert.Contains("C1",back.Root.Children[0].GetProps()["solver"]);Assert.Contains("D1",back.Root.Children[1].GetProps()["solver"]);
        Assert.Empty(new OpenXmlValidator().Validate(back.Package).Select(e=>e.Description));
        Mutations.Set(back.Root.Children[0],new Dictionary<string,string>{{"solver","null"}});
        Assert.Equal("null",back.Root.Children[0].GetProps()["solver"]);Assert.Contains("D1",back.Root.Children[1].GetProps()["solver"]);
    }
}
