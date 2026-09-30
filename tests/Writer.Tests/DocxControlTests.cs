using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
namespace Writer.Tests;
public class DocxControlTests
{
    [Fact] public void Bound_controls_update_custom_xml_and_all_mirrors_without_changing_other_data()
    {
        using var doc=(DocxDocument)new DocxAdapter().Create();
        var part=doc.Main.AddCustomXmlPart(DocumentFormat.OpenXml.Packaging.CustomXmlPartType.CustomXml);
        using(var bytes=new MemoryStream(System.Text.Encoding.UTF8.GetBytes("<form xmlns='urn:writer-form'><name>Before</name><other keep='yes'>Untouched</other></form>")))part.FeedData(bytes);
        const string store="{397AD51F-C7C7-4D58-8283-BAB11518916A}";
        part.AddNewPart<DocumentFormat.OpenXml.Packaging.CustomXmlPropertiesPart>().DataStoreItem=new DocumentFormat.OpenXml.CustomXmlDataProperties.DataStoreItem{ItemId=store};
        for(var i=1;i<=2;i++)doc.Main.Document!.Body!.PrependChild(new W.Paragraph(new W.SdtRun(new W.SdtProperties(new W.SdtId{Val=i},new W.DataBinding{StoreItemId=store,XPath="/f:form/f:name",PrefixMappings="xmlns:f='urn:writer-form'"},new W.SdtContentText()),new W.SdtContentRun(new W.Run(new W.Text("Before"))))));
        var controls=PathResolver.Query(doc.Root,"//control").ToArray();var data=JsonNode.Parse(controls[0].GetProps()["data"])!;Assert.True(data["bound"]!.GetValue<bool>());data["value"]="王 & Li";controls[0].SetProp("data",data.ToJsonString());
        Assert.All(controls,c=>Assert.Equal("王 & Li",c.GetProps()["text"]));
        using(var reader=new StreamReader(part.GetStream())){var xml=reader.ReadToEnd();Assert.Contains("王 &amp; Li",xml);Assert.Contains("Untouched",xml);Assert.Contains("keep=\"yes\"",xml);}
        using var saved=new MemoryStream();doc.Save(saved);using var reopened=new DocxAdapter().Open(new MemoryStream(saved.ToArray()));Assert.All(PathResolver.Query(reopened.Root,"//control"),c=>Assert.Equal("王 & Li",c.GetProps()["text"]));
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(doc.Package).Select(e=>e.Description));
        var before=controls[0].GetRaw();data["properties"]=data["properties"]!.GetValue<string>().Replace("/f:form/f:name","/f:form/f:missing");Assert.Throws<WriterException>(()=>controls[0].SetProp("data",data.ToJsonString()));Assert.Equal(before,controls[0].GetRaw());
    }
    static Dictionary<string,string> P(params (string,string)[] p)=>p.ToDictionary(x=>x.Item1,x=>x.Item2);
    [Theory]
    [InlineData("text","Name")]
    [InlineData("checkbox","☒")]
    [InlineData("date","2026-09-28")]
    [InlineData("dropdown","Name")]
    [InlineData("combo","Name")]
    public void Controls_are_native_editable_and_stable_beside_changed_text(string type,string text)
    {
        using var doc=new DocxAdapter().Create();var para=Mutations.Add(doc.Root.Children.Single(),"paragraph",P(("text","Left Right")),null);
        var spec=new JsonObject{["type"]=type,["title"]="Entry",["tag"]="form_field",["value"]="Name",["checked"]=true,["date"]="2026-09-28",["items"]=new JsonArray(new JsonObject{["label"]="Name",["value"]="name"})};
        var control=Mutations.Add(para,"control",P(("data",spec.ToJsonString()),("at","5")),null);
        Assert.True(para.Children.Any(c=>c.Kind=="control"),para.GetRaw());Assert.Equal(text,control.GetProps()["text"]);Assert.Equal("Left Right",para.GetProps()["text"]);
        Mutations.Set(para,P(("html","<b>Before Left Right After</b>")));Assert.Equal(text,para.Children.Single(c=>c.Kind=="control").GetProps()["text"]);
        using var ms=new MemoryStream();doc.Save(ms);using var reopened=new DocxAdapter().Open(new MemoryStream(ms.ToArray()));var again=PathResolver.Single(reopened.Root,"//control");
        Assert.Equal(type,JsonNode.Parse(again.GetProps()["data"])!["type"]!.GetValue<string>());Assert.Equal(text,again.GetProps()["text"]);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((DocxDocument)reopened).Package).Select(e=>e.Path?.XPath+": "+e.Description));
        var data=JsonNode.Parse(again.GetProps()["data"])!;data["type"]="text";data["value"]="Changed";again.SetProp("data",data.ToJsonString());Assert.Equal("Changed",again.GetProps()["text"]);
        again.Remove();Assert.Empty(((DocxDocument)reopened).Main.Document!.Descendants<W.SdtRun>());
    }
    [Fact] public void Imported_missing_and_duplicate_control_ids_are_stable_across_reads()
    {
        using var doc=(DocxDocument)new DocxAdapter().Create();var body=doc.Main.Document!.Body!;
        body.PrependChild(new W.Paragraph(new W.SdtRun(new W.SdtProperties(new W.SdtContentText()),new W.SdtContentRun(new W.Run(new W.Text("First")))),new W.SdtRun(new W.SdtProperties(new W.SdtContentText()),new W.SdtContentRun(new W.Run(new W.Text("Second"))))));
        using var stream=new MemoryStream();doc.Save(stream);var bytes=stream.ToArray();
        using var a=new DocxAdapter().Open(new MemoryStream(bytes));using var b=new DocxAdapter().Open(new MemoryStream(bytes));
        var left=PathResolver.Query(a.Root,"//control").Select(x=>x.GetProps()["id"]).ToArray();var right=PathResolver.Query(b.Root,"//control").Select(x=>x.GetProps()["id"]).ToArray();Assert.Equal(left,right);Assert.Equal(2,left.Distinct().Count());
    }
    [Fact] public void Invalid_list_edits_are_atomic_and_empty_controls_keep_their_placeholder()
    {
        using var doc=new DocxAdapter().Create();var para=Mutations.Add(doc.Root.Children.Single(),"paragraph",P(("text","")),null);
        var control=Mutations.Add(para,"control",P(("data","{\"type\":\"text\",\"title\":\"Fill here\",\"value\":\"\"}")),null);var before=control.GetRaw();Assert.Equal("Fill here",control.GetProps()["text"]);
        Assert.Throws<WriterException>(()=>control.SetProp("data","{\"type\":\"dropdown\",\"value\":\"wrong\",\"items\":[]}"));Assert.Equal(before,control.GetRaw());
    }
}
