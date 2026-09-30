using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using Writer.Formats.Pptx;

namespace Writer.Tests;
public class OfficeSmartArtTests
{
    [Theory]
    [InlineData("docx","list")]
    [InlineData("docx","hierarchy")]
    [InlineData("pptx","process")]
    [InlineData("pptx","cycle")]
    public void SmartArt_nodes_hierarchy_and_cached_text_survive_edit_save_and_reopen(string format,string layout)
    {
        using var doc=format=="docx"?new DocxAdapter().Create():new PptxAdapter().Create();
        var target=format=="docx"?doc.Root:Mutations.Add(doc.Root,"slide",new Dictionary<string,string>{{"layout","blank"}},null);
        var parent="{11111111-1111-1111-1111-111111111111}";var child="{22222222-2222-2222-2222-222222222222}";
        var spec=new JsonObject{["layout"]=layout,["nodes"]=new JsonArray(new JsonObject{["id"]=parent,["text"]="Plan"},new JsonObject{["id"]=child,["parent"]=parent,["text"]="Build"})};
        Mutations.Set(target,new Dictionary<string,string>{{"smartArtData",spec.ToJsonString()}});
        var obj=PathResolver.Single(doc.Root,"//object");var model=JsonNode.Parse(obj.GetProps()["smartart"])!;Assert.Equal(layout,model["layout"]!.GetValue<string>());Assert.Equal(parent,model["nodes"]![1]!["parent"]!.GetValue<string>());Assert.Equal(layout=="list"?2:layout=="cycle"?4:3,model["shapes"]!.AsArray().Count);if(layout!="list")Assert.Contains(model["shapes"]!.AsArray(),s=>s!["geom"]!.GetValue<string>()=="line");
        var package=format=="docx"?((DocxDocument)doc).Package:(DocumentFormat.OpenXml.Packaging.OpenXmlPackage)((PptxDocument)doc).Package;
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(package).Select(e=>$"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}"));
        spec["path"]=obj.Path;spec["layout"]="preserve";spec["nodes"]![1]!["text"]="Ship";
        Mutations.Set(target,new Dictionary<string,string>{{"smartArtData",spec.ToJsonString()}});
        using var stream=new MemoryStream();doc.Save(stream);stream.Position=0;using var back=format=="docx"?new DocxAdapter().Open(stream):new PptxAdapter().Open(stream);
        var reopened=JsonNode.Parse(PathResolver.Single(back.Root,"//object").GetProps()["smartart"])!;Assert.Equal("Ship",reopened["nodes"]![1]!["text"]!.GetValue<string>());Assert.Equal("Ship",reopened["shapes"]!.AsArray().Last()!["text"]!.GetValue<string>());
    }
}
