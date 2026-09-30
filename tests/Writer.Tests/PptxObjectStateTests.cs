using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Pptx;

namespace Writer.Tests;

public class PptxObjectStateTests
{
    static Dictionary<string,string> P(params (string,string)[] values) => values.ToDictionary(x=>x.Item1,x=>x.Item2);
    [Fact]
    public void Selection_pane_names_visibility_and_locks_survive_reopen_without_removing_aspect_locks()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root,"slide",P(("layout","blank")),null);
        var shape = Mutations.Add(slide,"shape",P(("text","Title"),("lockAspect","true")),null);
        Mutations.Add(slide,"shape",P(("text","Child")),null);
        Mutations.Add(slide,"connector",P(("geometry","straightConnector1")),null);
        Mutations.Add(slide,"table",P(("data","[[1,2],[3,4]]")),null);
        Mutations.Add(slide,"group",P(("members","shape[2]")),null);
        var nodes=slide.Children.SelectMany(n=>new[]{n}.Concat(n.Kind=="group"?n.Children:[])).ToList();
        foreach(var node in nodes) Mutations.Set(node,P(("name","Named "+node.Kind),("hidden","true"),("locked","true")));
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2016).Validate(((PptxDocument)doc).Package).Select(e=>e.Description));
        using var stream=new MemoryStream();doc.Save(stream);stream.Position=0;using var back=new PptxAdapter().Open(stream);
        var reopened=back.Root.Children.Single().Children.SelectMany(n=>new[]{n}.Concat(n.Kind=="group"?n.Children:[])).ToList();
        Assert.Equal(nodes.Count,reopened.Count);
        foreach(var node in reopened)
        {
            var props=node.GetProps();Assert.Equal("Named "+node.Kind,props["name"]);Assert.Equal("true",props["hidden"]);Assert.Equal("true",props["locked"]);
            Mutations.Set(node,P(("hidden","false"),("locked","false")));Assert.False(node.GetProps().ContainsKey("locked"));Assert.False(node.GetProps().ContainsKey("hidden"));
        }
        Assert.Equal("true",reopened.First(n=>n.Kind=="shape").GetProps()["lockAspect"]);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2016).Validate(((PptxDocument)back).Package).Select(e=>e.Description));
    }
}
