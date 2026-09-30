using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using System.Text.Json.Nodes;
using Writer.Core;
using Writer.Formats.Pptx;

namespace Writer.Tests;
public class PptxFreeformTests
{
    [Fact] public void Cubic_and_quadratic_curves_are_native_and_invalid_controls_are_atomic()
    {
        using var doc=new PptxAdapter().Create();var slide=Mutations.Add(doc.Root,"slide",new Dictionary<string,string>{{"layout","blank"}},null);
        var shape=Mutations.Add(slide,"shape",new Dictionary<string,string>{{"geometry","rect"}},null);
        shape.SetProp("pathData","{\"w\":100,\"h\":100,\"paths\":[{\"points\":[[0,0],[100,0],[100,100]],\"controls\":{\"1\":[[0,100],[100,100]],\"2\":[[50,50]]},\"closed\":true}]}");
        Assert.Contains("cubicBezTo",shape.GetRaw());Assert.Contains("quadBezTo",shape.GetRaw());
        using var bytes=new MemoryStream();doc.Save(bytes);bytes.Position=0;using var reopened=new PptxAdapter().Open(bytes);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((PptxDocument)reopened).Package).Select(e=>e.Description));
        var vector=PathResolver.Single(reopened.Root,"//shape");var data=JsonNode.Parse(vector.GetProps()["pathData"])!;
        Assert.Equal(1000,data["paths"]![0]!["controls"]!["1"]![0]![1]!.GetValue<double>());
        Assert.Single(data["paths"]![0]!["controls"]!["2"]!.AsArray());
        var before=vector.GetRaw();data["paths"]![0]!["controls"]!["0"]=new JsonArray(new JsonArray(1,2));
        Assert.Throws<ArgumentException>(()=>vector.SetProp("pathData",data.ToJsonString()));Assert.Equal(before,vector.GetRaw());
    }
    [Fact] public void Vector_strokes_and_polygons_are_editable_native_geometry_and_roundtrip()
    {
        using var doc=new PptxAdapter().Create();
        var slide=Mutations.Add(doc.Root,"slide",new Dictionary<string,string>{{"layout","blank"}},null);
        var shape=Mutations.Add(slide,"shape",new Dictionary<string,string>{{"geometry","rect"},{"line","FF3B30"},{"fill","none"}},null);
        var data="{\"w\":200,\"h\":100,\"paths\":[{\"points\":[[0,0],[100,100],[200,0]],\"closed\":false,\"fill\":\"none\",\"stroke\":true},{\"points\":[[20,20],[80,20],[50,60]],\"closed\":true,\"fill\":\"norm\"}]}";
        shape.SetProp("pathData",data);Assert.Equal("custom",shape.GetProps()["geometry"]);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((PptxDocument)doc).Package).Select(e=>e.Description));
        using var stream=new MemoryStream();doc.Save(stream);stream.Position=0;using var reopened=new PptxAdapter().Open(stream);
        var vector=PathResolver.Single(reopened.Root,"//shape");var read=JsonNode.Parse(vector.GetProps()["pathData"])!;
        Assert.Equal(2,read["paths"]!.AsArray().Count);Assert.Equal(500,read["paths"]![0]!["points"]![1]![0]!.GetValue<double>());
        var before=vector.GetRaw();Assert.Throws<ArgumentException>(()=>vector.SetProp("pathData","{\"w\":0,\"h\":100,\"paths\":[]}"));Assert.Equal(before,vector.GetRaw());
        read["paths"]![0]!["points"]![1]![0]=650;vector.SetProp("pathData",read.ToJsonString());Assert.Contains("650",vector.GetProps()["pathData"]);
        vector.SetProp("geometry","ellipse");Assert.False(vector.GetProps().ContainsKey("pathData"));
    }
}
