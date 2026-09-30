using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using Writer.Formats.Pptx;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Writer.Tests;

public class OfficeChartDataTests
{
    static Dictionary<string,string> P(string json) => new() { ["chartData"] = json };
    const string Data = """{"type":"combo","title":"Revenue","labels":["North","South"],"series":[{"name":"Actual","values":[12,18],"kind":"column"},{"name":"Target","values":[10,20],"kind":"line"}],"xTitle":"Region","yTitle":"USD","legend":"bottom","dataLabels":true}""";
    [Theory]
    [InlineData("docx")]
    [InlineData("pptx")]
    public void Multi_series_chart_edits_keep_frame_and_embedded_workbook_after_save(string format)
    {
        using var doc = format == "docx" ? new DocxAdapter().Create() : new PptxAdapter().Create();
        var target = format == "docx" ? doc.Root : Mutations.Add(doc.Root,"slide",new Dictionary<string,string>{{"layout","blank"}},null);
        Mutations.Set(target,P(Data));
        var container = format == "docx" ? doc.Root.Children.Single() : target;
        var chart=container.Children.Single(n=>n.Kind=="object");var raw=chart.GetRaw();
        using var model=JsonDocument.Parse(chart.GetProps()["chart"]);Assert.Equal(2,model.RootElement.GetProperty("series").GetArrayLength());
        var edited=Data.Replace("12,18","90,80").Replace("Revenue","Updated");
        edited=edited[..^1]+",\"path\":\""+(format=="docx"?"/body/object[1]":"/slide[1]/object[1]")+"\"}";
        Mutations.Set(target,P(edited));Assert.Equal(raw,chart.GetRaw());
        using var stream=new MemoryStream();doc.Save(stream);stream.Position=0;
        using OpenXmlPackage package = format=="docx"?WordprocessingDocument.Open(stream,false):PresentationDocument.Open(stream,false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(package).Select(e=>e.Description));
        var part=format=="docx"?((WordprocessingDocument)package).MainDocumentPart!.ChartParts.Single():((PresentationDocument)package).PresentationPart!.SlideParts.Single().ChartParts.Single();
        Assert.Contains("90",part.ChartSpace!.OuterXml);Assert.Contains("Updated",part.ChartSpace!.OuterXml);
        Assert.Equal(2,part.ChartSpace!.Descendants<C.SeriesText>().Count());
        using var embedded=part.EmbeddedPackagePart!.GetStream();using var wb=SpreadsheetDocument.Open(embedded,false);
        Assert.Contains("90",wb.WorkbookPart!.WorksheetParts.Single().Worksheet!.OuterXml);
    }
    [Fact]
    public void Scatter_charts_keep_distinct_x_values_and_reject_invalid_series()
    {
        using var doc=new DocxAdapter().Create();
        Mutations.Set(doc.Root,P("""{"type":"scatter","labels":["A","B"],"series":[{"name":"One","x":[1,2],"values":[3,4]},{"name":"Two","x":[10,20],"values":[30,null]}]}"""));
        using var model=JsonDocument.Parse(doc.Root.Children.Single().Children.Single(n=>n.Kind=="object").GetProps()["chart"]);
        Assert.Equal(20,model.RootElement.GetProperty("series")[1].GetProperty("x")[1].GetDouble());
        Assert.Throws<WriterException>(()=>Mutations.Set(doc.Root,P("""{"labels":["A","B"],"series":[{"values":[1]}]}""")));
    }
}
