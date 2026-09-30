using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxBookFeaturesTests
{
    [Fact] public void Manual_page_breaks_round_trip_and_reset_without_removing_print_options()
    {
        using var d = new XlsxAdapter().Create();
        d.Root.Children[0].SetProp("print", "{\"rowBreaks\":[10,20,10],\"colBreaks\":[3],\"header\":\"Report\"}");
        using var ms = new MemoryStream(); d.Save(ms);
        using (var package = SpreadsheetDocument.Open(new MemoryStream(ms.ToArray()), false))
            Assert.Empty(new OpenXmlValidator().Validate(package).Select(e => e.Description));
        using var reopened = new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));
        var sheet = reopened.Root.Children[0]; var p = JsonNode.Parse(sheet.GetProps()["print"])!;
        Assert.Equal("[10,20]", p["rowBreaks"]!.ToJsonString()); Assert.Equal("[3]", p["colBreaks"]!.ToJsonString());
        sheet.SetProp("print", "{\"rowBreaks\":[],\"colBreaks\":[]}");
        p = JsonNode.Parse(sheet.GetProps()["print"])!; Assert.Equal("[]", p["rowBreaks"]!.ToJsonString()); Assert.Equal("Report", p["header"]!.GetValue<string>());
        Assert.Throws<WriterException>(() => sheet.SetProp("print", "{\"rowBreaks\":[0]}"));
    }
    [Fact] public void Names_print_settings_visibility_and_protection_round_trip_in_valid_workbook()
    {
        using var d = new XlsxAdapter().Create();
        var sheet = d.Root.Children[0];
        PathResolver.Single(d.Root, "/sheet[1]/cell[A1]").SetProp("value", "12");
        d.Root.SetProp("names", "{\"Sales_Total\":\"Sheet1!$A$1\"}");
        PathResolver.Single(d.Root, "/sheet[1]/cell[B1]").SetProp("formula", "sales_total*2");
        sheet.SetProp("print", "{\"orientation\":\"landscape\",\"paper\":9,\"fitWidth\":1,\"fitHeight\":0,\"left\":0.3,\"header\":\"Report &P/&N\",\"footer\":\"&A\",\"area\":\"Sheet1!$A$1:$F$50\",\"titles\":\"Sheet1!$1:$2\",\"gridlines\":true}");
        sheet.SetProp("protected", "true");
        var hidden = Mutations.Add(d.Root, "sheet", new Dictionary<string, string>() { ["name"] = "Hidden" }, null);
        hidden.SetProp("visibility", "hidden");
        Assert.Throws<WriterException>(() => sheet.SetProp("visibility", "hidden"));
        using var ms = new MemoryStream(); d.Save(ms);
        using (var package = SpreadsheetDocument.Open(new MemoryStream(ms.ToArray()), false))
            Assert.Empty(new OpenXmlValidator().Validate(package).Select(e => e.Description));
        using var reopened = new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));
        Assert.Equal("24", PathResolver.Single(reopened.Root, "/sheet[1]/cell[B1]").GetProps()["value"]);
        Assert.Contains("Sales_Total", reopened.Root.GetProps()["names"]);
        var props = reopened.Root.Children[0].GetProps(); var p = JsonNode.Parse(props["print"])!;
        Assert.Equal("landscape", p["orientation"]!.GetValue<string>()); Assert.Equal(1, p["fitWidth"]!.GetValue<int>());
        Assert.Equal("Sheet1!$1:$2", p["titles"]!.GetValue<string>()); Assert.Equal("Report &P/&N", p["header"]!.GetValue<string>());
        Assert.Equal("true", props["protected"]); Assert.Equal("hidden", reopened.Root.Children[1].GetProps()["visibility"]);
        reopened.Root.SetProp("names", "{}"); Assert.Equal("Sheet1!$1:$2", JsonNode.Parse(reopened.Root.Children[0].GetProps()["print"])!["titles"]!.GetValue<string>());
        reopened.Root.Children[0].SetProp("protected", "false"); Assert.Equal("false", reopened.Root.Children[0].GetProps()["protected"]);
    }
}
