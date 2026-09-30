using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Xlsx;

namespace Writer.Tests;
public class XlsxAdvancedFilterTests
{
    [Fact] public void Criteria_and_hidden_rows_survive_rename_save_clear_and_invalid_edits()
    {
        using var doc=new XlsxAdapter().Create();var sheet=doc.Root.Children[0];
        const string spec="{\"range\":\"A1:C5\",\"criteria\":\"E1:F3\",\"unique\":true,\"hidden\":[2,4]}";
        sheet.SetProp("advancedFilter",spec);sheet.SetProp("hidden","{\"rows\":[3,5,10]}");sheet.SetProp("name","Renamed");
        using var stream=new MemoryStream();doc.Save(stream);using var opened=new XlsxAdapter().Open(new MemoryStream(stream.ToArray()));var again=opened.Root.Children[0];
        Assert.Equal(spec,again.GetProps()["advancedFilter"]);Assert.Contains("10",again.GetProps()["hidden"]);
        var metadata=again.GetProps()["advancedFilter"];Assert.Throws<WriterException>(()=>again.SetProp("advancedFilter","{\"range\":\"A1:XFE5\"}"));Assert.Equal(metadata,again.GetProps()["advancedFilter"]);
        foreach(var invalid in new[]{"{\"range\":\"A1:C5\",\"hidden\":true}","{\"range\":\"A1:C5\",\"manualHidden\":[-1]}","{\"range\":\"A1:C5\",\"unique\":\"yes\"}"}){
            Assert.Throws<WriterException>(()=>again.SetProp("advancedFilter",invalid));Assert.Equal(metadata,again.GetProps()["advancedFilter"]);
        }
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((XlsxDocument)opened).Package).Select(e=>e.Description));
        again.SetProp("advancedFilter","null");Assert.Equal("null",again.GetProps()["advancedFilter"]);Assert.Contains("10",again.GetProps()["hidden"]);
    }
}
