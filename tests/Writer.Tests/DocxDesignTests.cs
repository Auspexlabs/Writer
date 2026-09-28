using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class DocxDesignTests
{
    [Fact]
    public void Theme_style_set_borders_and_readonly_protection_round_trip()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, new Dictionary<string, string> { ["theme"] = "mist", ["styleSet"] = "formal", ["pageBorder"] = """{"style":"double","color":"112233","width":1.5,"space":24}""", ["protected"] = "true" });
        var stream = new MemoryStream(); doc.Save(stream);
        using var reopened = OpenDocx(stream.ToArray()); var native = (DocxDocument)reopened;
        Assert.Equal("Writer mist", reopened.Root.GetProps()["theme"]);
        Assert.Equal("true", reopened.Root.GetProps()["protected"]);
        Assert.Contains("112233", reopened.Root.GetProps()["pageBorder"]);
        Assert.Contains("24", native.Main.StyleDefinitionsPart!.Styles!.Elements<W.Style>().Single(s => s.StyleId == "Normal").OuterXml);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(native.Package).Select(e => e.Description));
        Mutations.Set(reopened.Root, new Dictionary<string, string> { ["protected"] = "false", ["pageBorder"] = """{"style":"none"}""" });
        Assert.False(reopened.Root.GetProps().ContainsKey("protected"));
        Assert.False(reopened.Root.GetProps().ContainsKey("pageBorder"));
    }
}
