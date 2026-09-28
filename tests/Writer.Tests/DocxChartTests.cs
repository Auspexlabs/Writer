using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class DocxChartTests
{
    [Fact]
    public void Word_chart_embeds_editable_workbook_and_keeps_its_place_on_update()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, new Dictionary<string, string> { ["chartData"] = """{"title":"Sales","type":"column","labels":["Jan","Feb"],"values":[4,-7]}""" });
        var chart = PathResolver.Single(doc.Root, "/body/object[1]");
        Assert.Equal("chart", chart.GetProps()["type"]);
        Assert.Contains("Jan", chart.GetProps()["chart"]);
        var bytes = new MemoryStream(); doc.Save(bytes);
        using var reopened = OpenDocx(bytes.ToArray());
        var package = ((DocxDocument)reopened).Package;
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(package).Select(e => e.Description));
        var part = package.MainDocumentPart!.ChartParts.Single();
        var embedded = (EmbeddedPackagePart)part.GetPartById(part.ChartSpace!.GetFirstChild<C.ExternalData>()!.Id!);
        using (var source = embedded.GetStream(FileMode.Open, FileAccess.Read)) using (var book = SpreadsheetDocument.Open(source, false)) Assert.Contains("-7", book.WorkbookPart!.WorksheetParts.Single().Worksheet!.OuterXml);
        var old = chart.GetRaw();
        Mutations.Set(reopened.Root, new Dictionary<string, string> { ["chartData"] = """{"path":"/body/object[1]","title":"New","type":"line","labels":["Mar"],"values":[22]}""" });
        var updated = PathResolver.Single(reopened.Root, "/body/object[1]");
        Assert.Contains("Mar", updated.GetProps()["chart"]);
        Assert.Contains("New", updated.GetProps()["chart"]);
        Assert.Single(reopened.Root.Children.Single().Children);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(package).Select(e => e.Description));
    }
}
