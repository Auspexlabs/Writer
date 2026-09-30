using Writer.Core;
using Writer.Formats.Pptx;
using A = DocumentFormat.OpenXml.Drawing;
namespace Writer.Tests;
public class PptxFontReplaceTests
{
    [Fact]
    public void Named_font_replacement_reaches_slide_and_theme_without_changing_other_fonts_or_text()
    {
        using var doc = new PptxAdapter().Create();
        var slide = Mutations.Add(doc.Root, "slide", new Dictionary<string,string>(), null);
        var shape = Mutations.Add(slide, "shape", new Dictionary<string,string> { ["text"] = "Keep all text", ["font"] = "Arial" }, null);
        var other = Mutations.Add(slide, "shape", new Dictionary<string,string> { ["text"] = "Georgia stays", ["font"] = "Georgia" }, null);
        var presentation = ((PptxDocument)doc).Presentation;
        var master = presentation.SlideMasterParts.First();
        foreach (var font in master.ThemePart!.Theme!.Descendants<A.LatinFont>()) font.Typeface = "Arial";
        Mutations.Set(doc.Root, new Dictionary<string,string> { ["replaceFont"] = "{\"from\":\"arial\",\"to\":\"Aptos\"}" });
        using var bytes = new MemoryStream(); doc.Save(bytes); bytes.Position=0;
        using var reopened = new PptxAdapter().Open(bytes);
        Assert.Equal("Aptos",PathResolver.Single(reopened.Root,shape.Path).GetProps()["font"]);
        Assert.Equal("Georgia",PathResolver.Single(reopened.Root,other.Path).GetProps()["font"]);
        Assert.Equal("Keep all text",PathResolver.Single(reopened.Root,shape.Path).GetProps()["text"]);
        Assert.All(((PptxDocument)reopened).Presentation.SlideMasterParts.First().ThemePart!.Theme!.Descendants<A.LatinFont>(), font=>Assert.Equal("Aptos",font.Typeface!.Value));
    }
}
