using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Html;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Paragraph shading (w:pPr/w:shd): the paragraph's own fill, the one its style gives (computed), written and kept.</summary>
public class DocxShadingTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static W.Paragraph Shaded(string text, string fill, string? style = null)
    {
        var p = P(text, style);
        (p.ParagraphProperties ??= new W.ParagraphProperties()).Shading = new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = fill };
        return p;
    }

    static void Note(MainDocumentPart main) => main.StyleDefinitionsPart!.Styles!.Append(new W.Style(
        new W.StyleName { Val = "Note" }, new W.StyleParagraphProperties(new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = "FFF2CC" }))
        { Type = W.StyleValues.Paragraph, StyleId = "Note" });

    [Fact]
    public void Paragraph_shading_is_read_from_the_paragraph_or_its_style_written_and_kept_when_the_text_changes()
    {
        using var doc = OpenDocx(Docx([Shaded("Light grey", "d9d9d9"), P("By its style", "Note"), Shaded("Cleared over the style", "auto", "Note"), P("Plain")], Note));
        Node Para(int i) => PathResolver.Single(doc.Root, $"/body/paragraph[{i}]");
        Assert.Equal("D9D9D9", Para(1).GetProps()["fill"]);
        Assert.False(Para(2).GetProps().ContainsKey("fill"));
        Assert.Equal("FFF2CC", Para(2).GetComputed(Para(2).GetProps())?.GetValueOrDefault("fill"));
        Assert.Null(Para(3).GetComputed(Para(3).GetProps())?.GetValueOrDefault("fill"));
        Assert.Null(Para(4).GetComputed(Para(4).GetProps())?.GetValueOrDefault("fill"));

        Mutations.Set(Para(1), Props(("html", "Light grey, <b>edited</b>")));
        Assert.Equal("D9D9D9", Para(1).GetProps()["fill"]);
        Assert.Equal("1F3864", Mutations.Set(Para(4), Props(("fill", "1F3864"))).GetProps()["fill"]);
        Assert.False(Mutations.Set(Para(4), Props(("fill", "none"))).GetProps().ContainsKey("fill"));
        var cleared = Mutations.Set(Para(2), Props(("fill", "none")));
        Assert.Null(cleared.GetComputed(cleared.GetProps())?.GetValueOrDefault("fill")); // none over the style's shading says so
        var heading = Mutations.Add(doc.Root.Children.Single(), "heading", Props(("text", "Shaded heading"), ("fill", "DDEBF7")), null);
        Assert.Equal("DDEBF7", heading.GetProps()["fill"]);
    }

    [Fact]
    public void The_html_view_draws_paragraph_shading_with_white_automatic_text_on_a_dark_one()
    {
        using var doc = OpenDocx(Docx(Shaded("Light", "D9D9D9"), Shaded("Dark", "1F3864")));
        var html = HtmlWriter.Render(doc);
        Assert.Contains("<p style=\"background:#D9D9D9\">Light</p>", html);
        Assert.Contains("<p style=\"background:#1F3864;color:#FFFFFF\">Dark</p>", html);
    }
}
