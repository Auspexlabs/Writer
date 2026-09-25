using Writer.Core;
using Writer.Formats.Common;
using Writer.Formats.Docx;

namespace Writer.Tests;

/// <summary>Character formatting beyond bold and colour: superscript and subscript, character spacing, Word's outline and shadow
/// effects, and the highlight pen (w:highlight for its fifteen colours), through html and run props both ways.</summary>
public class DocxCharacterTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    [Fact]
    public void Superscript_spacing_effects_and_the_highlight_pen_round_trip_through_html_and_props()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var html = "x<sup>2</sup> <span style=\"letter-spacing:2pt\">wide</span> <span style=\"" + InlineHtml.ShadowCss + "\">shade</span> <span style=\"" + InlineHtml.OutlineCss + "\">line</span> <span style=\"background-color:#FFFF00\">pen</span> <span style=\"background-color:#FFF2A8\">shading</span>";
        var p = Mutations.Add(body, "paragraph", Props(("html", html)), null);
        Assert.Equal(html, p.GetProps()["html"]);
        var runs = p.Children.Where(c => c.Kind == "run" && c.Text!.Trim().Length > 0).ToDictionary(r => r.Text!.Trim(), r => r.GetProps());
        Assert.Equal("superscript", runs["2"]["vertAlign"]);
        Assert.Equal("2pt", runs["wide"]["spacing"]);
        Assert.Equal("true", runs["shade"]["shadow"]);
        Assert.Equal("true", runs["line"]["outline"]);
        Assert.Equal("FFFF00", runs["pen"]["highlight"]);
        Assert.Equal("FFF2A8", runs["shading"]["highlight"]);
        var raw = p.GetRaw();
        Assert.Contains("<w:vertAlign w:val=\"superscript\"", raw);
        Assert.Contains("<w:spacing w:val=\"40\"", raw);
        Assert.Contains("<w:shadow", raw);
        Assert.Contains("<w:outline", raw);
        Assert.Contains("<w:highlight w:val=\"yellow\"", raw);
        Assert.Contains("w:fill=\"FFF2A8\"", raw);

        var run = p.Children.First(c => c.Kind == "run" && c.Text!.Trim() == "wide");
        run = Mutations.Set(run, Props(("vertAlign", "subscript"), ("spacing", "-0.5pt"), ("highlight", "00FF00")));
        Assert.Equal(("subscript", "-0.5pt", "00FF00"), (run.GetProps()["vertAlign"], run.GetProps()["spacing"], run.GetProps()["highlight"]));
        Assert.Contains("w:val=\"green\"", run.GetRaw());
        run = Mutations.Set(run, Props(("vertAlign", "baseline"), ("spacing", "none"), ("highlight", "none")));
        Assert.False(run.GetProps().ContainsKey("vertAlign"));
        Assert.False(run.GetProps().ContainsKey("spacing"));
        Assert.False(run.GetProps().ContainsKey("highlight"));

        Mutations.Set(p, Props(("html", "x2 wide shade line pen shading"))); // the editor took every mark off
        Assert.DoesNotContain("vertAlign", p.GetRaw());
        Assert.DoesNotContain("w:outline", p.GetRaw());
        Assert.DoesNotContain("w:highlight", p.GetRaw());
    }
}
