using System.Text.RegularExpressions;
using Writer.Core;
using Writer.Formats.Pptx;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Tests;

/// <summary>Slide text written back in place: a shape's html (what the editor saves) changes only the words that differ, and every
/// other run, paragraph, break, field and body setting stays as the file had it. text-runs.pptx mixes them on purpose.</summary>
public class PptxTextTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static byte[] Fixture(string file) => File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("pptx"), file));

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    static A.Paragraph[] Paragraphs(Node shape) => ((P.Shape)shape.Anchor).TextBody!.Elements<A.Paragraph>().ToArray();

    /// <summary>The fidelity audit's edit: every word of the html's text altered.</summary>
    static string Altered(string html) => Regex.Replace(html, @"<[^>]*>|[^<\s]+", m => m.Value[0] == '<' ? m.Value : m.Value + "~");

    [Theory]
    [MemberData(nameof(PptxFidelityTests.Files), MemberType = typeof(PptxFidelityTests))]
    public void Every_shape_retyped_and_set_back_is_the_file_it_was(string file)
    {
        var original = Fixture(file);
        using var doc = new PptxAdapter().Open(new MemoryStream(original));
        foreach (var path in PathResolver.Query(doc.Root, "//shape").Select(s => s.Path).ToList())
        {
            var html = PathResolver.Single(doc.Root, path).GetProps().GetValueOrDefault("html");
            if (string.IsNullOrEmpty(html)) continue;
            Mutations.Set(PathResolver.Single(doc.Root, path), Props(("html", Altered(html))));
            Mutations.Set(PathResolver.Single(doc.Root, path), Props(("html", html)));
        }
        var diff = PackageCompare.Diff(original, Save(doc));
        Assert.True(diff is null, diff);
    }

    [Fact]
    public void Retyping_a_word_rewrites_that_word_only()
    {
        using var doc = new PptxAdapter().Open(new MemoryStream(Fixture("text-runs.pptx")));
        var shape = PathResolver.Single(doc.Root, "/slide[1]/shape[1]");
        var before = ((P.Shape)shape.Anchor).OuterXml;
        Mutations.Set(shape, Props(("html", shape.GetProps()["html"].Replace("twice", "thrice").Replace("Numbered", "Counted"))));
        Assert.Equal(before.Replace("struck twice", "struck thrice").Replace("Numbered from", "Counted from"), ((P.Shape)shape.Anchor).OuterXml);

        // plain text keeps the formatting it cannot state: the German run stays German and unproofed
        var note = PathResolver.Single(doc.Root, "/slide[1]/shape[2]");
        before = ((P.Shape)note.Anchor).OuterXml;
        Mutations.Set(note, Props(("text", "Sidenote Randbemerkung")));
        Assert.Equal(before.Replace("Randnotiz", "Randbemerkung"), ((P.Shape)note.Anchor).OuterXml);
    }

    [Fact]
    public void Runs_read_and_write_baseline_caps_spacing_and_underline_styles()
    {
        using var doc = new PptxAdapter().Open(new MemoryStream(Fixture("text-runs.pptx")));
        var shape = PathResolver.Single(doc.Root, "/slide[1]/shape[1]");
        var runs = shape.Children[1].Children.Select(r => r.GetProps()).ToList();
        Assert.Equal(("small", "3pt"), (runs[0]["caps"], runs[0]["spacing"]));
        Assert.Equal(("all", "wavy"), (runs[2]["caps"], runs[2]["underlineStyle"]));
        Assert.Equal("superscript", runs[5]["vertAlign"]); // baseline 50%
        Assert.Equal(("subscript", "double"), (shape.Children[0].Children[3].GetProps()["vertAlign"], shape.Children[0].Children[1].GetProps()["underlineStyle"]));
        var html = shape.GetProps()["html"];
        Assert.Contains("<sub>", html);
        Assert.Contains("font-variant:small-caps", html);
        Assert.Contains("text-transform:uppercase", html);
        Assert.Contains("letter-spacing:3pt", html);
        Assert.Contains("text-decoration-style:double", html);

        // html states them: a new word takes them, a changed one loses or swaps them, and what html cannot say stays
        var edited = html.Replace("Numbered from III", "Numbered from III <sup>th</sup> <span style=\"letter-spacing:-0.5pt;text-transform:uppercase\">tight</span> <u style=\"text-decoration-style:dotted\">dots</u>")
            .Replace("<span style=\"font-size:18pt;letter-spacing:3pt;font-variant:small-caps\">Small caps</span>", "<span style=\"font-size:18pt\">Small caps</span>");
        Assert.NotEqual(html, edited);
        Mutations.Set(shape, Props(("html", edited)));
        var numbered = Paragraphs(shape)[3].Elements<A.Run>().ToList();
        Assert.Equal(30000, numbered.Single(r => r.Text!.Text == "th").RunProperties!.Baseline!.Value);
        var tight = numbered.Single(r => r.Text!.Text == "tight").RunProperties!;
        Assert.Equal((-50, "all", true), (tight.Spacing!.Value, tight.Capital!.InnerText, tight.Italic!.Value)); // italic from the run it continues
        Assert.Equal("dotted", numbered.Single(r => r.Text!.Text == "dots").RunProperties!.Underline!.InnerText);
        // plain again: the words take the properties of the plain run beside them, and join it
        var small = Paragraphs(shape)[1].Elements<A.Run>().First();
        Assert.Equal(("Small caps and ", null, null), (small.Text!.Text, small.RunProperties!.Capital, small.RunProperties.Spacing));
        var caps = Mutations.Refresh(shape).Children[1].Children.Single(r => r.Text == "all caps").GetProps();
        Assert.Equal(("all", "wavy"), (caps["caps"], caps["underlineStyle"])); // untouched: still wavyHeavy with its green line
        Assert.Contains("u=\"wavyHeavy\"", Paragraphs(shape)[1].OuterXml);

        var run = Mutations.Set(PathResolver.Single(doc.Root, "/slide[1]/shape[1]/paragraph[3]/run[1]"), Props(("vertAlign", "subscript"), ("caps", "small"), ("spacing", "1.5pt"), ("underlineStyle", "double")));
        Assert.Equal(("subscript", "small", "1.5pt", "true", "double"), (run.GetProps()["vertAlign"], run.GetProps()["caps"], run.GetProps()["spacing"], run.GetProps()["underline"], run.GetProps()["underlineStyle"]));
    }

    [Fact]
    public void New_lines_take_their_neighbours_paragraph_and_the_run_they_continue()
    {
        using var doc = new PptxAdapter().Open(new MemoryStream(Fixture("text-runs.pptx")));
        var shape = PathResolver.Single(doc.Root, "/slide[1]/shape[1]");
        var html = shape.GetProps()["html"];
        var bullet = Paragraphs(shape)[1];
        var at = html.IndexOf("<br>", html.IndexOf(" up high", StringComparison.Ordinal), StringComparison.Ordinal); // the end of the bullet line
        Mutations.Set(shape, Props(("html", html[..at] + "<br><sup><span style=\"font-size:18pt\">more</span></sup>" + html[at..])));
        var paragraphs = Paragraphs(shape);
        Assert.Equal(6, paragraphs.Length);
        Assert.Same(bullet, paragraphs[1]);
        Assert.Equal(bullet.ParagraphProperties!.OuterXml, paragraphs[2].ParagraphProperties!.OuterXml); // the § bullet, its level and indents
        Assert.Equal(bullet.GetFirstChild<A.EndParagraphRunProperties>()!.OuterXml, paragraphs[2].GetFirstChild<A.EndParagraphRunProperties>()!.OuterXml);
        Assert.Equal(bullet.Elements<A.Run>().Last().RunProperties!.OuterXml, paragraphs[2].Elements<A.Run>().Single().RunProperties!.OuterXml);

        // text typed into an empty paragraph looks like its end-of-paragraph mark
        var last = Paragraphs(shape)[^1];
        Mutations.Set(shape, Props(("html", shape.GetProps()["html"] + "<b><span style=\"font-size:10pt\">typed</span></b>")));
        var typed = last.Elements<A.Run>().Single().RunProperties!;
        Assert.Equal(("en-US", 1000, true, false), (typed.Language!.Value, typed.FontSize!.Value, typed.Bold!.Value, typed.HasChildren));
        Assert.NotNull(last.GetFirstChild<A.EndParagraphRunProperties>());

        // cleared by one save and typed into at the next (autosave), the text looks as it did; text without a size, colour or
        // font of its own keeps the run's, as the editor shows it
        var note = PathResolver.Single(doc.Root, "/slide[1]/shape[2]");
        Mutations.Set(note, Props(("html", "")));
        Assert.Equal("<a:endParaRPr lang=\"de-DE\" noProof=\"1\" />", Paragraphs(note).Single().GetFirstChild<A.EndParagraphRunProperties>()!.OuterXml.Replace(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"", ""));
        Mutations.Set(note, Props(("html", "Neu")));
        Assert.Equal(("de-DE", true), (Paragraphs(note).Single().Elements<A.Run>().Single().RunProperties!.Language!.Value, Paragraphs(note).Single().Elements<A.Run>().Single().RunProperties!.NoProof!.Value));
        Mutations.Set(shape, Props(("html", shape.GetProps()["html"].Replace("<span style=\"font-size:16pt\">中文段落，右对齐。</span>", "中文段落"))));
        Assert.Equal(1600, Paragraphs(shape)[3].Elements<A.Run>().Single().RunProperties!.FontSize!.Value);
    }

    [Fact]
    public void Writing_the_list_and_level_a_paragraph_has_keeps_its_own_bullet()
    {
        using var doc = new PptxAdapter().Open(new MemoryStream(Fixture("text-runs.pptx")));
        var shape = PathResolver.Single(doc.Root, "/slide[1]/shape[1]");
        var before = ((P.Shape)shape.Anchor).OuterXml;
        Mutations.Set(shape.Children[1], Props(("list", "bullet"), ("level", "1")));
        Mutations.Set(shape.Children[3], Props(("list", "number"), ("level", "0")));
        Mutations.Set(shape.Children[0], Props(("list", "none"), ("level", "0")));
        Assert.Equal(before, ((P.Shape)shape.Anchor).OuterXml);
        Mutations.Set(shape.Children[1], Props(("level", "2")));
        Assert.Equal(("bullet", "2"), (shape.Children[1].GetProps()["list"], Mutations.Refresh(shape).Children[1].GetProps()["level"]));
    }
}
