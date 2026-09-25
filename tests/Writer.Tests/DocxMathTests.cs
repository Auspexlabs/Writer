using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Equations: LaTeX written as Word's Office Math and read back as LaTeX that reads the same again.</summary>
public class DocxMathTests
{
    static string Twice(string latex) => OfficeMathTex.ToLatex(OfficeMathTex.ToOmml(latex));

    [Theory]
    [InlineData(@"\frac{a}{b}", @"\frac{a}{b}")]
    [InlineData(@"x^2+y_1", @"x^{2}+y_{1}")]
    [InlineData(@"x_1^2", @"x_{1}^{2}")]
    [InlineData(@"e^{i\pi}+1=0", @"e^{i\pi}+1=0")]
    [InlineData(@"\sqrt{x}+\sqrt[3]{y}", @"\sqrt{x}+\sqrt[3]{y}")]
    [InlineData(@"\sum_{i=1}^{n} i^2", @"\sum_{i=1}^{n}{i^{2}}")]
    [InlineData(@"\int_0^1 f(x)\,dx", @"\int_{0}^{1}{f}(x)\,dx")]
    [InlineData(@"\lim_{x\to 0}\frac{\sin x}{x}", @"\lim_{x\to0}\frac{\sin x}{x}")]
    [InlineData(@"\left( \frac{a}{b} \right)", @"\left(\frac{a}{b}\right)")]
    [InlineData(@"\begin{pmatrix} a & b \\ c & d \end{pmatrix}", @"\begin{pmatrix}a&b\\c&d\end{pmatrix}")]
    [InlineData(@"f(x)=\begin{cases} x & x>0 \\ -x & x\le 0 \end{cases}", @"f(x)=\begin{cases}x&x>0\\-x&x\leq0\end{cases}")]
    [InlineData(@"\hat{x}+\vec{v}+\overline{AB}", @"\hat{x}+\vec{v}+\overline{AB}")]
    [InlineData(@"\alpha\beta\Gamma \leq \infty", @"\alpha\beta\Gamma\leq\infty")]
    [InlineData(@"\text{if } x \in \mathbb{R}", @"\text{if }x\in\mathbb{R}")]
    [InlineData(@"\binom{n}{k}", @"\binom{n}{k}")]
    public void Latex_round_trips_through_office_math(string latex, string expected)
    {
        var once = Twice(latex);
        Assert.Equal(expected, once);
        Assert.Equal(once, Twice(once)); // what is read back reads the same again
    }

    [Fact]
    public void Office_math_has_the_structures_word_draws()
    {
        var xml = OfficeMathTex.ToOmml(@"\sum_{i=1}^{n}\frac{1}{i}+\left[x\right]+\begin{bmatrix}1&0\\0&1\end{bmatrix}+\sqrt{2}").OuterXml;
        foreach (var tag in new[] { "<m:nary>", "m:val=\"∑\"", "m:val=\"undOvr\"", "<m:f>", "<m:num>", "<m:den>", "<m:d>", "m:val=\"[\"", "<m:m>", "<m:mr>", "<m:count m:val=\"2\"/>", "<m:rad>", "<m:degHide m:val=\"1\"/>", "Cambria Math" })
            Assert.Contains(tag, xml);
    }

    [Fact]
    public void Equations_sit_in_paragraphs_at_offsets_and_survive_reopen()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", new Dictionary<string, string> { ["text"] = "Area is  here" }, null);
        Mutations.Add(p, "equation", new Dictionary<string, string> { ["latex"] = @"\pi r^2", ["at"] = "8" }, null);
        var display = Mutations.Add(body, "paragraph", new Dictionary<string, string>(), null);
        Mutations.Add(display, "equation", new Dictionary<string, string> { ["latex"] = @"E=mc^2", ["display"] = "true" }, null);
        Assert.Equal("Area is  here", p.GetProps()["text"]);
        var errors = new OpenXmlValidator().Validate(((DocxDocument)doc).Package).Where(e => e.Part is MainDocumentPart).Select(e => e.Path?.XPath + ": " + e.Description).ToList();
        Assert.Empty(errors);

        var ms = new MemoryStream();
        doc.Save(ms);
        using var reopened = OpenDocx(ms.ToArray());
        var eqs = PathResolver.Query(reopened.Root, "//equation").Select(n => n.GetProps()).ToList();
        Assert.Equal((@"\pi r^{2}", "8"), (eqs[0]["latex"], eqs[0]["at"]));
        Assert.False(eqs[0].ContainsKey("display"));
        Assert.Equal((@"E=mc^{2}", "true"), (eqs[1]["latex"], eqs[1]["display"]));

        var first = PathResolver.Single(reopened.Root, "/body/paragraph[1]/equation[1]");
        first = Mutations.Set(first, new Dictionary<string, string> { ["latex"] = @"\frac{1}{2}", ["at"] = "0" });
        Assert.Equal((@"\frac{1}{2}", "0"), (first.GetProps()["latex"], first.GetProps()["at"]));
        first.Remove();
        Assert.Single(PathResolver.Query(reopened.Root, "//equation"));
        Assert.Equal("Area is  here", PathResolver.Single(reopened.Root, "/body/paragraph[1]").GetProps()["text"]);
    }
}
