using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Docx;
using Writer.Formats.Html;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>What a Word document looks like where its paragraphs leave it to styles and the theme (computed), and the html view drawing
/// that, with each section's page, columns, header and notes where Word puts them.</summary>
public class DocxLookTests
{
    /// <summary>A document as a Chinese Word writes it: theme fonts in the defaults (the theme leaves the East Asian ones to its Hans entry),
    /// 正文 with a 2-character first line, and a heading style that cancels it and sets its own fonts, size and colour.</summary>
    static Document Themed() => OpenDocx(Docx([P("标题", "1"), P("正文段落")], main =>
    {
        main.StyleDefinitionsPart!.Styles = new W.Styles("""
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault>
            <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
            <w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:ind w:firstLineChars="200" w:firstLine="420"/></w:pPr></w:style>
            <w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/><w:pPr><w:spacing w:before="340" w:after="330"/><w:ind w:firstLineChars="0" w:firstLine="0"/><w:outlineLvl w:val="0"/></w:pPr>
            <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"/><w:b/><w:sz w:val="44"/><w:color w:val="2F5496"/></w:rPr></w:style>
            </w:styles>
            """);
        (main.ThemePart ?? main.AddNewPart<ThemePart>()).Theme = new A.Theme("""
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:fontScheme name="Office">
            <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hans" typeface="等线 Light"/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游明朝"/><a:font script="Hans" typeface="等线"/></a:minorFont>
            </a:fontScheme></a:themeElements></a:theme>
            """);
    }));

    static IReadOnlyDictionary<string, string> Computed(Node n) => n.GetComputed(n.GetProps()) ?? new Dictionary<string, string>();

    [Fact]
    public void Styles_defaults_and_theme_fonts_give_the_look_of_the_document_and_of_each_paragraph()
    {
        using var doc = Themed();
        var root = Computed(doc.Root);
        Assert.Equal(("Calibri", "等线", "10.5", "1.08", "8pt", "2ch"), (root["font"], root["fontEa"], root["size"], root["lineSpacing"], root["spaceAfter"], root["indentFirst"]));
        var heading = Computed(PathResolver.Single(doc.Root, "/body/heading[1]")); // only what differs from the document's look
        Assert.Equal(("Calibri Light", "等线 Light", "22", "true", "2F5496"), (heading["font"], heading["fontEa"], heading["size"], heading["bold"], heading["color"]));
        Assert.Equal(("17pt", "16.5pt", "0"), (heading["spaceBefore"], heading["spaceAfter"], heading["indentFirst"])); // its 0 undoes 正文's first line
        Assert.DoesNotContain("size", Computed(PathResolver.Single(doc.Root, "/body/paragraph[1]")).Keys);

        var html = HtmlWriter.Render(doc);
        Assert.Contains(".docx .page{font-family:'Calibri','等线';font-size:10.5pt}", html);
        Assert.Matches(@"\.docx \.page p,[^{]*\{line-height:1\.296;padding-bottom:8pt;text-indent:2em\}", html);
        Assert.Contains("<h1 style=\"padding-top:17pt;padding-bottom:16.5pt;text-indent:0;font-family:'Calibri Light','等线 Light';font-size:22pt;font-weight:bold;color:#2F5496\">标题</h1>", html);
        Assert.Contains("<p>正文段落</p>", html);
    }

    [Fact]
    public void Each_section_is_a_page_of_its_size_with_its_columns_and_notes()
    {
        using var doc = new DocxAdapter().Open(File.OpenRead(Path.Combine(FixtureDir("docx"), "sections.docx")));
        var html = HtmlWriter.Render(doc);
        Assert.Equal(3, System.Text.RegularExpressions.Regex.Count(html, "<section class=\"page\""));
        Assert.Contains("<section class=\"page\" style=\"width:21cm;min-height:29.7cm;padding:2.54cm 2.54cm 2.54cm 2.54cm\">\n<div style=\"column-count:2;column-gap:1cm\">", html);
        Assert.Contains("<section class=\"page\" style=\"width:29.7cm;min-height:21cm;padding:2cm 1.5cm 2cm 3cm\">", html); // the landscape section
        Assert.Matches("<aside class=\"notes\">\n<p><sup>1</sup> Column width", html); // the first section's footnotes under its text
        Assert.Contains("<sup class=\"note\">i</sup>", html); // endnotes count i, ii… and gather after the last page
        Assert.EndsWith("</section>\n<aside class=\"notes\">\n<p><sup>i</sup> Endnotes are collected per endnotePr.pos; here they gather at the document end.</p>\n<p><sup>ii</sup> Upper-Roman numbering restarts each section under endnotePr.numRestart=eachSect.</p>\n</aside>\n</body>\n</html>\n", html);
    }

    [Fact]
    public void Floating_pictures_wrap_and_tables_keep_their_column_widths()
    {
        string Html(string file) { using var doc = new DocxAdapter().Open(File.OpenRead(Path.Combine(FixtureDir("docx"), file))); return HtmlWriter.Render(doc); }
        var pictures = Html("pictures.docx");
        Assert.Contains("width:132.3px;height:132.3px;float:right;margin:0 0 4px 12px\" alt=\"Logo floated right with square wrap\">", pictures); // square wrap, right
        Assert.Matches("<p style=\"position:relative\"><img [^>]*position:absolute;z-index:-1;left:50%", pictures); // the watermark behind the text
        Assert.Contains("<colgroup><col style=\"width:2.44cm\">", Html("tables.docx"));
    }

    [Fact]
    public void Headers_and_footers_sit_in_the_margins()
    {
        using var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, new Dictionary<string, string> { ["header"] = "Report", ["footer"] = "Page {page} of {pages}" });
        var html = HtmlWriter.Render(doc);
        Assert.Contains("<header style=\"left:2.54cm;right:2.54cm;top:1.27cm\">Report</header>", html);
        Assert.Contains("<footer style=\"left:2.54cm;right:2.54cm;bottom:1.27cm\">Page 1 of 1</footer>", html);
    }
}
