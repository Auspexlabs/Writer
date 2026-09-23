using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Page breaks inside a paragraph (w:br w:type="page", Word's Ctrl+Enter after text) and paragraphs that start a page
/// (w:pageBreakBefore): read as such, and written back as such when the paragraph is edited, never as a line break.</summary>
public class DocxPageBreakTests
{
    const string PageBreakHtml = "<br style=\"page-break-before:always\">";

    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static Document Reopen(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return OpenDocx(ms.ToArray());
    }

    static W.Paragraph Broken(string before, string after) =>
        new(new W.Run(new W.Text(before) { Space = SpaceProcessingModeValues.Preserve }), new W.Run(new W.Break { Type = W.BreakValues.Page }),
            new W.Run(new W.Text(after) { Space = SpaceProcessingModeValues.Preserve }));

    static int PageBreaks(Node paragraph) => ((W.Paragraph)paragraph.Anchor).Descendants<W.Break>().Count(b => b.Type?.Value == W.BreakValues.Page);
    static int LineBreaks(Node paragraph) => ((W.Paragraph)paragraph.Anchor).Descendants<W.Break>().Count(b => b.Type is null);

    [Fact]
    public void A_page_break_inside_a_paragraph_reads_as_one_and_an_edit_writes_it_back_as_one()
    {
        using var doc = OpenDocx(Docx(Broken("Page one ends", "page two starts")));
        var p = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        Assert.Equal("Page one ends\fpage two starts", p.Text);
        Assert.Equal("Page one ends" + PageBreakHtml + "page two starts", p.GetProps()["html"]);

        // the editor sends the paragraph back with a word added: the break is still a page break, not a line break
        Mutations.Set(p, Props(("html", "Page one ends here" + PageBreakHtml + "page two starts")));
        using var back = Reopen(doc);
        p = PathResolver.Single(back.Root, "/body/paragraph[1]");
        Assert.Equal("Page one ends here\fpage two starts", p.Text);
        Assert.Equal((1, 0), (PageBreaks(p), LineBreaks(p)));

        Mutations.Set(p, Props(("html", "one<br>two" + PageBreakHtml + "three")));
        Assert.Equal((1, 1), (PageBreaks(p), LineBreaks(p)));
        Mutations.Set(p, Props(("text", "plain\ftext")));
        Assert.Equal((1, 0), (PageBreaks(p), LineBreaks(p)));
    }

    [Fact]
    public void A_page_break_in_a_table_cell_is_written_back_as_one()
    {
        var table = new W.Table(new W.TableRow(new W.TableCell(Broken("above", "below"))));
        using var doc = OpenDocx(Docx(table));
        var cell = PathResolver.Single(doc.Root, "/body/table[1]/row[1]/cell[1]");
        Assert.Equal("above" + PageBreakHtml + "below", cell.GetProps()["html"]);
        Mutations.Set(cell, Props(("html", "above, edited" + PageBreakHtml + "below")));
        Assert.Equal(1, ((W.TableCell)cell.Anchor).Descendants<W.Break>().Count(b => b.Type?.Value == W.BreakValues.Page));
    }

    [Fact]
    public void Page_break_before_is_read_from_the_paragraph_or_its_style_and_kept_when_the_text_changes()
    {
        static void Chapter(MainDocumentPart main) => main.StyleDefinitionsPart!.Styles!.Append(new W.Style(
            new W.StyleName { Val = "Chapter" }, new W.StyleParagraphProperties(new W.PageBreakBefore())) { Type = W.StyleValues.Paragraph, StyleId = "Chapter" });
        var direct = P("Starts a page");
        direct.PrependChild(new W.ParagraphProperties(new W.PageBreakBefore()));
        using var doc = OpenDocx(Docx([P("First"), direct, P("By its style", "Chapter"), P("Plain")], Chapter));
        var (own, styled, plain) = (PathResolver.Single(doc.Root, "/body/paragraph[2]"), PathResolver.Single(doc.Root, "/body/paragraph[3]"), PathResolver.Single(doc.Root, "/body/paragraph[4]"));
        Assert.Equal("true", own.GetProps()["pageBreakBefore"]);
        Assert.False(styled.GetProps().ContainsKey("pageBreakBefore"));
        Assert.Equal("true", styled.GetComputed(styled.GetProps())?.GetValueOrDefault("pageBreakBefore"));
        Assert.Null(plain.GetComputed(plain.GetProps())?.GetValueOrDefault("pageBreakBefore"));

        Mutations.Set(own, Props(("html", "Starts a page, <b>edited</b>")));
        using var back = Reopen(doc);
        Assert.Equal("true", PathResolver.Single(back.Root, "/body/paragraph[2]").GetProps()["pageBreakBefore"]);
        var set = Mutations.Set(PathResolver.Single(back.Root, "/body/paragraph[4]"), Props(("pageBreakBefore", "true")));
        Assert.Equal("true", set.GetProps()["pageBreakBefore"]);
        Assert.False(Mutations.Set(set, Props(("pageBreakBefore", "false"))).GetProps().ContainsKey("pageBreakBefore"));
    }

    [Fact]
    public void Exports_to_other_formats_carry_a_page_break_as_a_line_break()
    {
        using var doc = OpenDocx(Docx(Broken("Page one ends", "page two starts")));
        foreach (var adapter in new IFormatAdapter[] { new Writer.Formats.Markdown.MdAdapter(), new Writer.Formats.Pptx.PptxAdapter() })
        {
            var (target, _) = Exporter.Export(doc, adapter, null);
            using (target)
            {
                var text = string.Concat(PathResolver.Query(target.Root, "//paragraph").Select(n => n.Text));
                Assert.Contains("Page one ends", text);
                Assert.DoesNotContain('\f', text);
                target.Save(new MemoryStream()); // a form feed is not a character XML can hold
            }
        }
    }
}
