using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Citations as Word makes them: the sources in Word's bibliography part, a citation in the text as a citation content control
/// around a CITATION field, a works-cited list as a Bibliographies control around a BIBLIOGRAPHY field, Chicago's notes as footnotes;
/// all drawn again from the sources whenever a source, a citation or the style changes.</summary>
public class DocxCitationTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    const string PeggJson = """{"type":"article","authors":[{"last":"Pegg","first":"Ian L."}],"title":"Behavior of technetium in nuclear waste vitrification processes","container":"Journal of Radioanalytical and Nuclear Chemistry","year":"2015","volume":"305","issue":"1","pages":"287-292","doi":"10.1007/s10967-014-3900-9"}""";
    const string KuhnJson = """{"type":"book","authors":["Thomas S. Kuhn"],"title":"The structure of scientific revolutions","publisher":"University of Chicago Press","year":"1962"}""";

    static Document Reopen(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return OpenDocx(ms.ToArray());
    }

    static void AssertValid(Document doc) =>
        Assert.Empty(new OpenXmlValidator().Validate(((DocxDocument)doc).Package).Where(e => e.Part is MainDocumentPart or FootnotesPart).Select(e => e.Path?.XPath + ": " + e.Description));

    static Node Body(Document doc) => doc.Root.Children.Single();

    static (Document Doc, Node Para) Essay(string style = "mla")
    {
        var doc = new DocxAdapter().Create();
        Mutations.Set(doc.Root, Props(("citationStyle", style), ("source", PeggJson)));
        var p = Mutations.Add(Body(doc), "paragraph", Props(("text", "Pegg presents the experimental heart of the question .")), null); // a citation goes after the space before the period
        return (doc, p);
    }

    [Theory]
    [InlineData("ieee")]
    [InlineData("gb7714")]
    public void Numeric_citations_renumber_after_paragraph_reorder_and_keep_bibliography_in_sync(string style)
    {
        var (doc, first) = Essay(style); using var owned = doc;
        Mutations.Set(doc.Root, Props(("source", KuhnJson)));
        var second = Mutations.Add(Body(doc), "paragraph", Props(("text", "Second ")), null);
        Mutations.Add(first, "citation", Props(("sources", "Peg15")), null);
        Mutations.Add(second, "citation", Props(("sources", "Kuh62")), null);
        Mutations.Add(Body(doc), "bibliography", Props(), null);
        var native = ((DocxDocument)doc).Main.Document!.Body!;
        var a = native.Elements<W.Paragraph>().First(); var b = native.Elements<W.Paragraph>().Skip(1).First(); b.Remove(); native.InsertBefore(b, a);
        using var reopened = Reopen(doc);
        var paragraphs = ((DocxDocument)reopened).Main.Document!.Body!.Elements<W.Paragraph>().ToArray();
        Assert.Contains("[1]", paragraphs[0].InnerText);
        Assert.Contains("[2]", paragraphs[1].InnerText);
        Assert.Equal(style, reopened.Root.GetProps()["citationStyle"]);
        var bibliography = ((DocxDocument)reopened).Main.Document!.Body!.Descendants<W.SdtBlock>().First().InnerText;
        Assert.True(bibliography.IndexOf("Kuhn", StringComparison.OrdinalIgnoreCase) < bibliography.IndexOf("Pegg", StringComparison.OrdinalIgnoreCase));
        AssertValid(reopened);
    }

    [Fact]
    public void Sources_live_in_Words_bibliography_part_with_tags_as_Word_makes_them()
    {
        var (doc, _) = Essay();
        Mutations.Set(doc.Root, Props(("source", KuhnJson)));
        var props = doc.Root.GetProps();
        Assert.Contains("\"tag\":\"Peg15\"", props["sources"]);
        Assert.Contains("\"tag\":\"Kuh62\"", props["sources"]);
        Assert.Contains("\"authors\":[{\"last\":\"Kuhn\",\"first\":\"Thomas S.\"}]", props["sources"]);
        Assert.Equal("mla", props["citationStyle"]);
        var main = ((DocxDocument)doc).Main;
        var part = main.CustomXmlParts.Single();
        string xml; using (var r = new StreamReader(part.GetStream())) xml = r.ReadToEnd();
        Assert.Contains("MLASeventhEditionOfficeOnline.xsl", xml);
        Assert.Contains("<b:Tag>Peg15</b:Tag>", xml.Replace("<Tag>", "<b:Tag>").Replace("</Tag>", "</b:Tag>"));
        Assert.Matches("SourceType>JournalArticle<", xml);
        Assert.Matches("JournalName>Journal of Radioanalytical and Nuclear Chemistry<", xml);
        Assert.Matches("Last>Pegg</(b:)?Last><(b:)?First>Ian</(b:)?First><(b:)?Middle>L.</", xml);
        Assert.Contains("officeDocument/2006/bibliography", part.CustomXmlPropertiesPart!.DataStoreItem!.OuterXml);
        Assert.Empty(new OpenXmlValidator().Validate(((DocxDocument)doc).Package).Select(e => e.Part?.Uri + " " + e.Path?.XPath + ": " + e.Description)); // every part, the bibliography's too

        using var reopened = Reopen(doc);
        Assert.Equal(props["sources"], reopened.Root.GetProps()["sources"]);
        Mutations.Set(reopened.Root, Props(("source", """{"tag":"Kuh62","edition":"2","year":"none"}""")));
        Assert.Contains("\"edition\":\"2\"", reopened.Root.GetProps()["sources"]);
        Assert.DoesNotContain("1962", reopened.Root.GetProps()["sources"]);
        Mutations.Set(reopened.Root, Props(("source", """{"tag":"Kuh62","remove":true}""")));
        Assert.DoesNotContain("Kuh62", reopened.Root.GetProps()["sources"]);
    }

    [Fact]
    public void A_citation_is_a_citation_control_around_a_CITATION_field_at_an_offset_outside_the_paragraphs_text()
    {
        var (doc, p) = Essay();
        var cite = Mutations.Add(p, "citation", Props(("sources", "Peg15"), ("pages", "288"), ("at", "53")), null);
        var props = cite.GetProps();
        Assert.Equal(("(Pegg 288)", "Peg15", "288", "53"), (props["text"], props["sources"], props["pages"], props["at"]));
        Assert.Equal("Pegg presents the experimental heart of the question .", p.GetProps()["text"]); // the citation's text is not the paragraph's
        var raw = cite.GetRaw();
        Assert.Contains("<w:citation", raw);
        Assert.Contains(" CITATION Peg15 \\p 288 \\l 1033 ", raw);
        Assert.Equal("Pegg presents the experimental heart of the question (Pegg 288).", Views.Text(doc.Root).Trim());
        AssertValid(doc);

        Mutations.Set(doc.Root, Props(("citationStyle", "apa")));
        Assert.Equal("(Pegg, 2015, p. 288)", Body(doc).Children[0].Children.Single(c => c.Kind == "citation").GetProps()["text"]);
        Mutations.Set(Body(doc).Children[0].Children.Single(c => c.Kind == "citation"), Props(("noAuthor", "true")));
        Assert.Equal("(2015, p. 288)", Body(doc).Children[0].Children.Single(c => c.Kind == "citation").GetProps()["text"]);

        using var reopened = Reopen(doc);
        var again = PathResolver.Single(reopened.Root, "//citation[@id=" + props["id"] + "]");
        Assert.Equal(("(2015, p. 288)", "true"), (again.GetProps()["text"], again.GetProps()["noAuthor"]));
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(reopened.Root, Props(("source", """{"tag":"Peg15","remove":true}"""))));
        Assert.Contains("cited", ex.Message);
    }

    [Fact]
    public void Writing_the_paragraphs_text_around_a_citation_keeps_it_whole_where_it_is()
    {
        var (doc, p) = Essay();
        Mutations.Add(p, "citation", Props(("sources", "Peg15"), ("pages", "288"), ("at", "53")), null);
        Mutations.Set(p, Props(("html", "Pegg presents the <i>experimental</i> heart of the retention question .")));
        var para = Body(doc).Children[0];
        Assert.Equal("Pegg presents the experimental heart of the retention question .", para.GetProps()["text"]);
        var cite = para.Children.Single(c => c.Kind == "citation").GetProps();
        Assert.Equal(("(Pegg 288)", "63"), (cite["text"], cite["at"]));
        Assert.Equal("Pegg presents the experimental heart of the retention question (Pegg 288).", Views.Text(doc.Root).Trim());
        AssertValid(doc);
    }

    [Fact]
    public void A_works_cited_list_is_Words_bibliography_control_in_the_styles_order_and_follows_the_style_and_the_sources()
    {
        var (doc, p) = Essay();
        Mutations.Set(doc.Root, Props(("source", KuhnJson)));
        Mutations.Add(p, "citation", Props(("sources", "Peg15"), ("at", "53")), null);
        var list = Mutations.Add(Body(doc), "bibliography", Props(), null);
        var props = list.GetProps();
        Assert.Equal("Works Cited", props["title"]);
        Assert.Equal("Kuhn, Thomas S. The Structure of Scientific Revolutions. University of Chicago Press, 1962.\nPegg, Ian L. “Behavior of Technetium in Nuclear Waste Vitrification Processes.” Journal of Radioanalytical and Nuclear Chemistry, vol. 305, no. 1, 2015, pp. 287–92, https://doi.org/10.1007/s10967-014-3900-9.", props["text"]);
        Assert.StartsWith("<p>Kuhn, Thomas S. <i>The Structure of Scientific Revolutions</i>. University", props["html"]);
        var raw = list.GetRaw();
        Assert.Contains("w:docPartGallery w:val=\"Bibliographies\"", raw);
        Assert.Contains("<w:bibliography", raw);
        Assert.Contains(" BIBLIOGRAPHY ", raw);
        Assert.Contains("<w:pageBreakBefore", raw);
        Assert.Contains("w:hanging=\"720\"", raw);
        AssertValid(doc);

        Mutations.Set(doc.Root, Props(("citationStyle", "apa")));
        var apa = Body(doc).Children.Single(c => c.Kind == "bibliography").GetProps();
        Assert.Equal("References", apa["title"]);
        Assert.StartsWith("Kuhn, T. S. (1962). The structure of scientific revolutions.", apa["text"]);
        Assert.Matches("<w:b ?/>", Body(doc).Children.Single(c => c.Kind == "bibliography").GetRaw()); // APA's title is bold

        Mutations.Set(doc.Root, Props(("source", """{"type":"webpage","title":"About the Hanford Site","container":"Hanford.gov","year":"2023"}""")));
        Assert.Equal(3, Body(doc).Children.Single(c => c.Kind == "bibliography").GetProps()["text"].Split('\n').Length);

        using var reopened = Reopen(doc);
        var again = Body(reopened).Children.Single(c => c.Kind == "bibliography");
        Mutations.Set(again, Props(("title", "Works Consulted")));
        Assert.Equal("Works Consulted", Body(reopened).Children.Single(c => c.Kind == "bibliography").GetProps()["title"]);
        Assert.Contains("Works Consulted", Views.Text(reopened.Root));
    }

    [Fact]
    public void Chicago_cites_in_footnotes_in_full_the_first_time_and_short_after()
    {
        var (doc, p) = Essay("chicago");
        Mutations.Set(doc.Root, Props(("source", KuhnJson)));
        Mutations.Add(p, "citation", Props(("sources", "Peg15"), ("at", "4")), null); // a citation in the text (author-date here) is no note: the first note is still in full
        var first = Mutations.Add(p, "footnote", Props(("cite", "Peg15"), ("pages", "288"), ("at", "13")), null);
        var second = Mutations.Add(p, "footnote", Props(("cite", "Peg15"), ("pages", "290"), ("at", "53"), ("text", "Compare the melter data.")), null);
        Assert.Equal("Ian L. Pegg, “Behavior of Technetium in Nuclear Waste Vitrification Processes,” Journal of Radioanalytical and Nuclear Chemistry 305, no. 1 (2015): 288, https://doi.org/10.1007/s10967-014-3900-9.", first.GetProps()["citeText"]);
        var s = second.GetProps();
        Assert.Equal(("Pegg, “Behavior of Technetium,” 290.", "Compare the melter data.", "Peg15", "290"), (s["citeText"], s["text"], s["cite"], s["pages"]));
        Assert.Contains("<i>Journal of Radioanalytical and Nuclear Chemistry</i>", first.GetProps()["citeHtml"]);
        AssertValid(doc);

        Mutations.Set(second, Props(("text", "See also the melter data.")));
        Assert.Equal(("Pegg, “Behavior of Technetium,” 290.", "See also the melter data."), (second.GetProps()["citeText"], second.GetProps()["text"]));

        first.Remove(); // the second note is the source's first now: in full
        var left = Body(doc).Children[0].Children.Single(c => c.Kind == "footnote").GetProps();
        Assert.StartsWith("Ian L. Pegg, “Behavior of Technetium", left["citeText"]);
        Assert.EndsWith(": 290, https://doi.org/10.1007/s10967-014-3900-9.", left["citeText"]);

        using var reopened = Reopen(doc);
        Assert.Equal("chicago", reopened.Root.GetProps()["citationStyle"]);
        var list = Mutations.Add(Body(reopened), "bibliography", Props(), null);
        Assert.Equal("Bibliography", list.GetProps()["title"]);
        Assert.StartsWith("Kuhn, Thomas S. The Structure of Scientific Revolutions. University of Chicago Press, 1962.\nPegg, Ian L.", list.GetProps()["text"]);
    }

    [Fact]
    public void A_citation_Word_made_reads_as_one_and_is_drawn_again_in_the_documents_style()
    {
        var sources = """<b:Sources SelectedStyle="\APASixthEditionOfficeOnline.xsl" StyleName="APA" Version="6" xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"><b:Source><b:Tag>Smi20</b:Tag><b:SourceType>Book</b:SourceType><b:Guid>{11111111-2222-3333-4444-555555555555}</b:Guid><b:Author><b:Author><b:NameList><b:Person><b:Last>Smith</b:Last><b:First>Jane</b:First></b:Person></b:NameList></b:Author></b:Author><b:Title>Glass and waste</b:Title><b:Year>2020</b:Year><b:Publisher>Hanford Press</b:Publisher></b:Source></b:Sources>""";
        var p = new W.Paragraph(new W.Run(new W.Text("Glass holds it ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
            new W.SdtRun(new W.SdtProperties(new W.SdtId { Val = 77 }, new W.SdtContentCitation()), new W.SdtContentRun(
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }), new W.Run(new W.FieldCode(" CITATION Smi20 \\l 1033 ")),
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }), new W.Run(new W.Text("(Smith, 2020)")),
                new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }))),
            new W.Run(new W.Text(".")));
        var bytes = Docx([p], main =>
        {
            var part = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using var w = new StreamWriter(part.GetStream(FileMode.Create));
            w.Write(sources);
        });
        using var doc = OpenDocx(bytes);
        var para = Body(doc).Children[0];
        Assert.Equal("Glass holds it .", para.GetProps()["text"]);
        var cite = para.Children.Single(c => c.Kind == "citation").GetProps();
        Assert.Equal(("77", "Smi20", "(Smith, 2020)", "15"), (cite["id"], cite["sources"], cite["text"], cite["at"]));
        Assert.Equal("apa", doc.Root.GetProps()["citationStyle"]); // what Word's style box was set to
        Mutations.Set(doc.Root, Props(("citationStyle", "mla")));
        Assert.Equal("(Smith)", Body(doc).Children[0].Children.Single(c => c.Kind == "citation").GetProps()["text"]);
        AssertValid(doc);
    }
}
