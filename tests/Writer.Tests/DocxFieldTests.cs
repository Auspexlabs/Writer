using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class DocxFieldTests
{
    static Dictionary<string, string> P(params (string, string)[] pairs) => pairs.ToDictionary(x => x.Item1, x => x.Item2);
    static Document Reopen(Document doc) { var ms = new MemoryStream(); doc.Save(ms); return OpenDocx(ms.ToArray()); }
    static void Valid(Document doc) => Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((DocxDocument)doc).Package).Select(e => e.Description));

    [Fact]
    public void Chapter_captions_reset_and_figures_directory_contains_only_its_label()
    {
        using var doc = new DocxAdapter().Create(); var body = doc.Root.Children.Single();
        Mutations.Add(body, "paragraph", P(("text", "First"), ("style", "Heading1")), null);
        var one = Mutations.Add(body, "paragraph", P(("html", "图 1 First figure"), ("caption", "图"), ("captionChapter", "1")), null);
        var two = Mutations.Add(body, "paragraph", P(("html", "图 2 Second figure"), ("caption", "图"), ("captionChapter", "1")), null);
        Mutations.Add(body, "paragraph", P(("text", "Second"), ("style", "Heading1")), null);
        var three = Mutations.Add(body, "paragraph", P(("html", "图 3 Third figure"), ("caption", "图"), ("captionChapter", "1")), null);
        Mutations.Add(body, "paragraph", P(("caption", "表"), ("html", "表 1 Table only")), null);
        var toc = Mutations.Add(body, "toc", P(("caption", "图"), ("title", "Figures")), null);
        Assert.Equal("图 1-1 First figure", one.GetProps()["text"]);
        Assert.Equal("图 1-2 Second figure", two.GetProps()["text"]);
        Assert.Equal("图 2-1 Third figure", three.GetProps()["text"]);
        Assert.Contains("TOC \\c \"图\"", toc.GetRaw());
        Assert.DoesNotContain("Table only", toc.GetRaw());
        using var reopened = Reopen(doc); Valid(reopened);
        Assert.Equal("图", PathResolver.Single(reopened.Root, "/body/toc").GetProps()["caption"]);
        Mutations.Set(three, P(("captionChapter", "none")));
        Assert.DoesNotContain("STYLEREF", three.GetRaw());
    }

    [Fact]
    public void References_and_formatted_merge_fields_survive_edits_and_reopen()
    {
        using var doc = new DocxAdapter().Create(); var body = doc.Root.Children.Single();
        var target = Mutations.Add(body, "paragraph", P(("text", "Target"), ("bookmark", "Target")), null);
        var para = Mutations.Add(body, "paragraph", P(("html", "See <span data-field=\"REF Target \\h\"><b>Tar</b>get</span>; <span data-field=\"MERGEFIELD Name\"><b>Na</b>me</span>.")), null);
        Assert.Equal(2, ((DocxDocument)doc).Main.Document!.Body!.Descendants<W.SimpleField>().Count());
        var html = para.GetProps()["html"];
        Assert.Equal(2, System.Text.RegularExpressions.Regex.Matches(html, "data-field=").Count);
        Mutations.Set(para, P(("html", "Intro " + html + " Tail")));
        Mutations.Set(target, P(("text", "Changed target")));
        Mutations.Set(doc.Root, P(("fields", "all")));
        using var reopened = Reopen(doc); Valid(reopened);
        var p = PathResolver.Single(reopened.Root, "/body/paragraph[2]");
        Assert.Equal("Intro See Changed target; Name. Tail", p.GetProps()["text"]);
        Mutations.Set(reopened.Root, P(("mergeData", "{\"Name\":\"张三\"}")));
        Assert.Equal("Intro See Changed target; 张三. Tail", p.GetProps()["text"]);
        Assert.Single(((DocxDocument)reopened).Main.Document!.Body!.Descendants<W.SimpleField>());
        Assert.Contains("<w:b", p.GetRaw()); Valid(reopened);
    }

    [Fact]
    public void Formatting_a_reference_keeps_one_field_without_nesting()
    {
        using var doc = new DocxAdapter().Create(); var body = doc.Root.Children.Single();
        var para = Mutations.Add(body, "paragraph", P(("html", "<span data-field=\"REF Target \\h\">Target</span>")), null);
        Mutations.Set(para, P(("html", "<span data-field=\"REF Target \\h\"><b>Target</b></span>")));
        Assert.Single(((DocxDocument)doc).Main.Document!.Body!.Descendants<W.SimpleField>());
        Valid(doc);
    }

    [Fact]
    public void Imported_complex_merge_fields_and_header_fields_are_resolved()
    {
        using var doc = new DocxAdapter().Create();
        var native = (DocxDocument)doc;
        native.Main.Document!.Body!.PrependChild(new W.Paragraph(
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" MERGEFIELD Name \\* MERGEFORMAT ")),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            new W.Run(new W.RunProperties(new W.Bold()), new W.Text("«Name»")),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End })));
        Mutations.Set(doc.Root, P(("header", "Header")));
        native.Main.HeaderParts.First().Header!.Append(new W.Paragraph(new W.SimpleField(new W.Run(new W.Text("Name"))) { Instruction = " MERGEFIELD Name " }));
        Mutations.Set(doc.Root, P(("mergeData", "{\"Name\":\"王五\"}")));
        using var reopened = Reopen(doc); Valid(reopened);
        Assert.Contains("王五", ((DocxDocument)reopened).Main.HeaderParts.First().Header!.InnerText);
        Assert.Contains("王五", ((DocxDocument)reopened).Main.Document!.OuterXml);
        Assert.DoesNotContain("MERGEFIELD", ((DocxDocument)reopened).Main.Document!.OuterXml);
    }
}
