using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class DocxFieldTests
{
    [Fact] public void Bookmarked_table_references_use_live_values_and_cross_table_dependencies()
    {
        using var doc=new DocxAdapter().Create();var body=doc.Root.Children.Single();
        var source=Mutations.Add(body,"table",P(("data","[[3,4],[5,0]]")),null);
        source.Children[1].Children[1].SetProp("html","<span data-field='=PRODUCT(R1)'>0</span>");
        var table=(W.Table)source.Anchor;
        table.Descendants<W.Paragraph>().First().PrependChild(new W.BookmarkStart{Id="90",Name="Sales"});
        table.Descendants<W.Paragraph>().Last().Append(new W.BookmarkEnd{Id="90"});
        var output=Mutations.Add(body,"table",P(("data","[[0]]")),null);
        output.Children[0].Children[0].SetProp("html","<span data-field='=SUM(Sales R1C1:R2C2)'>0</span>");
        doc.Root.SetProp("fields","all");using var reopened=Reopen(doc);Valid(reopened);
        Assert.Equal(new[]{"12","24"},((DocxDocument)reopened).Main.Document!.Descendants<W.SimpleField>().Select(f=>f.InnerText));
        source.Children[0].Children[0].SetProp("text","6");doc.Root.SetProp("fields","all");
        Assert.Equal("39",output.Children[0].Children[0].GetProps()["text"]);
    }
    static Dictionary<string, string> P(params (string, string)[] pairs) => pairs.ToDictionary(x => x.Item1, x => x.Item2);
    static Document Reopen(Document doc) { var ms = new MemoryStream(); doc.Save(ms); return OpenDocx(ms.ToArray()); }
    static void Valid(Document doc) => Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((DocxDocument)doc).Package).Select(e => e.Description));

    [Fact] public void Formulas_use_bookmark_values_and_defined_names_after_reopen()
    {
        using var doc=new DocxAdapter().Create();var body=doc.Root.Children.Single();
        Mutations.Add(body,"paragraph",P(("text","1250"),("bookmark","Total")),null);
        var table=Mutations.Add(body,"table",P(("data","[[0,0]]")),null);
        table.Children[0].Children[0].SetProp("html","<span data-field='=Total*20%'>0</span>");
        table.Children[0].Children[1].SetProp("html","<span data-field='=IF(DEFINED(Total),SUM(R),0)'>0</span>");
        doc.Root.SetProp("fields","all");using var reopened=Reopen(doc);Valid(reopened);Assert.Equal(new[]{"250","250"},((DocxDocument)reopened).Main.Document!.Descendants<W.SimpleField>().Select(f=>f.InnerText));
    }
    [Fact]
    public void Table_formulas_are_real_fields_recalculate_dependencies_and_keep_number_format_on_reopen()
    {
        using var doc = new DocxAdapter().Create(); var body=doc.Root.Children.Single();
        var table=Mutations.Add(body,"table",P(("data","[[20,3,0],[10,2,0],[0,0,0]]")),null);
        var formulas=new[]{(0,2,"=A1*B1"),(1,2,"=PRODUCT(LEFT)"),(2,0,"=SUM(ABOVE) \\# \"0.00\""),(2,2,"=SUM(C1:C2)")};
        foreach(var (r,c,f) in formulas) Mutations.Set(table.Children[r].Children[c],P(("html","<span data-field='"+f+"'>0</span>")));
        Mutations.Set(doc.Root,P(("fields","all")));
        Assert.Equal(new[]{"60","20","30.00","80"},formulas.Select(x=>table.Children[x.Item1].Children[x.Item2].GetProps()["text"]));
        Mutations.Set(table.Children[0].Children[0],P(("text","40")));Mutations.Set(doc.Root,P(("fields","all")));
        using var back=Reopen(doc);Valid(back);var native=((DocxDocument)back).Main.Document!.Body!;
        Assert.Equal(new[]{"120","20","50.00","140"},native.Descendants<W.SimpleField>().Select(f=>f.InnerText));
        Assert.Contains("SUM(ABOVE)",native.OuterXml);Assert.Contains("data-field",back.Root.Children.Single().Children.First(n=>n.Kind=="table").Children[2].Children[0].GetProps()["html"]);
    }

    [Fact]
    public void Editing_or_removing_a_formula_replaces_the_instruction_without_nesting_or_resurrection()
    {
        using var doc=new DocxAdapter().Create();var table=Mutations.Add(doc.Root.Children.Single(),"table",P(("data","[[2,3,0]]")),null);var cell=table.Children[0].Children[2];
        Mutations.Set(cell,P(("html","<span data-field='=SUM(LEFT)'>5</span>")));
        Mutations.Set(cell,P(("html","<span data-field='=PRODUCT(LEFT)'>6</span>")));
        Mutations.Set(doc.Root,P(("fields","all")));Assert.Equal("6",cell.GetProps()["text"]);
        Assert.Single(((W.TableCell)cell.Anchor).Descendants<W.SimpleField>());
        Mutations.Set(cell,P(("html","")));Mutations.Set(doc.Root,P(("fields","all")));
        Assert.Empty(((W.TableCell)cell.Anchor).Descendants<W.SimpleField>());Assert.Equal("",cell.GetProps()["text"]);Valid(doc);
    }

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
