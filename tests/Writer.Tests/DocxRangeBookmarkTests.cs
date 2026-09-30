using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class DocxRangeBookmarkTests
{
    static Dictionary<string, string> P(string key, string value) => new() { [key] = value };

    [Fact]
    public void Partial_cross_paragraph_and_collapsed_bookmarks_round_trip_and_refresh_exact_reference()
    {
        using var doc = new DocxAdapter().Create(); var body = doc.Root.Children.Single();
        Mutations.Add(body, "paragraph", P("html", "aa<b>selected</b>"), null);
        Mutations.Add(body, "paragraph", P("text", "morezz"), null);
        Mutations.Add(body, "paragraph", P("html", "<span data-field='REF Selection \\h'>old</span>"), null);
        var patch = """{"set":[{"name":"Selection","start":"/body/paragraph[1]","startOffset":2,"end":"/body/paragraph[2]","endOffset":4},{"name":"Caret","start":"/body/paragraph[2]","startOffset":5,"end":"/body/paragraph[2]","endOffset":5}]}""";
        Mutations.Set(doc.Root, P("rangeBookmarks", patch)); Mutations.Set(doc.Root, P("fields", "all"));
        Assert.Equal("selectedmore", PathResolver.Single(doc.Root, "/body/paragraph[3]").Text);
        Assert.False(PathResolver.Single(doc.Root, "/body/paragraph[1]").GetProps().ContainsKey("bookmark"));
        var ms = new MemoryStream(); doc.Save(ms); using var back = OpenDocx(ms.ToArray());
        var ranges = JsonNode.Parse(back.Root.GetProps()["rangeBookmarks"])!.AsArray(); Assert.Equal(2, ranges.Count);
        Assert.Equal(2, ranges[0]!["startOffset"]!.GetValue<int>()); Assert.Equal(4, ranges[0]!["endOffset"]!.GetValue<int>());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((DocxDocument)back).Package));
        var before = ((DocxDocument)back).Main.Document!.OuterXml;
        Assert.Throws<WriterException>(() => Mutations.Set(back.Root, P("rangeBookmarks", patch.Replace("\"endOffset\":4", "\"endOffset\":400"))));
        Assert.Equal(before, ((DocxDocument)back).Main.Document!.OuterXml);
        Mutations.Set(back.Root, P("rangeBookmarks", """{"remove":["Selection"]}"""));
        Assert.Single(((DocxDocument)back).Main.Document!.Descendants<W.BookmarkStart>());
    }

    [Fact]
    public void Whole_paragraph_bookmark_edits_do_not_remove_nested_range_bookmarks()
    {
        using var doc = new DocxAdapter().Create(); var p = Mutations.Add(doc.Root.Children.Single(), "paragraph", P("text", "abcdef"), null);
        Mutations.Set(doc.Root, P("rangeBookmarks", """{"set":[{"name":"Partial","start":"/body/paragraph[1]","startOffset":1,"end":"/body/paragraph[1]","endOffset":3}]}"""));
        Mutations.Set(p, P("bookmark", "Whole")); Mutations.Set(p, P("bookmark", "none"));
        Assert.Contains("Partial", doc.Root.GetProps()["rangeBookmarks"]);
        Assert.Single(((DocxDocument)doc).Main.Document!.Descendants<W.BookmarkStart>());
    }
    [Fact]
    public void Cell_bookmarks_use_flattened_offsets_across_native_paragraphs()
    {
        using var doc=(DocxDocument)new DocxAdapter().Create();
        var table=Mutations.Add(doc.Root.Children.Single(),"table",P("data","[[\"first\",\"second\"]]"),null);
        var cell=table.Children[0].Children[0];((W.TableCell)cell.Anchor).Append(new W.Paragraph(new W.Run(new W.Text("third"))));
        const string path="/body/table[1]/row[1]/cell[1]";
        Mutations.Set(doc.Root,P("rangeBookmarks","{\"set\":[{\"name\":\"InCell\",\"start\":\""+path+"\",\"startOffset\":2,\"end\":\""+path+"\",\"endOffset\":9}]}"));
        var range=JsonNode.Parse(doc.Root.GetProps()["rangeBookmarks"])![0]!;
        Assert.Equal(path,range["start"]!.GetValue<string>());Assert.Equal(9,range["endOffset"]!.GetValue<int>());
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(doc.Package));
    }
}
