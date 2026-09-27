using System.Text.Json;
using Writer.Core;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class DocxCommentRangeTests
{
    [Fact]
    public void Cross_paragraph_anchor_roundtrips_and_invalid_range_keeps_old_anchor()
    {
        using var doc = OpenDocx(Docx(P("first paragraph"), P("middle paragraph"), P("last paragraph")));
        var range = """{"start":"/body/paragraph[1]","startOffset":6,"end":"/body/paragraph[3]","endOffset":4}""";
        var comment = Mutations.Add(PathResolver.Single(doc.Root, "/body/paragraph[3]"), "comment", new Dictionary<string, string> { ["text"] = "note", ["range"] = range }, null);
        Assert.Equal("paragraphmiddle paragraphlast", comment.GetProps()["quote"]);
        Assert.Equal(range, comment.GetProps()["range"]);
        var original = doc.Root.Children.Single().GetRaw();
        Assert.Throws<WriterException>(() => Mutations.Set(comment, new Dictionary<string, string> { ["range"] = range.Replace("\"endOffset\":4", "\"endOffset\":999") }));
        Assert.Equal(original, doc.Root.Children.Single().GetRaw());
        using var bytes = new MemoryStream(); doc.Save(bytes);
        using var reopened = OpenDocx(bytes.ToArray());
        Assert.Equal(range, PathResolver.Single(reopened.Root, "//comment").GetProps()["range"]);
        PathResolver.Single(reopened.Root, "//comment").Remove();
        Assert.Equal(new[] { "first paragraph", "middle paragraph", "last paragraph" }, PathResolver.Query(reopened.Root, "//paragraph").Select(n => n.Text));
    }
}
