using System.Text.Json;
using Writer.Core;

namespace Writer.Tests;

public class ViewsTests
{
    static FakeNode P(string kind, params (string, string)[] props) =>
        new(kind, props: props.ToDictionary(p => p.Item1, p => p.Item2));

    [Fact]
    public void Outline_lists_blocks_with_paths_and_previews()
    {
        var root = P("document", ("format", "fake")).Add(
            new FakeNode("body").Add(
                P("heading", ("text", "Intro"), ("level", "1")).Add(P("run", ("text", "Intro"))),
                P("table", ("rows", "1"), ("cols", "2"), ("data", "[[\"a\",\"b\"]]")).Add(
                    new FakeNode("row").Add(P("cell", ("text", "a"))))));
        Assert.Equal(
            "/  format=fake\n" +
            "  /body\n" +
            "    /body/heading[1]  level=1  \"Intro\"\n" +
            "    /body/table[1]  rows=1  cols=2\n",
            Views.Outline(root));
    }

    [Fact]
    public void Text_flattens_blocks_and_tables()
    {
        var root = new FakeNode("document").Add(new FakeNode("body").Add(
            P("paragraph", ("text", "One")).Add(P("run", ("text", "One"))),
            new FakeNode("table").Add(
                new FakeNode("row").Add(P("cell", ("text", "a\nb")), P("cell", ("text", "c"))),
                new FakeNode("row").Add(P("cell", ("text", "d")), P("cell", ("text", "e")))),
            P("paragraph", ("text", "")),
            P("paragraph", ("text", "Two"))));
        Assert.Equal("One\n\na b\tc\nd\te\n\nTwo\n", Views.Text(root));
    }

    [Fact]
    public void Json_has_props_children_and_raw_json_props()
    {
        var root = P("table", ("rows", "1"), ("data", "[[\"a\"]]")).Add(
            new FakeNode("row").Add(P("cell", ("text", "a"))));
        using var doc = JsonDocument.Parse(NodeJson.Serialize(root, 2));
        var r = doc.RootElement;
        Assert.Equal("table", r.GetProperty("kind").GetString());
        Assert.Equal("/", r.GetProperty("path").GetString());
        Assert.Equal(JsonValueKind.Array, r.GetProperty("props").GetProperty("data").ValueKind);
        var row = r.GetProperty("children")[0];
        Assert.Equal("/row[1]", row.GetProperty("path").GetString());
        var cell = row.GetProperty("children")[0];
        Assert.Equal("a", cell.GetProperty("text").GetString());
        Assert.False(cell.TryGetProperty("props", out _));
        Assert.False(JsonDocument.Parse(NodeJson.Serialize(root, 0)).RootElement.TryGetProperty("children", out _));
        Assert.Contains("中文", NodeJson.Serialize(P("paragraph", ("text", "中文")), 0));
        Assert.StartsWith("[", NodeJson.Summaries([root]));
        Assert.Equal("a very long…", NodeJson.Preview("a very long\nline of text", 12));
    }
}
