using System.Text;
using Writer.Core;
using Writer.Formats.Markdown;

namespace Writer.Tests;

public class MdTests
{
    const string Sample = "# Title\n\nIntro with **bold**, *it*, `code` and [link](https://x.y).\n\n- one\n- two **strong**\n  - nested\n- three\n\n1. first\n2. second\n\n```python\nprint(1)\n```\n\n| Name | Score |\n|:-----|------:|\n| Ann  | 90    |\n\n![Chart](chart.png)\n\n<div>raw html</div>\n\n> quote\n\nLast paragraph.\n";

    internal static MdDocument Open(string text) => (MdDocument)new MdAdapter().Open(new MemoryStream(Encoding.UTF8.GetBytes(text)));

    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    internal static string Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    [Fact]
    public void Projects_blocks_with_flattened_lists()
    {
        using var doc = Open(Sample);
        var body = doc.Root.Children.Single();
        Assert.Equal(
            new[]
            {
                "/body/heading[1]", "/body/paragraph[1]", "/body/paragraph[2]", "/body/paragraph[3]", "/body/paragraph[4]", "/body/paragraph[5]",
                "/body/paragraph[6]", "/body/paragraph[7]", "/body/code[1]", "/body/table[1]", "/body/image[1]", "/body/paragraph[8]",
            },
            body.Children.Select(c => c.Path));
        var two = body.Children[3];
        Assert.Equal("two strong", two.Text);
        Assert.Equal(("bullet", "0"), (two.GetProps()["list"], two.GetProps()["level"]));
        Assert.Equal("1", body.Children[4].GetProps()["level"]);
        Assert.Equal("number", body.Children[6].GetProps()["list"]);
        Assert.Equal("true", two.Children[1].GetProps()["bold"]);
        var runs = body.Children[1].Children.Select(r => r.GetProps()).ToList();
        Assert.Equal("https://x.y", runs.First(r => r.ContainsKey("link"))["link"]);
        Assert.Equal("true", runs.First(r => r["text"] == "code")["code"]);
        var code = body.Children[8];
        Assert.Equal(("print(1)", "python"), (code.Text, code.GetProps()["lang"]));
        var table = body.Children[9];
        Assert.Equal("[[\"Name\",\"Score\"],[\"Ann\",\"90\"]]", table.GetProps()["data"]);
        Assert.Equal("right", PathResolver.Single(doc.Root, "/body/table[1]/row[2]/cell[2]").GetProps()["align"]);
        var image = body.Children[10];
        Assert.Equal(("chart.png", "Chart"), (image.GetProps()["src"], image.GetProps()["alt"]));
        Assert.Equal("md", image.Format);
        Assert.False(body.Children[1].GetProps().ContainsKey("list"));
    }

    [Fact]
    public void Untouched_blocks_are_written_back_byte_for_byte()
    {
        using var doc = Open(Sample);
        _ = Views.Outline(doc.Root);
        _ = NodeJson.Serialize(doc.Root, int.MaxValue);
        Assert.Equal(Sample, Save(doc));
        using var crlf = Open("A\r\n\r\nB\r\n");
        Assert.Equal("A\r\n\r\nB\r\n", Save(crlf));
        using var bom = new MdAdapter().Open(new MemoryStream([0xEF, 0xBB, 0xBF, (byte)'H', (byte)'i']));
        var ms = new MemoryStream();
        bom.Save(ms);
        Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF, (byte)'H', (byte)'i' }, ms.ToArray());
    }

    [Fact]
    public void Editing_a_paragraph_rewrites_only_that_block()
    {
        using var doc = Open(Sample);
        Mutations.Set(PathResolver.Single(doc.Root, "/body/paragraph[8]"), Props(("md", "New **last** line")));
        Assert.Equal(Sample.Replace("Last paragraph.", "New **last** line"), Save(doc));
        using var crlf = Open("A\r\n\r\nB\r\n");
        Mutations.Set(PathResolver.Single(crlf.Root, "/body/paragraph[1]"), Props(("text", "A2\nA3")));
        Assert.Equal("A2\\\r\nA3\r\n\r\nB\r\n", Save(crlf));
    }

    [Fact]
    public void Editing_a_list_item_rewrites_the_list_and_keeps_nesting()
    {
        using var doc = Open(Sample);
        Mutations.Set(PathResolver.Single(doc.Root, "/body/paragraph[3]"), Props(("text", "TWO")));
        var saved = Save(doc);
        Assert.Contains("- one\n- TWO\n  - nested\n- three\n\n1. first\n2. second\n", saved);
        Assert.Contains("<div>raw html</div>\n\n> quote\n", saved);
    }

    [Fact]
    public void Add_remove_move_blocks()
    {
        using var doc = Open("# T\n\nA\n\nB\n");
        var body = doc.Root.Children.Single();
        Mutations.Add(body, "paragraph", Props(("md", "**New** one")), 2);
        Mutations.Add(body, "code", Props(("text", "x = 1"), ("lang", "py")), null);
        Mutations.Add(body, "heading", Props(("text", "Sub"), ("level", "2")), 1);
        Assert.Equal("## Sub\n\n# T\n\n**New** one\n\nA\n\nB\n\n```py\nx = 1\n```\n", Save(doc));
        var moved = Mutations.Move(PathResolver.Single(doc.Root, "/body/heading[2]"), body, 1);
        Assert.Equal("/body/heading[1]", moved.Path);
        PathResolver.Single(doc.Root, "/body/paragraph[1]").Remove();
        Assert.Equal("# T\n\n## Sub\n\nA\n\nB\n\n```py\nx = 1\n```\n", Save(doc));
        Mutations.Add(body, "image", Props(("src", "a.png"), ("alt", "A")), 1);
        Assert.StartsWith("![A](a.png)\n\n# T\n", Save(doc));
        Assert.Null(PathResolver.Single(doc.Root, "/body/image[1]").GetBinary());
    }

    [Fact]
    public void List_edits_split_merge_and_number()
    {
        using var doc = Open("- a\n- b\n- c\n");
        var body = doc.Root.Children.Single();
        Mutations.Set(body.Children[1], Props(("list", "none")));
        Assert.Equal("- a\n\nb\n\n- c\n", Save(doc));
        Mutations.Add(body, "paragraph", Props(("text", "d"), ("list", "number")), null);
        Mutations.Add(body, "paragraph", Props(("text", "e"), ("list", "number"), ("level", "1")), null);
        Mutations.Add(body, "paragraph", Props(("text", "f"), ("list", "number")), null);
        Assert.Equal("- a\n\nb\n\n- c\n1. d\n   1. e\n2. f\n", Save(doc));
        Assert.Equal(6, body.Children.Count);
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(body.Children[1], Props(("level", "1"))));
        Assert.Contains("list", ex.Message);
        Mutations.Add(body, "paragraph", Props(("text", "1. not a list *really*")), 1);
        Assert.StartsWith("1\\. not a list \\*really\\*\n\n- a\n", Save(doc));
    }

    [Fact]
    public void Complex_lists_are_read_only_until_replaced_raw()
    {
        using var doc = Open("- item\n\n  extra paragraph\n- two\n");
        var item = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        Assert.Equal("item", item.Text);
        var ex = Assert.Throws<WriterException>(() => Mutations.Set(item, Props(("text", "x"))));
        Assert.Contains("raw", ex.Hint);
        Assert.Contains("extra paragraph", item.GetRaw());
        item.SetRaw("- simple\n- list");
        Assert.Equal("- simple\n- list\n", Save(doc));
        Assert.Equal(2, doc.Root.Children.Single().Children.Count);
    }

    [Fact]
    public void Tables_edit_and_render()
    {
        using var doc = Open("| a | b |\n|---|---|\n| 1 | 2 |\n");
        var t = Mutations.Set(PathResolver.Single(doc.Root, "/body/table[1]"), Props(("data", "[[\"x\",\"y\",\"z\"],[\"1\",\"2\",\"3\"]]")));
        Assert.Equal(("2", "3"), (t.GetProps()["rows"], t.GetProps()["cols"]));
        Mutations.Set(PathResolver.Single(doc.Root, "/body/table[1]/row[1]/cell[3]"), Props(("align", "center")));
        Mutations.Add(t, "row", Props(("data", "[\"4\",\"5\",\"6|7\"]")), null);
        Assert.Equal("| x | y | z |\n| --- | --- | :---: |\n| 1 | 2 | 3 |\n| 4 | 5 | 6\\|7 |\n", Save(doc));
        Mutations.Set(PathResolver.Single(doc.Root, "//row[1]/cell[1]"), Props(("md", "**X**")));
        Assert.Equal("true", PathResolver.Single(doc.Root, "//row[1]/cell[1]/run[1]").GetProps()["bold"]);
        PathResolver.Single(doc.Root, "//row[3]").Remove();
        Assert.Equal("2", Mutations.Refresh(t).GetProps()["rows"]);
        using var narrow = Open("| a |\n|---|\n| 1 |\n");
        Assert.Throws<WriterException>(() => PathResolver.Single(narrow.Root, "//row[1]/cell[1]").Remove());
        PathResolver.Single(narrow.Root, "//row[2]").Remove();
        Assert.Throws<WriterException>(() => PathResolver.Single(narrow.Root, "//row[1]").Remove());
    }

    [Fact]
    public void Runs_can_be_added_formatted_and_moved()
    {
        using var doc = Open("Hello\n\nOther\n");
        var p = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        var run = Mutations.Add(p, "run", Props(("text", " world"), ("bold", "true"), ("link", "https://w")), null);
        Assert.Equal("/body/paragraph[1]/run[2]", run.Path);
        Assert.Equal("Hello [**world**](https://w)\n\nOther\n", Save(doc));
        run = Mutations.Set(run, Props(("bold", "false"), ("link", ""), ("code", "true")));
        Assert.Equal("Hello `world`\n\nOther\n", Save(doc));
        Mutations.Move(run, PathResolver.Single(doc.Root, "/body/paragraph[2]"), 1);
        Assert.Equal("Hello\n\n `world`Other\n", Save(doc));
        PathResolver.Single(doc.Root, "/body/paragraph[2]/run[1]").Remove();
        Assert.Equal("Hello\n\nOther\n", Save(doc));
    }

    [Fact]
    public void Body_raw_replaces_everything()
    {
        using var doc = Open("A\n");
        doc.Root.Children.Single().SetRaw("# New\n\ntext\n");
        Assert.Equal("# New\n\ntext\n", Save(doc));
        Assert.Equal("heading", doc.Root.Children.Single().Children[0].Kind);
    }
}

public class MdFidelityTests
{
    public static TheoryData<string> Files => new(TestDocs.Fixtures("md", "*.md").Select(f => Path.GetFileName(f)));

    [Theory]
    [MemberData(nameof(Files))]
    public void Open_read_everything_save_is_byte_identical(string file)
    {
        var original = File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("md"), file));
        using var doc = new MdAdapter().Open(new MemoryStream(original));
        _ = Views.Outline(doc.Root);
        _ = Views.Text(doc.Root);
        _ = NodeJson.Serialize(doc.Root, int.MaxValue);
        var ms = new MemoryStream();
        doc.Save(ms);
        Assert.Equal(original, ms.ToArray());
    }

    [Theory]
    [MemberData(nameof(Files))]
    public void Editing_one_paragraph_leaves_the_rest_untouched(string file)
    {
        var original = File.ReadAllText(Path.Combine(TestDocs.FixtureDir("md"), file));
        var doc = (MdDocument)new MdAdapter().Open(new MemoryStream(Encoding.UTF8.GetBytes(original)));
        var target = PathResolver.Query(doc.Root, "/body/paragraph").FirstOrDefault(p => p.Text!.Length > 0 && !p.GetProps().ContainsKey("list"));
        if (target is null) return;
        var chunk = doc.ChunkOf(target.Anchor);
        var index = doc.Chunks.IndexOf(chunk);
        var start = doc.Prefix.Length + doc.Chunks.Take(index).Sum(c => c.Original!.Length + c.Gap.Length);
        var end = start + chunk.Original!.Length;
        Mutations.Set(target, Props(("text", "EDITED")));
        var saved = MdTests.Save(doc);
        Assert.StartsWith(original[..start], saved);
        Assert.EndsWith(original[end..], saved);
        Assert.Equal("EDITED", saved[start..(saved.Length - (original.Length - end))]);
    }

    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);
}
