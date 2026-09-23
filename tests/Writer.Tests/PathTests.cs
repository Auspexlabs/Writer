using Writer.Core;

namespace Writer.Tests;

public class PathParserTests
{
    [Fact]
    public void Parses_child_steps_with_indexes()
    {
        var segs = PathParser.Parse("/body/paragraph[2]");
        Assert.Equal(2, segs.Count);
        Assert.False(segs[0].Deep);
        Assert.Equal("body", segs[0].Kind);
        Assert.Empty(segs[0].Selectors);
        Assert.Equal(new IndexSelector(2), segs[1].Selectors.Single());
    }

    [Fact]
    public void Parses_deep_negative_key_attr_and_quoted()
    {
        var segs = PathParser.Parse("//table[-1]/row[3]/cell[B3]");
        Assert.True(segs[0].Deep);
        Assert.Equal(new IndexSelector(-1), segs[0].Selectors.Single());
        Assert.Equal(new IndexSelector(3), segs[1].Selectors.Single());
        Assert.Equal(new KeySelector("B3"), segs[2].Selectors.Single());

        var attr = PathParser.Parse("//paragraph[@text~=\"Q4 [draft]\"][2]");
        Assert.Equal(new AttrSelector("text", true, "Q4 [draft]"), attr[0].Selectors[0]);
        Assert.Equal(new IndexSelector(2), attr[0].Selectors[1]);
        Assert.Equal(new AttrSelector("style", false, "Heading1"), PathParser.Parse("/body/paragraph[@style=Heading1]")[1].Selectors.Single());
        Assert.Equal(new KeySelector("3"), PathParser.Parse("/sheet[\"3\"]")[0].Selectors.Single());
        Assert.Equal(new KeySelector("Q1 Sales"), PathParser.Parse("/sheet['Q1 Sales']")[0].Selectors.Single());
    }

    [Fact]
    public void Root_and_wildcard()
    {
        Assert.Empty(PathParser.Parse("/"));
        Assert.Equal("*", PathParser.Parse("/body/*[3]")[1].Kind);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("/body/")]
    [InlineData("//")]
    [InlineData("/body/paragraph[0]")]
    [InlineData("/body/paragraph[")]
    [InlineData("/body/paragraph[@text]")]
    [InlineData("/body/paragraph[]")]
    [InlineData("/body/paragraph[1]x")]
    [InlineData("/body/paragraph[\"open]")]
    public void Rejects_malformed_paths(string path)
    {
        var ex = Assert.Throws<WriterException>(() => PathParser.Parse(path));
        Assert.Equal(ErrorCode.Usage, ex.Code);
        Assert.Contains("//heading[3]", ex.Hint);
    }
}

public class PathResolverTests
{
    static FakeNode N(string kind, string? text = null) =>
        new(kind, props: text is null ? null : new Dictionary<string, string> { ["text"] = text });

    static FakeNode Doc() =>
        N("document").Add(
            N("body").Add(
                N("heading", "Intro"),
                N("paragraph", "Revenue grew in Q4"),
                N("paragraph", "Costs fell"),
                N("table").Add(
                    N("row").Add(N("cell", "a"), N("cell", "b")),
                    N("row").Add(N("cell", "c"), N("cell", "Q4 total"))),
                N("heading", "Outlook"),
                N("paragraph", "Q4 was strong")));

    [Fact]
    public void Assigns_paths_by_kind_and_position()
    {
        var body = Doc().Children.Single();
        Assert.Equal("/body", body.Path);
        Assert.Equal(
            new[] { "/body/heading[1]", "/body/paragraph[1]", "/body/paragraph[2]", "/body/table[1]", "/body/heading[2]", "/body/paragraph[3]" },
            body.Children.Select(c => c.Path));
        Assert.Equal("/body/table[1]/row[2]/cell[2]", PathResolver.Single(Doc(), "/body/table[1]/row[2]/cell[2]").Path);
    }

    [Fact]
    public void Child_axis_indexes_count_within_the_parent()
    {
        Assert.Equal("Costs fell", PathResolver.Single(Doc(), "/body/paragraph[2]").Text);
        Assert.Equal("Q4 was strong", PathResolver.Single(Doc(), "/body/paragraph[-1]").Text);
        Assert.Equal("Costs fell", PathResolver.Single(Doc(), "/body/*[3]").Text);
    }

    [Fact]
    public void Deep_axis_counts_across_the_document()
    {
        Assert.Equal("Outlook", PathResolver.Single(Doc(), "//heading[2]").Text);
        Assert.Equal("Q4 total", PathResolver.Single(Doc(), "//cell[-1]").Text);
        Assert.Equal(4, PathResolver.Query(Doc(), "//cell").Count);
        Assert.Single(PathResolver.Query(Doc(), "//table//cell[@text~=\"q4\"]"));
    }

    [Fact]
    public void Attribute_selectors_match_display_values()
    {
        Assert.Equal(2, PathResolver.Query(Doc(), "//paragraph[@text~=\"Q4\"]").Count);
        Assert.Equal("/body/paragraph[2]", PathResolver.Single(Doc(), "//paragraph[@text=\"Costs fell\"]").Path);
        Assert.Empty(PathResolver.Query(Doc(), "//paragraph[@text=\"costs fell\"]"));
        Assert.Equal("/body/paragraph[3]", PathResolver.Single(Doc(), "//paragraph[@text~=\"Q4\"][2]").Path);
    }

    [Fact]
    public void Keys_names_and_virtual_children()
    {
        var sheet = new FakeNode("sheet", name: "Sales").Add(
            new FakeNode("row", key: "3").Add(new FakeNode("cell", key: "B3", props: new() { ["text"] = "42" })));
        sheet.Virtual = (kind, key) => kind == "cell" ? new FakeNode("cell", key: key.ToUpperInvariant()) : null;
        var root = N("document").Add(sheet);

        Assert.Equal("/sheet[1]/row[3]/cell[B3]", PathResolver.Single(root, "/sheet[Sales]/row[3]/cell[b3]").Path);
        Assert.Equal("42", PathResolver.Single(root, "/sheet[1]/row[3]/cell[B3]").Text);
        var virt = PathResolver.Single(root, "/sheet[1]/cell[c9]");
        Assert.Equal("/sheet[1]/cell[C9]", virt.Path);
        Assert.Same(sheet, virt.Parent);
        Assert.Empty(PathResolver.Query(root, "/sheet[1]/row[7]"));
    }

    [Fact]
    public void Single_explains_misses_and_ambiguity()
    {
        var miss = Assert.Throws<WriterException>(() => PathResolver.Single(Doc(), "/body/paragraph[9]"));
        Assert.Equal(ErrorCode.PathNotFound, miss.Code);
        Assert.Contains("/body/paragraph[3]", miss.Hint);

        var deep = Assert.Throws<WriterException>(() => PathResolver.Single(Doc(), "/body/table[1]/row[1]/cell[5]"));
        Assert.Contains("/body/table[1]/row[1]/cell[2]", deep.Hint);

        var many = Assert.Throws<WriterException>(() => PathResolver.Single(Doc(), "//paragraph"));
        Assert.Equal(ErrorCode.PathAmbiguous, many.Code);
        Assert.Contains("3 nodes", many.Message);

        Assert.Equal("document", PathResolver.Single(Doc(), "/").Kind);
    }
}
