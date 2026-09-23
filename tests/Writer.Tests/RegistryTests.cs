using Writer.Core;

namespace Writer.Tests;

public class RegistryTests
{
    [Fact]
    public void Every_docx_kind_is_documented_and_wired()
    {
        var kinds = Registry.ForFormat("docx").ToList();
        Assert.Contains(kinds, k => k.Name == "paragraph");
        foreach (var k in kinds)
        {
            Assert.False(string.IsNullOrWhiteSpace(k.Description), k.Name);
            foreach (var parent in k.Parents) Assert.NotNull(Registry.Find(parent));
            foreach (var p in Registry.PropsFor(k, "docx"))
            {
                Assert.False(string.IsNullOrWhiteSpace(p.Description), $"{k.Name}.{p.Name}");
                if (p.Type == PropType.Enum) Assert.NotEmpty(p.Values!);
                if (p.ReadOnly) continue;
                Assert.False(string.IsNullOrEmpty(p.Example), $"{k.Name}.{p.Name} needs an example");
                Assert.NotEmpty(Registry.NormalizeValue(p, p.Example!));
            }
        }
    }

    [Fact]
    public void Get_rejects_kinds_the_format_lacks()
    {
        var ex = Assert.Throws<WriterException>(() => Registry.Get("docx", "slide"));
        Assert.Equal(ErrorCode.UnsupportedKind, ex.Code);
        Assert.Contains("paragraph", ex.Hint);
    }

    [Fact]
    public void Normalize_converts_and_resolves_aliases()
    {
        var props = Registry.Normalize("docx", "run", new Dictionary<string, string>
        {
            ["b"] = "yes", ["size"] = "14pt", ["color"] = "#ff0000", ["text"] = "hi",
        });
        Assert.Equal("true", props["bold"]);
        Assert.Equal("14", props["size"]);
        Assert.Equal("FF0000", props["color"]);
        Assert.Equal("hi", props["text"]);
        Assert.Equal("720000", Registry.Normalize("docx", "image", new Dictionary<string, string> { ["width"] = "2cm" })["width"]);
    }

    [Theory]
    [InlineData("paragraph", "level", "9")]
    [InlineData("paragraph", "list", "circle")]
    [InlineData("run", "bold", "maybe")]
    [InlineData("table", "data", "[oops")]
    [InlineData("image", "width", "wide")]
    [InlineData("table", "rows", "0")]
    public void Normalize_rejects_bad_values(string kind, string prop, string value)
    {
        var ex = Assert.Throws<WriterException>(() => Registry.Normalize("docx", kind, new Dictionary<string, string> { [prop] = value }));
        Assert.Equal(ErrorCode.Validation, ex.Code);
        Assert.Contains(prop, ex.Message);
    }

    [Fact]
    public void Normalize_rejects_unknown_and_read_only_props()
    {
        var unknown = Assert.Throws<WriterException>(() => Registry.Normalize("docx", "paragraph", new Dictionary<string, string> { ["colour"] = "red" }));
        Assert.Contains("text", unknown.Hint);
        var readOnly = Assert.Throws<WriterException>(() => Registry.Normalize("docx", "paragraph", new Dictionary<string, string> { ["id"] = "x" }));
        Assert.Contains("read-only", readOnly.Message);
    }

    [Fact]
    public void Enum_values_are_case_insensitive() =>
        Assert.Equal("center", Registry.Normalize("docx", "paragraph", new Dictionary<string, string> { ["align"] = "Center" })["align"]);

    [Fact]
    public void ToDisplay_formats_lengths_and_points()
    {
        var shown = Registry.ToDisplay("image", new Dictionary<string, string> { ["width"] = "720000", ["src"] = "/word/media/image1.png" });
        Assert.Equal("2cm", shown["width"]);
        Assert.Equal("/word/media/image1.png", shown["src"]);
        Assert.Equal("14pt", Registry.ToDisplay("run", new Dictionary<string, string> { ["size"] = "14" })["size"]);
        Assert.Equal("x", Registry.ToDisplay("mystery", new Dictionary<string, string> { ["k"] = "x" })["k"]);
        Assert.True(Registry.IsJson("table", "data"));
        Assert.False(Registry.IsJson("table", "rows"));
    }

    [Fact]
    public void CheckParent_enforces_where_kinds_may_go()
    {
        Registry.CheckParent("docx", "paragraph", "body");
        var ex = Assert.Throws<WriterException>(() => Registry.CheckParent("docx", "run", "body"));
        Assert.Contains("paragraph", ex.Hint);
    }
}
