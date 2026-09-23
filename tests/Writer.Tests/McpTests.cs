using System.Text.Json;
using ModelContextProtocol.Protocol;
using Writer.Cli;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class McpTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("writer-mcp").FullName;

    public void Dispose() => Directory.Delete(_dir, true);

    [Theory]
    [InlineData("view a.docx outline", new[] { "view", "a.docx", "outline" })]
    [InlineData("set a.docx /body/paragraph[1] --prop text=\"Hello world\"", new[] { "set", "a.docx", "/body/paragraph[1]", "--prop", "text=Hello world" })]
    [InlineData("add a.docx /body --type paragraph --prop 'md=**bold** text'", new[] { "add", "a.docx", "/body", "--type", "paragraph", "--prop", "md=**bold** text" })]
    [InlineData("query a.docx \"//paragraph[@text~=\\\"Q4\\\"]\"", new[] { "query", "a.docx", "//paragraph[@text~=\"Q4\"]" })]
    [InlineData("  help   docx  ", new[] { "help", "docx" })]
    [InlineData("get a.docx --raw", new[] { "get", "a.docx", "--raw" })]
    public void Tokenizes_like_a_shell(string command, string[] expected) => Assert.Equal(expected, Mcp.Tokenize(command));

    [Fact]
    public void Unterminated_quotes_are_usage_errors()
    {
        var result = Mcp.Call("set a.docx /body --prop text=\"oops");
        Assert.True(result.IsError);
        Assert.Contains("USAGE", ((TextContentBlock)result.Content[0]).Text);
    }

    [Fact]
    public void Call_runs_commands_and_reports_errors()
    {
        var file = Path.Combine(_dir, "doc.docx");
        File.WriteAllBytes(file, Docx(P("Intro", "Heading1"), P("Body")));

        var outline = Mcp.Call($"writer view \"{file}\" outline");
        Assert.False(outline.IsError);
        Assert.Contains("/body/heading[1]", ((TextContentBlock)outline.Content[0]).Text);

        var set = Mcp.Call($"set \"{file}\" /body/paragraph[1] --prop text=\"Changed text\"");
        Assert.False(set.IsError);
        Assert.Equal("Changed text", JsonDocument.Parse(((TextContentBlock)set.Content[0]).Text).RootElement.GetProperty("props").GetProperty("text").GetString());

        var missing = Mcp.Call($"get \"{file}\" /body/paragraph[9]");
        Assert.True(missing.IsError);
        Assert.Contains("PATH_NOT_FOUND", ((TextContentBlock)missing.Content[0]).Text);

        var nested = Mcp.Call("mcp");
        Assert.True(nested.IsError);
    }
}
