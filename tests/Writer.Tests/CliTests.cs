using System.Text.Json;
using Writer.Cli;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class CliTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("writer-cli").FullName;

    public void Dispose() => Directory.Delete(_dir, true);

    string In(string name) => Path.Combine(_dir, name);

    static (int Code, string Out, string Err) Run(params string[] argv)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var code = Runner.Run(argv, stdout, stderr);
        return (code, stdout.ToString(), stderr.ToString());
    }

    static JsonElement Json(string s) => JsonDocument.Parse(s).RootElement.Clone();

    static string Code(string err) => Json(err).GetProperty("error").GetProperty("code").GetString()!;

    [Fact]
    public void Create_then_get_root()
    {
        var file = In("new.docx");
        var (code, output, _) = Run("create", file);
        Assert.Equal(0, code);
        Assert.True(File.Exists(file));
        Assert.Equal("docx", Json(output).GetProperty("format").GetString());

        var again = Run("create", file);
        Assert.Equal(3, again.Code);
        Assert.Equal("VALIDATION", Code(again.Err));

        var get = Run("get", file);
        Assert.Equal(0, get.Code);
        Assert.Equal("document", Json(get.Out).GetProperty("kind").GetString());
        Assert.Equal("/body", Json(get.Out).GetProperty("children")[0].GetProperty("path").GetString());
    }

    [Fact]
    public void Create_from_template_copies_it()
    {
        var template = In("t.docx");
        File.WriteAllBytes(template, Docx(P("Template text", "Heading1")));
        var file = In("copy.docx");
        Assert.Equal(0, Run("create", file, "--from", template).Code);
        Assert.Equal("Template text\n", Run("view", file, "text").Out);
        Assert.Equal("FILE_NOT_FOUND", Code(Run("create", In("x.docx"), "--from", In("missing.docx")).Err));
    }

    [Fact]
    public void Get_query_view_on_a_real_document()
    {
        var file = In("doc.docx");
        File.WriteAllBytes(file, Docx(P("Intro", "Heading1"), P("Revenue grew in Q4"), P("Costs fell")));

        var get = Run("get", file, "/body/paragraph[2]");
        Assert.Equal("Costs fell", Json(get.Out).GetProperty("props").GetProperty("text").GetString());

        var deep = Run("get", file, "/body", "--depth", "2");
        Assert.Equal("Intro", Json(deep.Out).GetProperty("children")[0].GetProperty("props").GetProperty("text").GetString());

        var query = Run("query", file, "//paragraph[@text~=\"q4\"]");
        var hits = Json(query.Out);
        Assert.Equal(1, hits.GetArrayLength());
        Assert.Equal("/body/paragraph[1]", hits[0].GetProperty("path").GetString());

        Assert.Contains("/body/heading[1]  level=1  \"Intro\"", Run("view", file, "outline").Out);
        Assert.Contains("/body/heading[1]", Run("view", file).Out);
        Assert.Equal("Intro\n\nRevenue grew in Q4\n\nCosts fell\n", Run("view", file, "text").Out);
        Assert.Equal("document", Json(Run("view", file, "json").Out).GetProperty("kind").GetString());
        Assert.StartsWith("<w:p ", Run("get", file, "/body/paragraph[1]", "--raw").Out);
    }

    [Fact]
    public void Errors_are_json_on_stderr_with_exit_codes()
    {
        var file = In("doc.docx");
        File.WriteAllBytes(file, Docx(P("one")));
        File.WriteAllText(In("notes.xyz"), "hi");

        var missing = Run("get", In("nope.docx"));
        Assert.Equal(2, missing.Code);
        Assert.Equal("", missing.Out);
        Assert.Equal("FILE_NOT_FOUND", Code(missing.Err));

        var badPath = Run("get", file, "/body/paragraph[9]");
        Assert.Equal(2, badPath.Code);
        Assert.Equal("PATH_NOT_FOUND", Code(badPath.Err));
        Assert.Contains("/body/paragraph[1]", Json(badPath.Err).GetProperty("error").GetProperty("hint").GetString());

        var unknownFormat = Run("get", In("notes.xyz"));
        Assert.Equal(4, unknownFormat.Code);
        Assert.Equal("UNKNOWN_FORMAT", Code(unknownFormat.Err));

        Assert.Equal(1, Run("frobnicate").Code);
        Assert.Equal("USAGE", Code(Run("frobnicate").Err));
        Assert.Equal("USAGE", Code(Run("get", file, "--bogus", "1").Err));
        Assert.Equal("USAGE", Code(Run("get").Err));
        Assert.Equal("USAGE", Code(Run("get", file, "--depth", "deep").Err));
        Assert.Equal("USAGE", Code(Run("view", file, "sideways").Err));
        Assert.Equal("USAGE", Code(Run("get", file, "body").Err));
    }

    [Fact]
    public void Add_set_remove_move_round_trip()
    {
        var file = In("edit.docx");
        Run("create", file);
        var added = Run("add", file, "/body", "--type", "heading", "--prop", "text=Intro", "--prop", "level=1");
        Assert.Equal(0, added.Code);
        Assert.Equal("/body/heading[1]", Json(added.Out).GetProperty("path").GetString());
        Run("add", file, "/body", "--type", "paragraph", "--prop", "md=Revenue **grew**");
        Run("add", file, "/body", "--type", "paragraph", "--prop", "text=Costs fell", "--after", "/body/heading[1]");
        Assert.Equal("Intro\n\nCosts fell\n\nRevenue grew\n", Run("view", file, "text").Out);

        var set = Run("set", file, "/body/paragraph[1]", "--prop", "align=center");
        Assert.Equal("center", Json(set.Out).GetProperty("props").GetProperty("align").GetString());
        var all = Run("set", file, "//paragraph", "--prop", "style=Quote", "--all");
        Assert.Equal(2, Json(all.Out).GetArrayLength());
        Assert.Equal("PATH_AMBIGUOUS", Code(Run("set", file, "//paragraph", "--prop", "style=Quote").Err));
        Assert.Equal("USAGE", Code(Run("set", file, "/body/paragraph[1]").Err));
        Assert.Equal("USAGE", Code(Run("set", file, "/body/paragraph[1]", "--prop", "novalue").Err));
        Assert.Equal("VALIDATION", Code(Run("set", file, "/body/paragraph[1]", "--prop", "colour=red").Err));
        var raw = Run("set", file, "/body/paragraph[1]", "--raw", "<w:p><w:r><w:t>Raw text</w:t></w:r></w:p>");
        Assert.Equal("Raw text", Json(raw.Out).GetProperty("props").GetProperty("text").GetString());

        var moved = Run("move", file, "/body/paragraph[2]", "--to", "/body", "--index", "1");
        Assert.Equal("/body/paragraph[1]", Json(moved.Out).GetProperty("path").GetString());
        Assert.StartsWith("Revenue grew", Run("view", file, "text").Out);

        var removed = Run("remove", file, "//paragraph", "--all");
        Assert.Equal(2, Json(removed.Out).GetProperty("removed").GetArrayLength());
        Assert.Equal("Intro\n", Run("view", file, "text").Out);
        Assert.Equal("USAGE", Code(Run("add", file, "/body", "--prop", "text=x").Err));
        Assert.Equal("UNSUPPORTED_KIND", Code(Run("add", file, "/body", "--type", "slide").Err));
        Assert.Equal("VALIDATION", Code(Run("add", file, "/body", "--type", "run", "--prop", "text=x").Err));
        Assert.Equal("VALIDATION", Code(Run("add", file, "/body", "--type", "paragraph", "--after", "/body").Err));
        Assert.Equal("USAGE", Code(Run("move", file, "/body/heading[1]").Err));
    }

    [Fact]
    public void Markdown_files_export_and_render()
    {
        var md = In("notes.md");
        File.WriteAllText(md, "# Notes\n\nSome **bold** text.\n");
        Assert.Equal("Notes\n\nSome bold text.\n", Run("view", md, "text").Out);
        Assert.Contains("<strong>bold</strong>", Run("view", md, "html").Out);

        var docx = In("notes.docx");
        var exported = Json(Run("export", md, "--to", docx).Out);
        Assert.Equal("docx", exported.GetProperty("format").GetString());
        Assert.Equal(0, exported.GetProperty("warnings").GetArrayLength());
        Assert.Equal("heading", Json(Run("get", docx, "/body/heading[1]").Out).GetProperty("kind").GetString());

        var html = In("notes.html");
        Assert.Equal(0, Run("export", docx, "--to", html).Code);
        Assert.Matches("<h1[^>]*>Notes</h1>", File.ReadAllText(html)); // the heading as its style draws it
        Assert.Equal(0, Run("export", docx, "--to", In("notes.json")).Code);
        Assert.Equal("USAGE", Code(Run("export", md).Err));
        Assert.Equal("UNKNOWN_FORMAT", Code(Run("export", md, "--to", In("notes.xyz")).Err));

        Assert.Equal(0, Run("set", md, "/body/paragraph[1]", "--prop", "md=Plain now").Code);
        Assert.Equal("# Notes\n\nPlain now\n", File.ReadAllText(md));
    }

    [Fact]
    public void Pdf_is_readable_and_exportable_but_not_writable()
    {
        var pdf = In("report.pdf");
        File.WriteAllBytes(pdf, TestDocs.Pdf());
        Assert.Contains("Quarterly Report", Run("view", pdf, "text").Out);
        Assert.Equal("2", Json(Run("get", pdf).Out).GetProperty("props").GetProperty("pages").GetString());
        var set = Run("set", pdf, "/page[1]/text[1]", "--prop", "x=1cm");
        Assert.Equal(3, set.Code);
        Assert.Equal("FORMAT_READONLY", Code(set.Err));
        Assert.Equal("FORMAT_READONLY", Code(Run("create", In("new.pdf")).Err));
        Assert.Equal("FORMAT_READONLY", Code(Run("export", pdf, "--to", In("copy.pdf")).Err));
        Assert.Equal(0, Run("export", pdf, "--to", In("report.md")).Code);
        Assert.Contains("Second page", File.ReadAllText(In("report.md")));
    }

    [Fact]
    public void App_accepts_no_browser_for_desktop_hosts()
    {
        var parsed = Args.Parse(["app", "--no-browser", "--dir", _dir]);
        Assert.True(parsed.Flag("no-browser"));
        Assert.Equal(_dir, parsed.Opt("dir"));
        Assert.False(Args.Parse(["app", "--dir", _dir]).Flag("no-browser"));
        Assert.Equal("USAGE", Code(Run("serve", "--no-browser").Err));
        Assert.Contains("app [--dir folder] [--list-depth n] [--port n] [--no-browser]", Run("help").Out);
        Assert.Equal(0, Args.Parse(["app", "--list-depth", "0"]).Int("list-depth", 3));
    }

    [Fact]
    public void Help_covers_commands_formats_and_elements()
    {
        var overview = Run("help");
        Assert.Equal(0, overview.Code);
        Assert.Contains("Usage: writer", overview.Out);
        Assert.Contains("query <file> <path>", overview.Out);
        Assert.Equal(0, Run().Code);
        Assert.Contains("docx", Json(Run("help", "--json").Out).GetProperty("formats")[0].GetString());

        var format = Run("help", "word");
        Assert.Contains("paragraph", format.Out);
        Assert.Contains("under: body, cell", format.Out);
        Assert.Equal("docx", Json(Run("help", "docx", "--json").Out).GetProperty("format").GetString());

        var kind = Run("help", "docx", "paragraph");
        Assert.Contains("none | bullet | number", kind.Out);
        Assert.Contains("read-only", kind.Out);
        Assert.Contains("writer add file.docx /body --type paragraph", kind.Out);

        Assert.DoesNotContain("read-only", Run("help", "docx", "set", "paragraph").Out);
        Assert.DoesNotContain("write-only", Run("help", "docx", "get", "paragraph").Out);

        var json = Json(Run("help", "docx", "run", "--json").Out);
        Assert.Equal("run", json.GetProperty("element").GetString());
        Assert.Contains(json.GetProperty("props").EnumerateArray(), p => p.GetProperty("name").GetString() == "bold");

        Assert.Equal("UNKNOWN_FORMAT", Code(Run("help", "wordperfect").Err));
        Assert.Equal("UNSUPPORTED_KIND", Code(Run("help", "docx", "slide").Err));
        Assert.Contains(Runner.Version, Run("--version").Out);
    }
}
