using System.Text.Json;
using Writer.Cli;
using Writer.Core;
using Writer.Formats.Xlsx;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>The editing verbs beyond get/set/add/remove: structure, search, section, replace, formula, batch, and the range's fill and sort.</summary>
public class EditTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("writer-edits").FullName;

    public void Dispose() => Directory.Delete(_dir, true);

    string In(string name) => Path.Combine(_dir, name);

    static (int Code, string Out, string Err) Run(params string[] argv)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var code = Runner.Run(argv, stdout, stderr);
        return (code, stdout.ToString(), stderr.ToString());
    }

    static string Ok(params string[] argv)
    {
        var (code, output, error) = Run(argv);
        Assert.True(code == 0, error);
        return output;
    }

    static JsonElement Json(string s) => JsonDocument.Parse(s).RootElement.Clone();

    string Md(string text)
    {
        var file = In("notes.md");
        File.WriteAllText(file, text);
        return file;
    }

    const string Notes = "# Title\n\nIntro **bold** here.\n\n## Part A\n\nAlpha Q3 text.\n\n- item one\n- item two\n\n| Name | Score |\n| --- | --- |\n| Ann | 90 |\n\n## Part B\n\nBeta Q3 text.\n";

    string Book()
    {
        var file = In("book.xlsx");
        using var doc = new XlsxAdapter().Create();
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/range[A1:C5]"), new Dictionary<string, string> { ["values"] = """[["Name","Qty","Price"],["Ann",3,10],["Bob","n/a",20],["Cid",5,7],["Dee",1,100]]""" });
        Mutations.Set(PathResolver.Single(doc.Root, "/sheet[1]/cell[A5]"), new Dictionary<string, string> { ["bold"] = "true" });
        using var fs = File.Create(file);
        doc.Save(fs);
        return file;
    }

    [Fact]
    public void Structure_lists_headings_with_what_follows_them_and_a_workbook_by_sheet_with_its_header_row()
    {
        var structure = Ok("view", Md(Notes), "structure");
        Assert.Equal(
            "/body  headings=3  paragraphs=5  tables=1\n"
            + "/body/heading[1]  level=1  \"Title\"\n"
            + "  1 paragraph  /body/paragraph[1]  \"Intro bold here.\"\n"
            + "/body/heading[2]  level=2  \"Part A\"\n"
            + "  3 paragraphs  /body/paragraph[2] … /body/paragraph[4]  \"Alpha Q3 text.\"\n"
            + "  /body/table[1]  2 rows × 2 cols  \"Name | Score\"\n"
            + "/body/heading[3]  level=2  \"Part B\"\n"
            + "  1 paragraph  /body/paragraph[5]  \"Beta Q3 text.\"\n", structure);

        var book = Ok("view", Book(), "structure");
        Assert.StartsWith("/sheet[1]  name=Sheet1  range=A1:C5  (5 rows × 3 cols)\n  header (row 1): A Name | B Qty | C Price\n  row 2: Ann | 3 | 10\n", book);

        var docx = In("doc.docx");
        File.WriteAllBytes(docx, Docx(P("Report", "Heading1"), P("Body", "Quote"), P("More")));
        Assert.Contains("styles=Quote(1)", Ok("view", docx, "structure"));
        Assert.Contains("2 paragraphs  /body/paragraph[1] … /body/paragraph[2]", Ok("view", docx, "structure"));
    }

    [Fact]
    public void Search_prints_matching_blocks_with_their_paths_and_a_count()
    {
        var file = Md(Notes);
        Assert.Equal("/body/paragraph[2]  \"Alpha Q3 text.\"\n/body/paragraph[5]  \"Beta Q3 text.\"\n2 matches\n", Ok("search", file, "Q3"));
        Assert.Equal("0 matches\n", Ok("search", file, "q3"));
        Assert.StartsWith("/body/paragraph[2]", Ok("search", file, "q3", "--ignore-case"));
        Assert.Equal("/body/table[1]/row[2]/cell[1]  \"Ann\"\n1 match\n", Ok("search", file, "Ann"));
        var book = Book();
        Assert.Equal("/sheet[1]/row[3]/cell[B3]  \"n/a\"\n1 match\n", Ok("search", book, "n/a"));
    }

    [Fact]
    public void Section_reads_rewrites_and_removes_a_heading_with_its_blocks_in_markdown()
    {
        var file = Md(Notes);
        var read = Ok("section", file, "/body/heading[2]");
        Assert.StartsWith("/body/heading[2]  level=2  \"Part A\"\n/body/paragraph[2]  \"Alpha Q3 text.\"\n/body/paragraph[3]  list=bullet  level=0  \"item one\"\n", read);
        Assert.Contains("/body/table[1]  rows=2  cols=2\n    Name\tScore\n    Ann\t90\n", read);

        var rewritten = Ok("section", file, "/body/heading[2]", "--md", "## Part A2\n\nNew *body* line.\n\n1. one\n2. two");
        Assert.Equal("Now:\n/body/heading[2]  level=2  \"Part A2\"\n/body/paragraph[2]  \"New body line.\"\n/body/paragraph[3]  list=number  level=0  \"one\"\n/body/paragraph[4]  list=number  level=0  \"two\"\n", rewritten);
        Assert.Equal("# Title\n\nIntro **bold** here.\n\n## Part A2\n\nNew *body* line.\n\n1. one\n2. two\n\n## Part B\n\nBeta Q3 text.\n", File.ReadAllText(file)); // written as given, the rest byte for byte

        var removed = Json(Ok("section", file, "/body/heading[3]", "--remove"));
        Assert.Equal(["/body/heading[3]", "/body/paragraph[5]"], removed.GetProperty("removed").EnumerateArray().Select(p => p.GetString()));
        Assert.EndsWith("1. one\n2. two\n", File.ReadAllText(file));

        var notHeading = Run("section", file, "/body/paragraph[1]");
        Assert.Equal(3, notHeading.Code);
        Assert.Contains("not a heading", notHeading.Err);
    }

    [Fact]
    public void Section_rewrites_a_word_section_from_markdown_blocks_and_leaves_the_rest_alone()
    {
        var file = In("doc.docx");
        File.WriteAllBytes(file, Docx(P("Report", "Heading1"), P("Intro"), P("Costs", "Heading2"), P("Flat."), P("Detail."), P("Outlook", "Heading2"), P("Sunny.")));
        var now = Ok("section", file, "/body/heading[2]", "--md", "## Costs (revised)\n\nCosts **fell**.\n\n- one\n- two\n\n| A | B |\n| --- | --- |\n| 1 | 2 |");
        Assert.StartsWith("Now:\n/body/heading[2]  level=2  \"Costs (revised)\"\n/body/paragraph[2]  \"Costs fell.\"\n/body/paragraph[3]  style=ListParagraph  list=bullet  level=0  \"one\"\n", now);
        Assert.Contains("/body/table[1]  rows=2  cols=2  style=TableGrid\n    A\tB\n    1\t2\n", now);
        using var doc = OpenDocx(File.ReadAllBytes(file));
        var blocks = doc.Root.Children[0].Children;
        Assert.Equal(["heading", "paragraph", "heading", "paragraph", "paragraph", "paragraph", "table", "heading", "paragraph"], blocks.Select(b => b.Kind));
        Assert.Equal(["Report", "Intro", "Costs (revised)", "Costs fell.", "one", "two", null, "Outlook", "Sunny."], blocks.Select(b => b.Kind == "table" ? null : b.Text));
        Assert.Equal("Costs <b>fell</b>.", blocks[3].GetProps()["html"]);
    }

    [Fact]
    public void Replace_keeps_word_formatting_previews_a_count_and_changes_markdown_runs_and_worksheet_values()
    {
        var docx = In("doc.docx");
        File.WriteAllBytes(docx, Docx(P("Costs", "Heading1"), P("Flat in Q3.")));
        Ok("set", docx, "/body/paragraph[1]", "--prop", "html=Revenue in <b>Q3</b> grew; Q3 was strong.");
        var preview = Json(Ok("replace", docx, "--find", "Q3", "--with", "Q4", "--preview"));
        Assert.Equal((2, 1, false), (preview.GetProperty("matches").GetInt32(), preview.GetProperty("nodes").GetInt32(), preview.GetProperty("replaced").GetBoolean()));
        using (var before = OpenDocx(File.ReadAllBytes(docx))) Assert.Contains("<b>Q3</b>", before.Root.Children[0].Children[1].GetProps()["html"]); // a preview writes nothing
        var done = Json(Ok("replace", docx, "--find", "Q3", "--with", "Q4"));
        Assert.True(done.GetProperty("replaced").GetBoolean());
        Assert.Equal("/body/paragraph[1]", done.GetProperty("hits")[0].GetProperty("path").GetString());
        using (var after = OpenDocx(File.ReadAllBytes(docx))) Assert.Equal("Revenue in <b>Q4</b> grew; Q4 was strong.", after.Root.Children[0].Children[1].GetProps()["html"]);

        var md = Md("# T\n\nSay **Q3** and Q3.\n");
        Ok("replace", md, "--find", "Q3", "--with", "Q4");
        Assert.Equal("# T\n\nSay **Q4** and Q4.\n", File.ReadAllText(md));
        Ok("replace", md, "/body/heading[1]", "--find", "T", "--with", "Title");
        Assert.StartsWith("# Title\n", File.ReadAllText(md));

        var book = Book();
        Ok("formula", book, "/sheet[1]/cell[D2]", "B2*C2");
        var cells = Json(Ok("replace", book, "--find", "n/a", "--with", "0"));
        Assert.Equal(1, cells.GetProperty("matches").GetInt32());
        Assert.Equal("0", Json(Ok("get", book, "/sheet[1]/cell[B3]")).GetProperty("props").GetProperty("value").GetString());
        Assert.Equal(1, Run("replace", book, "--find", "x").Code); // --with is required unless previewing
        Assert.Equal(0, Run("replace", book, "--find", "x", "--preview").Code);
    }

    [Fact]
    public void Formula_writes_a_cell_fills_a_range_with_shifted_references_and_reports_what_it_refers_to()
    {
        var book = Book();
        var filled = Json(Ok("formula", book, "/sheet[1]/range[D2:D5]", "=B2*C2"));
        Assert.Equal(("D2:D5", "B2*C2", true), (filled.GetProperty("target").GetString(), filled.GetProperty("formula").GetString(), filled.GetProperty("written").GetBoolean()));
        Assert.Equal(("3", "10"), (filled.GetProperty("refs").GetProperty("B2").GetString(), filled.GetProperty("refs").GetProperty("C2").GetString()));
        Assert.Empty(filled.GetProperty("warnings").EnumerateArray());
        string Formula(string cell) => Json(Ok("get", book, $"/sheet[1]/cell[{cell}]")).GetProperty("props").GetProperty("formula").GetString()!;
        Assert.Equal(["B2*C2", "B3*C3", "B4*C4", "B5*C5"], new[] { "D2", "D3", "D4", "D5" }.Select(Formula));

        var pinned = Json(Ok("formula", book, "/sheet[1]/range[E2:F3]", "$B2*C$1+SUM($A$1:B1)&\"C2\""));
        Assert.Equal("$B3*D$1+SUM($A$1:C2)&\"C2\"", Formula("F3")); // the string literal stays; relative parts move by row and column

        var sum = Json(Ok("formula", book, "/sheet[1]/cell[B6]", "SUM(B2:B5)"));
        Assert.Equal("4 cells: 3 numbers, 1 text, 0 blank", sum.GetProperty("refs").GetProperty("B2:B5").GetString());
        Assert.Equal(["B2:B5 has 1 text cell (B3 'n/a'): SUM and AVERAGE skip them"], sum.GetProperty("warnings").EnumerateArray().Select(w => w.GetString()));
        Assert.Equal("SUM(B2:B5)", Formula("B6"));

        var checkedOnly = Json(Ok("formula", book, "/sheet[1]/cell[D6]", "SUMM(D2:D6", "--check"));
        Assert.False(checkedOnly.GetProperty("written").GetBoolean());
        var warnings = checkedOnly.GetProperty("warnings").EnumerateArray().Select(w => w.GetString()!).ToList();
        Assert.Contains("parentheses are not balanced", warnings);
        Assert.Contains(warnings, w => w.StartsWith("SUMM is not a function this editor computes"));
        Assert.Contains("D2:D6 overlaps the target D6: a circular reference", warnings);
        Assert.Equal("", Json(Ok("get", book, "/sheet[1]/cell[D6]")).GetProperty("props").GetProperty("value").GetString());

        Assert.Contains("no sheet named Other", Json(Ok("formula", book, "/sheet[1]/cell[G1]", "Other!A1+1", "--check")).GetProperty("warnings")[0].GetString());
        Assert.Equal(3, Run("formula", Md("# T\n"), "/body/heading[1]", "1").Code); // not a workbook
    }

    [Fact]
    public void Sort_moves_rows_with_their_formatting_and_formulas_and_leaves_the_header()
    {
        var book = Book();
        Ok("formula", book, "/sheet[1]/range[D2:D5]", "B2*C2");
        Ok("set", book, "/sheet[1]/range[A2:D5]", "--prop", "sort=C:desc");
        var values = Json(Ok("get", book, "/sheet[1]/range[A1:D5]")).GetProperty("props").GetProperty("values");
        Assert.Equal(["Name", "Dee", "Bob", "Ann", "Cid"], values.EnumerateArray().Select(r => r[0].GetString()));
        Assert.Equal(["Price", "100", "20", "10", "7"], values.EnumerateArray().Select(r => r[2].GetString()));
        var dee = Json(Ok("get", book, "/sheet[1]/cell[A2]")).GetProperty("props");
        Assert.Equal("true", dee.GetProperty("bold").GetString()); // Dee's bold name came along from row 5
        Assert.Equal("B2*C2", Json(Ok("get", book, "/sheet[1]/cell[D2]")).GetProperty("props").GetProperty("formula").GetString()); // and its formula follows its new row
        Assert.Equal("B5*C5", Json(Ok("get", book, "/sheet[1]/cell[D5]")).GetProperty("props").GetProperty("formula").GetString());

        Ok("set", book, "/sheet[1]/range[A2:D5]", "--prop", "sort=A");
        Assert.Equal(["Ann", "Bob", "Cid", "Dee"], Json(Ok("get", book, "/sheet[1]/range[A2:A5]")).GetProperty("props").GetProperty("values").EnumerateArray().Select(r => r[0].GetString()));
        Ok("set", book, "/sheet[1]/range[A2:D5]", "--prop", "sort=B"); // numbers first, then text, blanks last
        Assert.Equal(["Dee", "Ann", "Cid", "Bob"], Json(Ok("get", book, "/sheet[1]/range[A2:A5]")).GetProperty("props").GetProperty("values").EnumerateArray().Select(r => r[0].GetString()));
        var bad = Run("set", book, "/sheet[1]/range[A2:D5]", "--prop", "sort=Z");
        Assert.Equal(3, bad.Code);
        Assert.Contains("does not name a column", bad.Err);
    }

    [Fact]
    public void Batch_saves_once_after_every_command_and_nothing_when_one_fails()
    {
        var file = In("doc.docx");
        File.WriteAllBytes(file, Docx(P("One"), P("Two")));
        var failed = Run("batch", file, "--run", $"set {file} /body/paragraph[1] --prop text=Eins", "--run", $"set {file} /body/paragraph[9] --prop text=Nein");
        Assert.Equal(3, failed.Code);
        var error = Json(failed.Err).GetProperty("error");
        Assert.StartsWith("Command 2 of 2 failed, nothing was written — set ", error.GetProperty("message").GetString());
        Assert.Contains("Nothing matches /body/paragraph[9]", error.GetProperty("message").GetString());
        Assert.DoesNotContain("writer-batch-", error.GetProperty("message").GetString() + error.GetProperty("hint").GetString()); // the copy's name stays out of sight
        using (var untouched = OpenDocx(File.ReadAllBytes(file))) Assert.Equal(["One", "Two"], untouched.Root.Children[0].Children.Select(p => p.Text));

        var done = Json(Ok("batch", file, "--run", "set doc.docx /body/paragraph[1] --prop text=Eins", "--run", "add doc.docx /body --type paragraph --prop md=\"**Drei**\"", "--run", "view doc.docx text"));
        Assert.Equal(3, done.GetProperty("commands").GetInt32());
        Assert.Equal("Eins  Two  Drei", done.GetProperty("results")[2].GetProperty("output").GetString()); // outputs come back on one line each
        using (var written = OpenDocx(File.ReadAllBytes(file))) Assert.Equal(["Eins", "Two", "Drei"], written.Root.Children[0].Children.Select(p => p.Text));

        var refused = Run("batch", file, "--run", $"export {file} --to out.md");
        Assert.Equal(1, refused.Code);
        Assert.Contains("cannot run in a batch", refused.Err);
        Assert.Equal(1, Run("batch", file).Code);
    }
}
