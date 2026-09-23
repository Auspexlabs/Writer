using System.Text;
using Writer.Cli;
using Writer.Core;
using Writer.Formats;
using Writer.Formats.Docx;
using Writer.Formats.Html;
using Writer.Formats.Markdown;
using Writer.Formats.MindMap;

namespace Writer.Tests;

public class MmTests
{
    /// <summary>What FreeMind writes: a leading comment, timestamps, fonts, edges, hooks, icons, a note, and rich node text.</summary>
    const string Sample = """
        <!-- To view this file, download free mind mapping software FreeMind from http://freemind.sourceforge.net -->
        <map version="1.0.1">
        <node CREATED="1700000000000" ID="ID_1" MODIFIED="1700000000001" TEXT="Plan">
        <font BOLD="true" NAME="SansSerif" SIZE="16"/>
        <hook NAME="accessories/plugins/AutomaticLayout.properties"/>
        <node CREATED="1700000000002" ID="ID_2" MODIFIED="1700000000003" POSITION="right" TEXT="Alpha">
        <edge COLOR="#ff0000" WIDTH="thin"/>
        <icon BUILTIN="idea"/>
        <richcontent TYPE="NOTE"><html>
          <head>
          </head>
          <body>
            <p>
              alpha   note
            </p>
          </body>
        </html>
        </richcontent>
        <node CREATED="1700000000004" ID="ID_3" MODIFIED="1700000000005" TEXT="Alpha 1"/>
        </node>
        <node BACKGROUND_COLOR="#ffff00" COLOR="#0000ff" CREATED="1700000000006" FOLDED="true" ID="ID_4" LINK="https://example.com" MODIFIED="1700000000007" POSITION="left">
        <richcontent TYPE="NODE"><html><body><p>Rich <b>Beta</b></p></body></html></richcontent>
        </node>
        </node>
        </map>

        """;

    static Document Open(string text) => new MmAdapter().Open(new MemoryStream(Encoding.UTF8.GetBytes(text)));

    static Document Reopen(Document doc) => Open(MdTests.Save(doc));

    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static Node At(Document doc, string path) => PathResolver.Single(doc.Root, path);

    [Fact]
    public void Create_makes_a_map_with_one_central_topic()
    {
        using var doc = new MmAdapter().Create();
        var root = At(doc, "/topic[1]");
        Assert.Equal(("Central Topic", "mm", "document"), (root.Text, root.Format, root.Parent!.Kind));
        Assert.Matches("^<map version=\"1.0.1\">\n<node TEXT=\"Central Topic\" ID=\"ID_[0-9]+\"/>\n</map>\n$", MdTests.Save(doc));
        Assert.Equal("mm", Adapters.ForName("freemind").Format);
        Assert.Equal("mm", Adapters.ForPath("x.mm").Format);
    }

    [Fact]
    public void Children_keep_every_property_across_save_and_reopen()
    {
        using var doc = new MmAdapter().Create();
        var root = At(doc, "/topic[1]");
        var alpha = Mutations.Add(root, "topic", Props(("text", "Alpha"), ("note", "first\nsecond"), ("collapsed", "true"), ("side", "left"),
            ("link", "https://x.y"), ("color", "#ff0000"), ("fill", "yellow"), ("icon", "idea")), null);
        Assert.Equal("/topic[1]/topic[1]", alpha.Path);
        Assert.Matches(@"^ID_\d{9,10}$", alpha.GetProps()["id"]);
        var beta = Mutations.Add(root, "topic", Props(("md", "**Beta** *two*")), 1);
        Assert.Equal(("/topic[1]/topic[1]", "Beta two"), (beta.Path, beta.Text));
        Assert.Equal("Gamma&", Mutations.Add(root, "topic", Props(("html", "<b>Gamma</b>&amp;")), null).Text);

        using var again = Reopen(doc);
        Assert.Equal(new[] { "Beta two", "Alpha", "Gamma&" }, At(again, "/topic[1]").Children.Select(c => c.Text));
        var props = At(again, "/topic[1]/topic[2]").GetProps();
        Assert.Equal(("Alpha", "first\nsecond", "true", "left"), (props["text"], props["note"], props["collapsed"], props["side"]));
        Assert.Equal(("https://x.y", "FF0000", "FFFF00", "idea"), (props["link"], props["color"], props["fill"], props["icon"]));
        Assert.Equal(alpha.GetProps()["id"], props["id"]);
        Assert.NotEqual(props["id"], At(again, "/topic[1]/topic[1]").GetProps()["id"]);

        var cleared = Mutations.Set(At(again, "/topic[1]/topic[2]"), Props(("note", ""), ("collapsed", "false"), ("link", ""), ("color", "none"), ("icon", "none")));
        Assert.Equal(new[] { "text", "side", "fill", "id" }, cleared.GetProps().Keys);
        var xml = MdTests.Save(again);
        Assert.DoesNotContain("<icon", xml);
        Assert.DoesNotContain("richcontent", xml);
        Assert.DoesNotContain("FOLDED", xml);
    }

    [Fact]
    public void Topics_are_addressable_by_id_and_rich_text_reads_as_text()
    {
        using var doc = Open(Sample);
        Assert.Equal("Alpha 1", At(doc, "//topic[@id=ID_3]").Text);
        var beta = At(doc, "//topic[@id=\"ID_4\"]");
        Assert.Equal("Rich Beta", beta.Text);
        Assert.Equal(("true", "left", "https://example.com", "0000FF", "FFFF00"),
            (beta.GetProps()["collapsed"], beta.GetProps()["side"], beta.GetProps()["link"], beta.GetProps()["color"], beta.GetProps()["fill"]));
        Assert.Equal("alpha note", At(doc, "//topic[@id=ID_2]").GetProps()["note"]);
        Assert.Equal("idea", At(doc, "//topic[@id=ID_2]").GetProps()["icon"]);
        Assert.Equal(ErrorCode.PathNotFound, Assert.Throws<WriterException>(() => At(doc, "//topic[@id=ID_9]")).Code);
    }

    [Fact]
    public void Move_and_remove_honour_index_and_refuse_root_changes()
    {
        using var doc = Open(Sample);
        var root = At(doc, "/topic[1]");
        var moved = Mutations.Move(At(doc, "//topic[@id=ID_3]"), root, 1);
        Assert.Equal("/topic[1]/topic[1]", moved.Path);
        Assert.Equal(new[] { "Alpha 1", "Alpha", "Rich Beta" }, root.Children.Select(c => c.Text));
        Mutations.Move(At(doc, "//topic[@id=ID_2]"), At(doc, "//topic[@id=ID_4]"), null);
        Mutations.Move(At(doc, "//topic[@id=ID_3]"), At(doc, "//topic[@id=ID_4]"), 1);
        Assert.Equal(new[] { "Alpha 1", "Alpha" }, At(doc, "//topic[@id=ID_4]").Children.Select(c => c.Text));
        Assert.Single(root.Children);

        Assert.Equal(ErrorCode.Validation, Refused(() => Mutations.Add(doc.Root, "topic", Props(("text", "Second root")), null), "one root topic"));
        Assert.Equal(ErrorCode.Validation, Refused(() => root.Remove(), "root topic"));
        Assert.Equal(ErrorCode.Validation, Refused(() => Mutations.Move(At(doc, "//topic[@id=ID_4]"), doc.Root, null), "one root topic"));
        Assert.Equal(ErrorCode.Validation, Refused(() => Mutations.Move(At(doc, "//topic[@id=ID_4]"), At(doc, "//topic[@id=ID_2]"), null), "under itself"));
        Assert.Equal(ErrorCode.Validation, Refused(() => Mutations.Move(root, At(doc, "//topic[@id=ID_2]"), null), "root topic"));
        Assert.Equal(ErrorCode.UnsupportedKind, Refused(() => Mutations.Add(root, "paragraph", Props(("text", "x")), null), "not a mm element"));

        At(doc, "//topic[@id=ID_4]").Remove();
        Assert.Empty(root.Children);
        using var again = Reopen(doc);
        Assert.Empty(At(again, "/topic[1]").Children);
    }

    static ErrorCode Refused(Action action, string message)
    {
        var ex = Assert.Throws<WriterException>(action);
        Assert.Contains(message, ex.Message);
        return ex.Code;
    }

    [Fact]
    public void Everything_not_edited_is_written_back_byte_for_byte()
    {
        using var untouched = Open(Sample);
        _ = Views.Outline(untouched.Root);
        Assert.Equal(Sample, MdTests.Save(untouched));

        using var doc = Open(Sample);
        Mutations.Set(At(doc, "//topic[@id=ID_3]"), Props(("text", "Alpha One")));
        Assert.Equal(Sample.Replace("TEXT=\"Alpha 1\"", "TEXT=\"Alpha One\""), MdTests.Save(doc));

        Mutations.Set(At(doc, "//topic[@id=ID_4]"), Props(("text", "Beta")));
        var xml = MdTests.Save(doc);
        Assert.DoesNotContain("TYPE=\"NODE\"", xml);
        Assert.Contains("POSITION=\"left\" TEXT=\"Beta\">", xml);
        Assert.Contains("<font BOLD=\"true\" NAME=\"SansSerif\" SIZE=\"16\"/>", xml);
        Assert.Contains("<edge COLOR=\"#ff0000\" WIDTH=\"thin\"/>", xml);
        Assert.Contains("TYPE=\"NOTE\"><html>\n  <head>", xml);

        Mutations.Set(At(doc, "//topic[@id=ID_3]"), Props(("note", "todo")));
        Assert.Contains("<node CREATED=\"1700000000004\" ID=\"ID_3\" MODIFIED=\"1700000000005\" TEXT=\"Alpha One\"><richcontent TYPE=\"NOTE\"><html><head/><body><p>todo</p></body></html></richcontent></node>", MdTests.Save(doc));

        using var declared = Open("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<map version=\"1.0.1\">\n<node TEXT=\"R\"/>\n</map>\n");
        Assert.StartsWith("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<map", MdTests.Save(declared));
        Assert.Equal(ErrorCode.FormatError, Assert.Throws<WriterException>(() => Open("<html/>")).Code);
        Assert.Equal(ErrorCode.FormatError, Assert.Throws<WriterException>(() => Open("not xml")).Code);
    }

    [Fact]
    public void Raw_xml_round_trips_a_topic()
    {
        using var doc = Open(Sample);
        var topic = At(doc, "//topic[@id=ID_3]");
        Assert.Equal("<node CREATED=\"1700000000004\" ID=\"ID_3\" MODIFIED=\"1700000000005\" TEXT=\"Alpha 1\" />", topic.GetRaw());
        topic.SetRaw("<node ID=\"ID_3\" TEXT=\"Raw\"><icon BUILTIN=\"flag\"/></node>");
        Assert.Equal(("Raw", "flag"), (At(doc, "//topic[@id=ID_3]").Text, At(doc, "//topic[@id=ID_3]").GetProps()["icon"]));
        Assert.Equal(ErrorCode.Validation, Assert.Throws<WriterException>(() => topic.SetRaw("<icon/>")).Code);
    }

    [Fact]
    public void Views_show_the_map_as_an_outline()
    {
        using var doc = Open(Sample);
        Assert.Equal("Plan\n- Alpha\n  - Alpha 1\n- Rich Beta\n", Views.Text(doc.Root));
        var outline = Views.Outline(doc.Root);
        Assert.StartsWith("/  format=mm\n  /topic[1]  \"Plan\"\n    /topic[1]/topic[1]  note=alpha note  side=right  icon=idea  \"Alpha\"\n      /topic[1]/topic[1]/topic[1]  \"Alpha 1\"\n", outline);
        var html = HtmlWriter.Render(doc);
        Assert.Contains("<ul class=\"mindmap\">\n<li>Plan<ul>\n<li>Alpha<ul>\n<li>Alpha 1</li>\n</ul>\n</li>\n<li><a href=\"https://example.com\">Rich Beta</a></li>\n</ul>\n</li>\n</ul>\n", html);
    }

    [Fact]
    public void Exports_through_a_markdown_outline()
    {
        using var doc = Open(Sample);
        var (md, warnings) = Exporter.Export(doc, new MdAdapter(), null);
        Assert.Equal("# Plan\n\n- Alpha\n  - Alpha 1\n- Rich Beta\n", MdTests.Save(md));
        Assert.Equal(["1 note(s) dropped: markdown has no equivalent for topic notes"], warnings);

        var (docx, docxWarnings) = Exporter.Export(doc, new DocxAdapter(), null);
        Assert.Single(docxWarnings);
        Assert.Equal(new[] { "heading", "paragraph", "paragraph", "paragraph" }, docx.Root.Children.Single().Children.Select(c => c.Kind));
        Assert.Equal(("bullet", "1"), (At(docx, "/body/paragraph[2]").GetProps()["list"], At(docx, "/body/paragraph[2]").GetProps()["level"]));

        var (deck, _) = Exporter.Export(doc, new Writer.Formats.Pptx.PptxAdapter(), null);
        Assert.Equal("Plan", deck.Root.Children[0].GetProps()["title"]);
        Assert.Equal(ErrorCode.Validation, Assert.Throws<WriterException>(() => Exporter.Export(doc, new MmAdapter(), null)).Code);
    }

    [Fact]
    public void Headings_and_list_items_become_topics()
    {
        using var md = MdTests.Open("# Report\n\nIntro para.\n\n## Costs\n\n- rent\n  - office\n- power\n\n## Sales\n\nplain\n\n| a |\n|---|\n| 1 |\n\n#### Deep\n");
        var (map, warnings) = Exporter.Export(md, new MmAdapter(), null);
        Assert.Equal("Report\n- Costs\n  - rent\n    - office\n  - power\n- Sales\n  - Deep\n", Views.Text(map.Root));
        Assert.Equal(["/body/table[1]: table has no equivalent in mm; skipped", "2 paragraph(s) skipped: only headings and list items become topics"], warnings);
        Assert.Matches(@"^ID_\d{9,10}$", At(map, "/topic[1]/topic[1]").GetProps()["id"]);

        using var titled = new DocxAdapter().Create();
        Mutations.Set(titled.Root, Props(("title", "Deck")));
        Mutations.Add(titled.Root.Children.Single(), "heading", Props(("text", "H1"), ("level", "1")), null);
        Mutations.Add(titled.Root.Children.Single(), "heading", Props(("text", "H2"), ("level", "2")), null);
        Assert.Equal("Deck\n- H1\n  - H2\n", Views.Text(Exporter.Export(titled, new MmAdapter(), null).Target.Root));

        var longLine = new string('x', 100);
        using var prose = MdTests.Open("# Notes\n\nJust a line.\n\n" + longLine + "\n");
        var (flat, flatWarnings) = Exporter.Export(prose, new MmAdapter(), null);
        Assert.Empty(flatWarnings);
        Assert.Equal(new[] { "Just a line.", new string('x', 79) + "…" }, At(flat, "/topic[1]").Children.Select(c => c.Text));
        Assert.Equal("Notes", At(flat, "/topic[1]").Text);
    }

    [Fact]
    public void Cli_creates_adds_with_index_and_views()
    {
        var dir = Directory.CreateTempSubdirectory("writer-mm").FullName;
        try
        {
            var file = Path.Combine(dir, "map.mm");
            Assert.Equal(0, Run("create", file));
            Assert.Equal(0, Run("add", file, "/topic[1]", "--type", "topic", "--prop", "text=A", "--prop", "icon=flag"));
            Assert.Equal(0, Run("add", file, "/topic[1]", "--type", "topic", "--prop", "text=B", "--index", "1"));
            Assert.Equal(0, Run("add", file, "/topic[1]/topic[2]", "--type", "topic", "--prop", "md=**A1**"));
            var stdout = new StringWriter();
            Assert.Equal(0, Runner.Run(["view", file, "text"], stdout, new StringWriter()));
            Assert.Equal("Central Topic\n- B\n- A\n  - A1\n", stdout.ToString());
            Assert.Equal(3, Run("add", file, "/", "--type", "topic", "--prop", "text=Another root"));
            Assert.Equal(0, Run("export", file, "--to", Path.Combine(dir, "map.md")));
            Assert.Equal("# Central Topic\n\n- B\n- A\n  - A1\n", File.ReadAllText(Path.Combine(dir, "map.md")));
        }
        finally
        {
            Directory.Delete(dir, true);
        }

        static int Run(params string[] argv) => Runner.Run(argv, new StringWriter(), new StringWriter());
    }

    [Fact]
    public void Every_topic_gets_an_id_before_the_first_change_but_reading_adds_none()
    {
        using var fresh = new MmAdapter().Create();
        Assert.StartsWith("ID_", fresh.Root.Children[0].GetProps()["id"]);

        var xml = "<map version=\"1.0.1\">\n<node TEXT=\"Root\"><node TEXT=\"A\"/><node TEXT=\"B\" ID=\"ID_1\"/></node>\n</map>\n";
        using var doc = new MmAdapter().Open(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml)));
        Assert.False(doc.Root.Children[0].GetProps().ContainsKey("id"));
        var untouched = new MemoryStream();
        doc.Save(untouched);
        Assert.Equal(xml, System.Text.Encoding.UTF8.GetString(untouched.ToArray()));

        Mutations.Set(doc.Root.Children[0], new Dictionary<string, string> { ["text"] = "Root" });
        var topics = PathResolver.Query(doc.Root, "//topic");
        Assert.Equal(3, topics.Count);
        Assert.All(topics, t => Assert.StartsWith("ID_", t.GetProps()["id"]));
        Assert.Equal("ID_1", topics[2].GetProps()["id"]);
    }
}
