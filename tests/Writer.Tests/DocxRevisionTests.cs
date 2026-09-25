using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Cli;
using Writer.Core;
using Writer.Formats.Common;
using Writer.Formats.Docx;
using M = DocumentFormat.OpenXml.Math;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Tracked changes (w:ins / w:del, the track setting, accept and reject) and comments in Word documents.</summary>
public class DocxRevisionTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    static Document Reopen(Document doc) => OpenDocx(Save(doc));

    static byte[] Fixture() => File.ReadAllBytes(Path.Combine(FixtureDir("docx"), "revisions.docx"));

    static Node Para(Document doc, string contains) => PathResolver.Single(doc.Root, $"//paragraph[@text~=\"{contains}\"]");

    static List<(string? Text, string? Change)> Changes(Node p) =>
        p.Children.Where(c => c.Kind == "run").Select(c => (c.Text, c.GetProps().GetValueOrDefault("change"))).ToList();

    /// <summary>No schema errors in the main document, the comments parts or the settings, validated as Word 2013 (w15 done marks).</summary>
    static void AssertValid(Document doc)
    {
        var errors = new OpenXmlValidator(FileFormatVersions.Office2013).Validate(((DocxDocument)doc).Package)
            .Where(e => e.Part is MainDocumentPart or WordprocessingCommentsPart or WordprocessingCommentsExPart or DocumentSettingsPart)
            .Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    [Fact]
    public void Insertions_and_deletions_read_as_run_props_and_html_but_not_text()
    {
        using var doc = OpenDocx(Fixture());
        var inserted = Para(doc, "marked as an INSERTION");
        var run = inserted.Children.Single(c => c.Kind == "run").GetProps();
        Assert.Equal(("inserted", "Alice", "2026-05-25T18:00:00+08:00"), (run["change"], run["author"], run["date"]));
        Assert.Equal("<ins data-author=\"Alice\" data-date=\"2026-05-25T18:00:00+08:00\">This run will be marked as an INSERTION.</ins>", inserted.GetProps()["html"]);

        var mixed = Para(doc, "7a. The ");
        Assert.Equal("7a. The cat jumped and another fox ran fast. (regex tracks only the 1st 'fox'→'cat')", mixed.Text);
        Assert.StartsWith("7a. The <del data-author=\"Iris\" data-date=\"2026-05-25T18:50:00+08:00\">fox</del><ins data-author=\"Iris\"", mixed.GetProps()["html"]);
        var deleted = mixed.Children[1].GetProps();
        Assert.Equal(("fox", "deleted"), (deleted["text"], deleted["change"]));

        var props = doc.Root.GetProps();
        Assert.Equal("14", props["revisions"]);
        Assert.False(props.ContainsKey("track"));
        Assert.DoesNotContain("marked as a DELETION", Views.Text(doc.Root));
        Assert.DoesNotContain("change=", Views.Outline(doc.Root));
        Assert.DoesNotContain("/comment[", Views.Outline(doc.Root));
    }

    [Fact]
    public void Html_with_ins_and_del_writes_revision_wrappers_word_accepts()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "Keep old text")), null);
        p = Mutations.Set(p, Props(("html", "Keep <del>old</del><ins data-author=\"Ann\" data-date=\"2026-01-02T03:04:05Z\">new</ins> <b><ins>bold</ins></b> text")));
        Assert.Equal("Keep new bold text", p.Text);
        var raw = p.GetRaw();
        Assert.Contains("<w:del w:author=\"Writer\"", raw);
        Assert.Contains("<w:delText xml:space=\"preserve\">old</w:delText>", raw);
        Assert.Contains("<w:ins w:author=\"Ann\" w:date=\"2026-01-02T03:04:05Z\"", raw);
        var ids = System.Text.RegularExpressions.Regex.Matches(raw, "w:id=\"(\\d+)\"").Select(m => m.Groups[1].Value).ToList();
        Assert.Equal(3, ids.Count);
        Assert.Equal(ids.Distinct().Count(), ids.Count);
        Assert.Equal("3", doc.Root.GetProps()["revisions"]);
        AssertValid(doc);

        using var reopened = Reopen(doc);
        var again = PathResolver.Single(reopened.Root, "/body/paragraph[1]");
        Assert.Equal("Keep <del data-author=\"Writer\" data-date=\"" + again.Children[1].GetProps()["date"] + "\">old</del><ins data-author=\"Ann\" data-date=\"2026-01-02T03:04:05Z\">new</ins> <ins data-author=\"Writer\" data-date=\"" + again.Children[4].GetProps()["date"] + "\"><b>bold</b></ins> text", again.GetProps()["html"]);
        Assert.Equal("true", again.Children[4].GetProps()["bold"]);
        Assert.EndsWith("Z", again.Children[1].GetProps()["date"]);

        // The editor's own ins / del carry no author or date: saving the paragraph again leaves the stored revisions alone.
        var stored = ((W.Paragraph)again.Anchor).Descendants<W.InsertedRun>().Select(i => i.Id!.Value).ToList();
        again = Mutations.Set(again, Props(("html", "Keep <del>old</del><ins>new</ins> <b><ins>bold</ins></b> text <ins>more</ins>")));
        var ins = ((W.Paragraph)again.Anchor).Descendants<W.InsertedRun>().ToList();
        Assert.Equal(stored, ins.Take(2).Select(i => i.Id!.Value));
        Assert.Equal(("Ann", "2026-01-02T03:04:05Z"), (ins[0].Author!.Value, ins[0].Date!.InnerText));
        Assert.Equal(3, ins.Count);
    }

    [Fact]
    public void Track_records_a_word_level_diff_for_text_md_and_html()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "The quick brown fox")), null);
        var cjk = Mutations.Add(body, "paragraph", Props(("text", "这是一个测试")), null);
        var root = Mutations.Set(doc.Root, Props(("track", "true"), ("author", "Ann")));
        Assert.Equal(("true", "Ann"), (root.GetProps()["track"], root.GetProps()["author"]));
        var settings = ((DocxDocument)doc).Main.DocumentSettingsPart!.Settings!.OuterXml;
        Assert.Contains("<w:trackRevisions", settings);
        Assert.Contains("<w:docVar w:name=\"WriterAuthor\" w:val=\"Ann\" />", settings);
        Assert.Null(((DocxDocument)doc).Package.PackageProperties.Creator);

        p = Mutations.Set(p, Props(("text", "The quick red fox")));
        Assert.Equal("The quick red fox", p.Text);
        Assert.Equal([("The quick ", null), ("brown", "deleted"), ("red", "inserted"), (" fox", null)], Changes(p));
        Assert.All(p.Children.Where(c => c.GetProps().ContainsKey("change")), c => Assert.Equal("Ann", c.GetProps()["author"]));
        Assert.Equal("2", doc.Root.GetProps()["revisions"]);

        // A second edit keeps the pending revisions; bolding a word is tracked too, as the plain word deleted and the bold one inserted.
        p = Mutations.Set(p, Props(("md", "The quick red **fox** runs")));
        Assert.Equal("The quick red fox runs", p.Text);
        Assert.Equal([("The quick ", null), ("brown", "deleted"), ("red", "inserted"), (" ", null), ("fox", "deleted"), ("fox", "inserted"), (" runs", "inserted")], Changes(p));
        Assert.Equal("true", p.Children[5].GetProps()["bold"]);
        // Withdrawing one's own insertion just removes it.
        p = Mutations.Set(p, Props(("text", "The quick red fox")));
        Assert.Equal([("The quick ", null), ("brown", "deleted"), ("red", "inserted"), (" ", null), ("fox", "deleted"), ("fox", "inserted")], Changes(p));

        // CJK: every character is a word of its own.
        cjk = Mutations.Set(cjk, Props(("html", "这是一次测试")));
        Assert.Equal([("这是一", null), ("个", "deleted"), ("次", "inserted"), ("测试", null)], Changes(cjk));

        // Explicit markers in html are written as given, no diff.
        var explicitly = Mutations.Set(cjk, Props(("html", "全新<ins>内容</ins>")));
        Assert.Equal([("全新", null), ("内容", "inserted")], Changes(explicitly));
        AssertValid(doc);

        Mutations.Set(doc.Root, Props(("track", "false")));
        Assert.False(doc.Root.GetProps().ContainsKey("track"));
        var plain = Mutations.Set(explicitly, Props(("text", "plain again")));
        Assert.Single(plain.Children);

        // A paragraph added while tracking is inserted mark and all: rejecting it leaves no empty line.
        Mutations.Set(doc.Root, Props(("track", "true")));
        var added = Mutations.Add(body, "paragraph", Props(("text", "brand new")), null);
        Assert.Equal("inserted", added.Children.Single().GetProps()["change"]);
        Assert.Contains("<w:rPr><w:ins ", added.GetRaw());
        AssertValid(doc);
        Mutations.Set(doc.Root, Props(("reject", "all")));
        Assert.Equal(["The quick brown fox", "plain again"], doc.Root.Children.Single().Children.Select(b => b.Text));
        Assert.Equal("0", doc.Root.GetProps()["revisions"]);

        Mutations.Set(doc.Root, Props(("author", "")));
        Assert.False(doc.Root.GetProps().ContainsKey("author"));
        Assert.DoesNotContain("docVar", ((DocxDocument)doc).Main.DocumentSettingsPart!.Settings!.OuterXml);
    }

    /// <summary>A paragraph with a simple field, a bookmark around a word, an equation, an inline picture, a complex field, a content
    /// control and a footnote reference between its words.</summary>
    static Document Rich()
    {
        var doc = new DocxAdapter().Create();
        var png = Path.Combine(Path.GetTempPath(), $"writer-rich-{Guid.NewGuid():N}.png");
        File.WriteAllBytes(png, FakePng(4, 3));
        try { Mutations.Add(doc.Root.Children.Single(), "image", Props(("src", png)), null); }
        finally { File.Delete(png); }
        var picture = ((DocxDocument)doc).Main.Document!.Body!.Descendants<W.Drawing>().Single().Parent!;
        var holder = (W.Paragraph)picture.Parent!;
        picture.Remove();
        holder.InsertBeforeSelf(new W.Paragraph(
            T("Alpha one "),
            new W.SimpleField(T("7")) { Instruction = " PAGE " },
            T(" two "),
            new W.BookmarkStart { Id = "90", Name = "mark" }, T("three"), new W.BookmarkEnd { Id = "90" },
            T(" "),
            new M.OfficeMath(new M.Run(new M.Text("x=1"))),
            T(" four "),
            picture,
            T(" five "),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" DATE ") { Space = SpaceProcessingModeValues.Preserve }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            T("today"),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }),
            T(" omega "),
            new W.SdtRun(new W.SdtProperties(new W.SdtId { Val = 5 }), new W.SdtContentRun(T("name"))),
            new W.Run(new W.FootnoteReference { Id = 1 }),
            T(" end")));
        holder.Remove();
        ((DocxDocument)doc).Main.AddNewPart<FootnotesPart>().Footnotes = new W.Footnotes(new W.Footnote(new W.Paragraph(T("A note."))) { Id = 1 });
        return doc;
    }

    static W.Run T(string text) => new(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });

    /// <summary>The paragraph's content in order: text as is, a simple field as [F:result], a bookmark as [ ], the equation, the picture,
    /// a complex field as {code|result}, a content control as &lt;C: &gt;, a footnote reference as ^, revisions as &lt;+ &gt; and &lt;- &gt;.</summary>
    static string Signature(Node paragraph)
    {
        var sb = new StringBuilder();
        foreach (var e in ((W.Paragraph)paragraph.Anchor).ChildElements)
        {
            sb.Append(e switch
            {
                W.ParagraphProperties => "",
                W.SimpleField f => "[F:" + f.InnerText + "]",
                W.BookmarkStart => "[",
                W.BookmarkEnd => "]",
                M.OfficeMath => "[MATH]",
                W.InsertedRun ins => "<+" + ins.InnerText + ">",
                W.DeletedRun del => "<-" + del.InnerText + ">",
                W.SdtRun control => "<C:" + control.InnerText + ">",
                W.Run r when r.GetFirstChild<W.FootnoteReference>() is not null => "^",
                W.Run r when r.GetFirstChild<W.Drawing>() is not null => "[PIC]",
                W.Run r when r.GetFirstChild<W.FieldChar>()?.FieldCharType?.Value is { } type =>
                    type == W.FieldCharValues.Begin ? "{" : type == W.FieldCharValues.Separate ? "|" : "}",
                W.Run r when r.GetFirstChild<W.FieldCode>() is { } code => code.Text.Trim(),
                _ => e.InnerText,
            });
        }
        return sb.ToString();
    }

    [Fact]
    public void Editing_words_keeps_fields_equations_pictures_and_bookmarks_in_place()
    {
        const string tail = " <C:name>^ end";
        using var doc = Rich();
        AssertValid(doc);
        var p = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        Assert.Equal("Alpha one 7 two three  four  five today omega name end", p.Text);
        Assert.Equal("Alpha one [F:7] two [three] [MATH] four [PIC] five {DATE|today} omega" + tail, Signature(p));

        // Writing back what was read changes nothing at all.
        var before = p.GetRaw();
        p = Mutations.Set(p, Props(("html", p.GetProps()["html"])));
        p = Mutations.Set(p, Props(("text", p.Text!)));
        Assert.Equal(before, p.GetRaw());

        // A word before them, a word after them (html, as the editor saves), and words beside them (plain text).
        p = Mutations.Set(p, Props(("html", p.GetProps()["html"].Replace("Alpha", "Beta"))));
        Assert.Equal("Beta one [F:7] two [three] [MATH] four [PIC] five {DATE|today} omega" + tail, Signature(p));
        p = Mutations.Set(p, Props(("html", p.GetProps()["html"].Replace("omega", "<b>gamma</b>"))));
        Assert.Equal("Beta one [F:7] two [three] [MATH] four [PIC] five {DATE|today} gamma" + tail, Signature(p));
        Assert.Equal("true", p.Children.Single(c => c.Text == "gamma").GetProps()["bold"]);
        p = Mutations.Set(p, Props(("text", p.Text!.Replace("three", "drei").Replace("four", "vier"))));
        Assert.Equal("Beta one [F:7] two [drei] [MATH] vier [PIC] five {DATE|today} gamma" + tail, Signature(p));
        // Typed right after a field's result: outside the field.
        p = Mutations.Set(p, Props(("text", p.Text!.Replace("today gamma", "today, gamma"))));
        Assert.Equal("Beta one [F:7] two [drei] [MATH] vier [PIC] five {DATE|today}, gamma" + tail, Signature(p));
        // A word inside a content control stays inside it; the footnote reference keeps its place.
        p = Mutations.Set(p, Props(("text", p.Text!.Replace("name end", "Jane fin"))));
        Assert.Equal("Beta one [F:7] two [drei] [MATH] vier [PIC] five {DATE|today}, gamma <C:Jane>^ fin", Signature(p));
        AssertValid(doc);

        // Words on both sides of the equation changed in one write: each is rewritten where it was, the equation stays between.
        p = Mutations.Set(p, Props(("text", p.Text!.Replace("drei  vier", "3 4"))));
        Assert.Equal("Beta one 7 two 3 4  five today, gamma Jane fin", p.Text);
        var signature = Signature(p);
        Assert.True(signature.IndexOf('3') < signature.IndexOf("[MATH]") && signature.IndexOf("[MATH]") < signature.IndexOf('4') && signature.IndexOf('4') < signature.IndexOf("[PIC]"), signature);
        Assert.StartsWith("Beta one [F:7] two [", signature);
        Assert.EndsWith("[PIC] five {DATE|today}, gamma <C:Jane>^ fin", signature);

        // Tracked: only the word is marked, everything else stays.
        Mutations.Set(doc.Root, Props(("track", "true")));
        p = Mutations.Set(p, Props(("text", p.Text!.Replace("five", "fünf"))));
        Assert.Equal(signature.Replace("five", "<-five><+fünf>"), Signature(p));
        AssertValid(doc);
        using var reopened = OpenDocx(Save(doc));
        Assert.Equal("Beta one 7 two 3 4  fünf today, gamma Jane fin", PathResolver.Single(reopened.Root, "/body/paragraph[1]").Text);
    }

    [Fact]
    public void Text_and_a_picture_in_one_run_stay_apart_when_commented_or_tracked()
    {
        using var doc = new DocxAdapter().Create();
        var png = Path.Combine(Path.GetTempPath(), $"writer-mixed-{Guid.NewGuid():N}.png");
        File.WriteAllBytes(png, FakePng(4, 3));
        try { Mutations.Add(doc.Root.Children.Single(), "image", Props(("src", png), ("alt", "logo")), null); }
        finally { File.Delete(png); }
        var body = ((DocxDocument)doc).Main.Document!.Body!;
        var drawing = body.Descendants<W.Drawing>().Single();
        drawing.InsertBeforeSelf(new W.Text("see ") { Space = SpaceProcessingModeValues.Preserve });
        drawing.InsertAfterSelf(new W.Text(" here") { Space = SpaceProcessingModeValues.Preserve });
        var p = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        Assert.Equal("see  here", p.Text);

        Assert.Equal("here", Mutations.Add(p, "comment", Props(("text", "which?"), ("quote", "here")), null).GetProps()["quote"]);
        Assert.Single(body.Descendants<W.Drawing>());

        Mutations.Set(doc.Root, Props(("track", "true")));
        p = Mutations.Set(PathResolver.Single(doc.Root, "/body/paragraph[1]"), Props(("text", "look")));
        Assert.Equal("look", p.Text);
        Assert.Null(Assert.Single(body.Descendants<W.Drawing>()).Ancestors<W.DeletedRun>().FirstOrDefault());
        Assert.Equal(["see ", " ", "here"], body.Descendants<W.DeletedText>().Select(t => t.Text));
        Assert.Equal("look", PathResolver.Single(doc.Root, "//comment[@id=0]").GetProps()["quote"]);
        AssertValid(doc);
    }

    [Fact]
    public void Edits_inside_links_and_insertions_keep_them_whole_and_stay_valid()
    {
        using var doc = new DocxAdapter().Create();
        var p = Mutations.Add(doc.Root.Children.Single(), "paragraph",
            Props(("html", "Go <a href=\"https://example.com\">to the site</a> <ins data-author=\"Ann\" data-date=\"2026-01-01T00:00:00Z\">now or later</ins>")), null);
        var html = p.GetProps()["html"].Replace("the site", "our site").Replace("now or", "now and");
        p = Mutations.Set(p, Props(("html", html)));
        Assert.Equal("Go to our site now and later", p.Text);
        const string url = "https://example.com";
        var runs = p.Children.Where(c => c.Kind == "run")
            .Select(c => (c.Text, c.GetProps().GetValueOrDefault("link"), c.GetProps().GetValueOrDefault("change"), c.GetProps().GetValueOrDefault("author"))).ToList();
        Assert.Equal([("Go ", null, null, null), ("to our site", url, null, null), (" ", null, null, null), ("now and later", null, "inserted", "Ann")], runs);
        var body = ((DocxDocument)doc).Main.Document!.Body!;
        Assert.Single(body.Descendants<W.Hyperlink>());
        Assert.Single(((DocxDocument)doc).Main.HyperlinkRelationships);
        Assert.Single(body.Descendants<W.InsertedRun>());
        var ids = System.Text.RegularExpressions.Regex.Matches(p.GetRaw(), "w:id=\"(\\d+)\"").Select(m => m.Groups[1].Value).ToList();
        Assert.Equal(ids.Distinct().Count(), ids.Count);
        AssertValid(doc);
    }

    [Fact]
    public void Accept_and_reject_all_on_document_and_paragraph()
    {
        using var accepted = OpenDocx(Fixture());
        var root = Mutations.Set(accepted.Root, Props(("accept", "all")));
        Assert.Equal("0", root.GetProps()["revisions"]);
        Assert.Equal("7a. The cat jumped and another fox ran fast. (regex tracks only the 1st 'fox'→'cat')", Para(accepted, "7a. The ").Text);
        Assert.Equal("8a. Remove the  token here. (delete-only via find — no insertion)", Para(accepted, "8a.").Text);
        Assert.Empty(PathResolver.Query(accepted.Root, "//paragraph[@text~=\"tracked-deleted by Bob\"]"));
        Assert.Single(PathResolver.Query(accepted.Root, "//paragraph[@text~=\"inserted by Alice\"]"));
        var body = ((DocxDocument)accepted).Main.Document!.Body!;
        Assert.Empty(body.Descendants<W.InsertedRun>());
        Assert.Empty(body.Descendants<W.DeletedRun>());
        Assert.DoesNotContain(body.Descendants<W.ParagraphMarkRunProperties>().SelectMany(m => m.ChildElements), e => e is W.Inserted or W.Deleted);
        var raw = body.OuterXml;
        Assert.Contains("w:rPrChange", raw);
        Assert.Contains("w:moveFrom", raw);
        Assert.Contains("w:cellIns", raw);
        AssertValid(accepted);

        using var rejected = OpenDocx(Fixture());
        Mutations.Set(rejected.Root, Props(("reject", "all")));
        Assert.Equal("0", rejected.Root.GetProps()["revisions"]);
        Assert.Equal("7a. The fox jumped and another fox ran fast. (regex tracks only the 1st 'fox'→'cat')", Para(rejected, "7a. The ").Text);
        Assert.Equal("8a. Remove the OBSOLETE token here. (delete-only via find — no insertion)", Para(rejected, "8a.").Text);
        Assert.Empty(PathResolver.Query(rejected.Root, "//paragraph[@text~=\"inserted by Alice\"]"));
        Assert.Single(PathResolver.Query(rejected.Root, "//paragraph[@text~=\"tracked-deleted by Bob\"]"));
        Assert.DoesNotContain("w:delText", rejected.Root.Children.Single().GetRaw());

        using var one = OpenDocx(Fixture());
        var p = Mutations.Set(Para(one, "7a. The "), Props(("reject", "all")));
        Assert.Equal("7a. The fox jumped and another fox ran fast. (regex tracks only the 1st 'fox'→'cat')", p.Text);
        Assert.Equal("12", one.Root.GetProps()["revisions"]);
        Assert.Contains("<w:del ", Para(one, "8a.").GetRaw());
        using var reopened = Reopen(one);
        Assert.Equal("12", reopened.Root.GetProps()["revisions"]);
    }

    [Fact]
    public void Comments_are_created_read_updated_and_removed()
    {
        var original = Docx(P("Revenue grew 25% this year"), P("Costs were flat"));
        using var doc = OpenDocx(original);
        var second = PathResolver.Single(doc.Root, "/body/paragraph[2]");
        var whole = Mutations.Add(second, "comment", Props(("text", "Really?"), ("author", "Bob")), null);
        Assert.Equal("/body/paragraph[2]/comment[1]", whole.Path);
        var props = whole.GetProps();
        Assert.Equal(("0", "Bob", "Really?", "Costs were flat"), (props["id"], props["author"], props["text"], props["quote"]));
        Assert.EndsWith("Z", props["date"]);
        Assert.False(props.ContainsKey("resolved"));

        var first = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        var quoted = Mutations.Add(first, "comment", Props(("text", "Check this\nagainst the ledger"), ("quote", "grew 25%")), null);
        Assert.Equal("1", quoted.GetProps()["id"]);
        Assert.Equal("grew 25%", quoted.GetProps()["quote"]);
        Assert.Equal("Writer", quoted.GetProps()["author"]);
        first = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        Assert.Equal("Revenue grew 25% this year", first.Text);
        Assert.Equal(["Revenue ", "grew 25%", " this year"], first.Children.Where(c => c.Kind == "run").Select(c => c.Text));
        Assert.Equal("/body/paragraph[1]/comment[1]", first.Children.Last().Path);
        Assert.Equal("2", doc.Root.GetProps()["comments"]);
        Assert.Equal(2, PathResolver.Query(doc.Root, "//comment").Count);
        Assert.Same(quoted.Anchor, PathResolver.Single(doc.Root, "//comment[@id=1]").Anchor);
        var ex = Assert.Throws<WriterException>(() => Mutations.Add(first, "comment", Props(("text", "x"), ("quote", "missing")), null));
        Assert.Equal(ErrorCode.Validation, ex.Code);
        AssertValid(doc);

        using var reopened = Reopen(doc);
        var back = PathResolver.Single(reopened.Root, "//comment[@id=1]");
        Assert.Equal(("Check this\nagainst the ledger", "grew 25%"), (back.GetProps()["text"], back.GetProps()["quote"]));
        back = Mutations.Set(back, Props(("text", "Checked"), ("resolved", "true"), ("quote", "this year")));
        Assert.Equal(("Checked", "true", "this year"), (back.GetProps()["text"], back.GetProps()["resolved"], back.GetProps()["quote"]));
        Assert.Matches("w15:done=\"(1|true)\"", ((DocxDocument)reopened).Main.WordprocessingCommentsExPart!.CommentsEx!.OuterXml);
        AssertValid(reopened);
        using var again = Reopen(reopened);
        Assert.Equal(("Checked", "true"), (PathResolver.Single(again.Root, "//comment[@id=1]").GetProps()["text"], PathResolver.Single(again.Root, "//comment[@id=1]").GetProps()["resolved"]));
        Assert.Equal("Revenue grew 25% this year", PathResolver.Single(again.Root, "/body/paragraph[1]").Text);
        using (var copy = OpenDocx(Save(again)))
        {
            PathResolver.Single(copy.Root, "/body/paragraph[1]").Remove();
            Assert.Equal("1", copy.Root.GetProps()["comments"]);
            Assert.Equal("0", PathResolver.Single(copy.Root, "//comment").GetProps()["id"]);
            AssertValid(copy);
        }

        var all = PathResolver.Query(again.Root, "//comment");
        Assert.Equal(["/body/paragraph[1]/comment[1]", "/body/paragraph[2]/comment[1]"], all.Select(c => c.Path));
        foreach (var c in all.Reverse()) c.Remove();
        Assert.Equal("0", again.Root.GetProps()["comments"]);
        Assert.Null(((DocxDocument)again).Main.WordprocessingCommentsPart);
        Assert.Null(((DocxDocument)again).Main.WordprocessingCommentsExPart);
        var saved = Save(again);
        Assert.Null(PackageCompare.Diff(original, saved, "word/document.xml"));
        using var clean = OpenDocx(saved);
        Assert.DoesNotContain("comment", clean.Root.Children.Single().GetRaw());
        Assert.Equal(["Revenue grew 25% this year", "Costs were flat"], clean.Root.Children.Single().Children.Select(b => b.Text));
    }

    [Fact]
    public void Comment_initials_are_read_written_and_kept()
    {
        using var doc = OpenDocx(Docx(P("Revenue grew 25% this year")));
        var p = PathResolver.Single(doc.Root, "/body/paragraph[1]");
        Assert.Equal("AL", Mutations.Add(p, "comment", Props(("text", "Source?"), ("author", "Ann Lee"), ("initials", "AL")), null).GetProps()["initials"]);
        Assert.False(Mutations.Add(p, "comment", Props(("text", "ok")), null).GetProps().ContainsKey("initials"), "none unless given, as before");
        AssertValid(doc);

        using var reopened = Reopen(doc);
        var comment = PathResolver.Single(reopened.Root, "//comment[@id=0]");
        Assert.Equal("AL", comment.GetProps()["initials"]);
        Assert.Equal("A", Mutations.Set(comment, Props(("initials", "A"))).GetProps()["initials"]);
        Assert.False(Mutations.Set(comment, Props(("initials", ""))).GetProps().ContainsKey("initials"), "empty removes them");
        AssertValid(reopened);
    }

    [Fact]
    public void Comment_anchors_survive_text_replacement_and_deleted_text_is_not_quoted()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("html", "Keep <del>gone</del> this")), null);
        var c = Mutations.Add(p, "comment", Props(("text", "note")), null);
        Assert.Equal("Keep  this", c.GetProps()["quote"]);
        p = Mutations.Set(p, Props(("text", "Replaced")));
        Assert.Equal("Replaced", p.Text);
        Assert.Single(p.Children, x => x.Kind == "comment");
        Assert.Equal("Replaced", PathResolver.Single(doc.Root, "//comment[@id=0]").GetProps()["quote"]);
    }

    [Fact]
    public void Cli_adds_comments_and_resolves_revisions()
    {
        var dir = Directory.CreateTempSubdirectory("writer-rev").FullName;
        try
        {
            var file = Path.Combine(dir, "r.docx");
            File.WriteAllBytes(file, Fixture());
            var (code, output, _) = Run("add", file, "//paragraph[@text~=\"7a. The \"]", "--type", "comment", "--prop", "text=Why cat?", "--prop", "quote=cat");
            Assert.Equal(0, code);
            var node = JsonDocument.Parse(output).RootElement;
            Assert.Equal("comment", node.GetProperty("kind").GetString());
            Assert.Equal("cat", node.GetProperty("props").GetProperty("quote").GetString());
            var get = Run("get", file, "//paragraph[@text~=\"7a. The \"]", "--depth", "2");
            Assert.Contains("\"change\": \"deleted\"", get.Out);
            Assert.Contains("\"kind\": \"comment\"", get.Out);
            var set = Run("set", file, "/", "--prop", "accept=all");
            Assert.Equal(0, set.Code);
            Assert.Contains("\"revisions\": \"0\"", set.Out);
            Assert.Contains("\"comments\": \"1\"", set.Out);
            Assert.Equal(0, Run("remove", file, "//comment[@id=0]").Code);
            Assert.Contains("\"comments\": \"0\"", Run("get", file).Out);
            var help = Run("help", "docx", "comment");
            Assert.Contains("quote", help.Out);
            Assert.Contains("resolved", help.Out);
            Assert.Contains("accept", Run("help", "docx", "paragraph").Out);
            Assert.Contains("track", Run("help", "docx", "document").Out);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Inline_html_parses_and_renders_changes_and_export_drops_deletions()
    {
        var runs = InlineHtml.Parse("a <ins data-author=\"Ann\" data-date=\"2026-01-01T00:00:00Z\">b</ins><del data-author=\"\">c</del> <u>d</u>");
        Assert.Equal(
        [
            new RunSpec("a "),
            new RunSpec("b", Change: "inserted", Author: "Ann", Date: "2026-01-01T00:00:00Z"),
            new RunSpec("c", Change: "deleted"),
            new RunSpec(" "),
            new RunSpec("d", Underline: true),
        ], runs);
        Assert.Equal("a <ins data-author=\"Ann\" data-date=\"2026-01-01T00:00:00Z\">b</ins><del>c</del> <u>d</u>", InlineHtml.Render(runs));

        using var written = new DocxAdapter().Create();
        var para = Mutations.Add(written.Root.Children.Single(), "paragraph", Props(("md", "Keep <del>old</del><ins data-author=\"Bo\">**new**</ins> end")), null);
        Assert.Equal([("Keep ", null), ("old", "deleted"), ("new", "inserted"), (" end", null)], Changes(para));
        Assert.Equal(("Bo", "true"), (para.Children[2].GetProps()["author"], para.Children[2].GetProps()["bold"]));
        Mutations.Set(written.Root, Props(("header", "a <ins>b</ins><del>c</del>")));
        Assert.Equal("a <u>b</u><s>c</s>", written.Root.GetProps()["header"]);
        using var deck = new Writer.Formats.Pptx.PptxAdapter().Create();
        var slide = Mutations.Add(deck.Root, "slide", Props(("layout", "Blank")), null);
        var shape = Mutations.Add(slide, "shape", Props(("html", "a <ins>b</ins><del>c</del>")), null);
        var slideRuns = shape.Children.Single().Children;
        Assert.Equal(("true", "true"), (slideRuns.Single(r => r.Text == "b").GetProps()["underline"], slideRuns.Single(r => r.Text == "c").GetProps()["strike"]));

        using var doc = OpenDocx(Fixture());
        var (md, _) = Writer.Formats.Exporter.Export(doc, new Writer.Formats.Markdown.MdAdapter(), null);
        using (md)
        {
            var text = Views.Text(md.Root);
            Assert.Contains("7a. The cat jumped", text);
            Assert.DoesNotContain("marked as a DELETION", text);
        }
    }

    static (int Code, string Out, string Err) Run(params string[] argv)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var code = Runner.Run(argv, stdout, stderr);
        return (code, stdout.ToString(), stderr.ToString());
    }
}
