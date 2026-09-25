using System.Text.Json;
using Writer.Core;
using Writer.Formats.Docx;

namespace Writer.Tests;

/// <summary>The style gallery: styles.xml read as JSON with each style's own look, styles defined or changed from the same
/// fields, paragraph styles applied as pStyle and character styles as rStyle (in html, span data-style).</summary>
public class DocxStyleGalleryTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static Dictionary<string, JsonElement> Styles(Document doc) =>
        JsonDocument.Parse(doc.Root.GetProps()["styles"]).RootElement.EnumerateArray().ToDictionary(s => s.GetProperty("id").GetString()!, s => s.Clone());

    [Fact]
    public void The_gallery_lists_styles_with_their_looks_and_a_defined_style_is_applied_as_pStyle_or_rStyle()
    {
        using var doc = new DocxAdapter().Create();
        var styles = Styles(doc);
        Assert.Equal(("heading 1", 1, "true", "16"), (styles["Heading1"].GetProperty("name").GetString(), styles["Heading1"].GetProperty("heading").GetInt32(), styles["Heading1"].GetProperty("look").GetProperty("bold").GetString(), styles["Heading1"].GetProperty("look").GetProperty("size").GetString()));
        Assert.Equal(("character", "true"), (styles["Emphasis"].GetProperty("type").GetString(), styles["Emphasis"].GetProperty("look").GetProperty("italic").GetString()));
        Assert.Equal("595959", styles["Quote"].GetProperty("look").GetProperty("color").GetString());
        Assert.False(styles.ContainsKey("TableGrid"), "table styles are not in the paragraph gallery");

        Mutations.Set(doc.Root, Props(("style", """{"id":"Note","name":"备注","italic":true,"color":"595959","spaceAfter":"12pt","lineSpacing":"1.5","font":"Kaiti"}""")));
        var note = Styles(doc)["Note"];
        Assert.Equal(("备注", "paragraph", "Normal"), (note.GetProperty("name").GetString(), note.GetProperty("type").GetString(), note.GetProperty("basedOn").GetString()));
        Assert.Equal(("true", "595959", "12pt", "1.5", "Kaiti"), (note.GetProperty("look").GetProperty("italic").GetString(), note.GetProperty("look").GetProperty("color").GetString(),
            note.GetProperty("look").GetProperty("spaceAfter").GetString(), note.GetProperty("look").GetProperty("lineSpacing").GetString(), note.GetProperty("look").GetProperty("font").GetString()));

        Mutations.Set(doc.Root, Props(("style", """{"id":"Note","bold":true,"italic":"none","spaceAfter":"none","size":"10.5"}"""))); // a change keeps what it does not name
        note = Styles(doc)["Note"];
        Assert.Equal(("true", "595959", "10.5"), (note.GetProperty("look").GetProperty("bold").GetString(), note.GetProperty("look").GetProperty("color").GetString(), note.GetProperty("look").GetProperty("size").GetString()));
        Assert.False(note.GetProperty("look").TryGetProperty("italic", out _));
        Assert.False(note.GetProperty("look").TryGetProperty("spaceAfter", out _));

        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("html", "plain <span data-style=\"Emphasis\">stressed</span> <b>bold</b>"), ("style", "Note")), null);
        Assert.Equal("Note", p.GetProps()["style"]);
        Assert.Equal("plain <span data-style=\"Emphasis\">stressed</span> <b>bold</b>", p.GetProps()["html"]);
        var run = p.Children.Single(r => r.Kind == "run" && r.Text == "stressed");
        Assert.Equal("Emphasis", run.GetProps()["style"]);
        Assert.Contains("w:rStyle w:val=\"Emphasis\"", run.GetRaw());
        Assert.False(run.GetProps().ContainsKey("italic"), "the style gives the slant, not direct formatting");

        Mutations.Set(run, Props(("style", "Strong")));
        Assert.Equal("Strong", p.Children.Single(r => r.Kind == "run" && r.Text == "stressed").GetProps()["style"]);
        Mutations.Set(p, Props(("html", "plain stressed <b>bold</b>"))); // the editor took the span off
        Assert.DoesNotContain("rStyle", p.GetRaw());
        Assert.Throws<WriterException>(() => Mutations.Set(doc.Root, Props(("style", """{"bold":true}"""))));
    }
}
