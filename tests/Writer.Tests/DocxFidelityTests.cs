using Writer.Core;
using Writer.Formats.Docx;

namespace Writer.Tests;

/// <summary>The promise behind the whole design: opening, reading everything and saving changes nothing.</summary>
public class DocxFidelityTests
{
    public static TheoryData<string> Files => new(TestDocs.Fixtures("docx", "*.docx").Select(f => Path.GetFileName(f)));

    [Theory]
    [MemberData(nameof(Files))]
    public void Open_read_everything_save_changes_nothing(string file)
    {
        var original = File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("docx"), file));
        using var doc = new DocxAdapter().Open(new MemoryStream(original));
        _ = Views.Outline(doc.Root);
        _ = Views.Text(doc.Root);
        _ = NodeJson.Serialize(doc.Root, int.MaxValue);
        using var saved = new MemoryStream();
        doc.Save(saved);
        Assert.Null(PackageCompare.Diff(original, saved.ToArray()));
    }

    /// <summary>What a save from the editor does to paragraphs the user retyped: every word changed, then set back. The runs cut apart on
    /// the way join again, so the file is as it was (no split or empty runs piling up with each save).</summary>
    [Theory]
    [MemberData(nameof(Files))]
    public void Retyping_every_paragraph_and_undoing_it_leaves_the_file_as_it_was(string file)
    {
        var original = File.ReadAllBytes(Path.Combine(TestDocs.FixtureDir("docx"), file));
        using var doc = new DocxAdapter().Open(new MemoryStream(original));
        foreach (var path in PathResolver.Query(doc.Root, "//paragraph").Concat(PathResolver.Query(doc.Root, "//heading")).Select(n => n.Path).ToList())
        {
            var node = PathResolver.Single(doc.Root, path);
            if (node.GetProps().GetValueOrDefault("html") is not { Length: > 0 } html) continue;
            var altered = string.Concat(System.Text.RegularExpressions.Regex.Split(html, "(<[^>]*>)").Select(s => s.StartsWith('<') ? s : System.Text.RegularExpressions.Regex.Replace(s, @"(\S+)", "$1~")));
            Mutations.Set(node, new Dictionary<string, string> { ["html"] = altered });
            Mutations.Set(PathResolver.Single(doc.Root, path), new Dictionary<string, string> { ["html"] = html });
        }
        using var saved = new MemoryStream();
        doc.Save(saved);
        Assert.Null(PackageCompare.Diff(original, saved.ToArray()));
    }
}
