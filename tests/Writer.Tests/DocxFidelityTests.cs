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
}
