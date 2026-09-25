using System.Text.Json;
using Writer.Core;
using Writer.Formats;

namespace Writer.Tests;

/// <summary>For every (format, kind) in the registry: add it, write every documented property, read it back, remove it.
/// New kinds join automatically once they have a recipe.</summary>
public class AdapterMatrixTests : IDisposable
{
    static readonly (string Format, string Kind, string Parent)[] Recipes =
    [
        ("docx", "heading", "/body"),
        ("docx", "paragraph", "/body"),
        ("docx", "code", "/body"),
        ("docx", "table", "/body"),
        ("docx", "image", "/body"),
        ("docx", "run", "/body/paragraph[1]"),
        ("docx", "row", "/body/table[1]"),
        ("docx", "cell", "/body/table[1]/row[1]"),
        ("docx", "pagebreak", "/body"),
        ("docx", "toc", "/body"),
        ("docx", "comment", "/body/paragraph[1]"),
        ("docx", "footnote", "/body/paragraph[1]"),
        ("docx", "equation", "/body/paragraph[1]"),
        ("docx", "shape", "/body/paragraph[1]"),
        ("md", "heading", "/body"),
        ("md", "paragraph", "/body"),
        ("md", "code", "/body"),
        ("md", "table", "/body"),
        ("md", "image", "/body"),
        ("md", "run", "/body/paragraph[1]"),
        ("md", "row", "/body/table[1]"),
        ("md", "cell", "/body/table[1]/row[1]"),
        ("pptx", "slide", "/"),
        ("pptx", "shape", "/slide[1]"),
        ("pptx", "image", "/slide[1]"),
        ("pptx", "table", "/slide[1]"),
        ("pptx", "connector", "/slide[1]"),
        ("pptx", "group", "/slide[1]"),
        ("pptx", "paragraph", "/slide[1]/shape[1]"),
        ("pptx", "run", "/slide[1]/shape[1]/paragraph[1]"),
        ("pptx", "row", "/slide[1]/table[1]"),
        ("pptx", "cell", "/slide[1]/table[1]/row[1]"),
        ("xlsx", "sheet", "/"),
        ("xlsx", "row", "/sheet[1]"),
        ("xlsx", "cell", "/sheet[1]/row[1]"),
        ("mm", "topic", "/topic[1]"),
        ("xlsx", "chart", "/sheet[1]"),
        ("xlsx", "image", "/sheet[1]"),
    ];

    public static TheoryData<string, string> Cases
    {
        get
        {
            var data = new TheoryData<string, string>();
            foreach (var r in Recipes) data.Add(r.Format, r.Kind);
            return data;
        }
    }

    readonly string _png = Path.Combine(Path.GetTempPath(), $"writer-matrix-{Guid.NewGuid():N}.png");

    public AdapterMatrixTests() => File.WriteAllBytes(_png, TestDocs.FakePng(4, 3));

    public void Dispose() => File.Delete(_png);

    [Fact]
    public void Every_addable_kind_has_a_recipe()
    {
        foreach (var kind in Registry.Kinds.Where(k => k.Parents.Length > 0))
            foreach (var format in kind.Formats.Where(f => Adapters.ForName(f).CanWrite))
                Assert.Contains(Recipes, r => r.Format == format && r.Kind == kind.Name);
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void Add_set_read_back_and_remove(string format, string kind)
    {
        var recipe = Recipes.First(r => r.Format == format && r.Kind == kind);
        using var doc = Adapters.ForName(format).Create();
        Seed(doc);
        var parent = PathResolver.Single(doc.Root, recipe.Parent);
        var before = parent.Children.Count;
        var definition = Registry.Get(format, kind);
        var props = Registry.PropsFor(definition, format).Where(p => !p.ReadOnly && !p.WriteOnly && p.Example is not null).ToList();

        var node = Mutations.Add(parent, kind, CreationProps(kind), null);
        Assert.Equal(kind, node.Kind);
        Assert.StartsWith(recipe.Parent == "/" ? "/" : recipe.Parent + "/", node.Path);
        Assert.Equal(before + 1, parent.Children.Count);

        foreach (var p in props)
        {
            var value = p.Name == "src" ? _png : p.Example!;
            node = Mutations.Set(node, new Dictionary<string, string> { [p.Name] = value });
            if (p.Name == "restart") continue; // contextual: reads true only after another list of the same kind (DocxEditTests covers it)
            var canonical = node.GetProps();
            Assert.True(canonical.ContainsKey(p.Name), $"{kind}.{p.Name} is not readable after being set");
            var expected = Registry.NormalizeValue(p, value);
            if (p.Type == PropType.Length) // docx stores twips, so lengths round trip to the displayed precision
                Assert.Equal(Units.FormatLength(long.Parse(expected)), Units.FormatLength(long.Parse(canonical[p.Name])));
            else if (format == "docx" && p.Name == "widths")
            {
                using var requested = JsonDocument.Parse(expected);
                using var stored = JsonDocument.Parse(canonical[p.Name]);
                var left = requested.RootElement.EnumerateArray().ToArray();
                var right = stored.RootElement.EnumerateArray().ToArray();
                Assert.Equal(left.Length, right.Length);
                for (var i = 0; i < left.Length; i++)
                    Assert.InRange(Math.Abs(Units.ParseLength(left[i].GetString()!) - Units.ParseLength(right[i].GetString()!)), 0L, 635L);
            }
            else if (p.Name != "src") Assert.Equal(expected, canonical[p.Name]);
        }

        node.Remove();
        Assert.Equal(before, parent.Children.Count);
    }

    Dictionary<string, string> CreationProps(string kind) => kind switch
    {
        "run" => new() { ["text"] = "seed" },
        "image" => new() { ["src"] = _png },
        "table" => new() { ["rows"] = "2", ["cols"] = "2" },
        _ => new(),
    };

    static void Seed(Document doc)
    {
        if (doc.Format is "xlsx" or "mm") return;
        if (doc.Format == "pptx")
        {
            var slide = Mutations.Add(doc.Root, "slide", new Dictionary<string, string> { ["layout"] = "Blank" }, null);
            var shape = Mutations.Add(slide, "shape", new Dictionary<string, string> { ["text"] = "seed shape" }, null);
            Mutations.Add(shape, "paragraph", new Dictionary<string, string> { ["text"] = "seed paragraph" }, null);
            Mutations.Add(slide, "table", new Dictionary<string, string> { ["rows"] = "2", ["cols"] = "2" }, null);
            return;
        }
        var body = doc.Root.Children.Single();
        Mutations.Add(body, "paragraph", new Dictionary<string, string> { ["text"] = "seed paragraph" }, null);
        Mutations.Add(body, "table", new Dictionary<string, string> { ["rows"] = "2", ["cols"] = "2" }, null);
    }
}
