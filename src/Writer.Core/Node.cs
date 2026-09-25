using System.Globalization;

namespace Writer.Core;

/// <summary>One file format: how to create and open documents of that format.</summary>
public interface IFormatAdapter
{
    string Format { get; }
    IReadOnlyList<string> Extensions { get; }
    bool CanWrite { get; }
    Document Create();
    Document Open(Stream stream);
}

/// <summary>An open document. Save writes the native tree and nothing else.</summary>
public abstract class Document : IDisposable
{
    public abstract string Format { get; }
    public abstract Node Root { get; }
    /// <summary>Where the document was loaded from, when known. Relative references (markdown images) resolve against it.</summary>
    public string? SourcePath { get; set; }
    public abstract void Save(Stream stream);
    public virtual void Dispose() { }
}

/// <summary>
/// A node of the unified tree. Nodes are projections over a native object (<see cref="Anchor"/>) and are
/// re-created on every access; they carry no state of their own besides their position.
/// </summary>
public abstract class Node
{
    public Node? Parent { get; private set; }
    public string Path { get; private set; } = "/";

    public abstract string Kind { get; }
    /// <summary>The native object this node projects. Two nodes are the same node when their anchors are the same object.</summary>
    public abstract object Anchor { get; }
    /// <summary>Set for nodes addressed by key rather than position, e.g. cell[B3].</summary>
    public virtual string? Key => null;
    /// <summary>Optional name matched by kind[Name] in paths, e.g. sheet[Sales].</summary>
    public virtual string? Name => null;
    public virtual string Format => Parent?.Format ?? throw new InvalidOperationException("A detached node has no format.");

    /// <summary>Children, projected from the native tree on every call.</summary>
    public IReadOnlyList<Node> Children => Attach(ProjectChildren());
    protected virtual IEnumerable<Node> ProjectChildren() => [];

    /// <summary>Canonical property values: EMU lengths, point sizes, RRGGBB colors, true/false. Only properties that are set.</summary>
    public abstract IReadOnlyDictionary<string, string> GetProps();

    public string? Text => GetProps().GetValueOrDefault("text");

    /// <summary>What the node looks like where its own <paramref name="props"/> are silent: values inherited from a theme,
    /// layout, master or style, in the same canonical units. Null when the format has nothing to resolve.</summary>
    public virtual IReadOnlyDictionary<string, string>? GetComputed(IReadOnlyDictionary<string, string> props) => null;

    /// <summary>A child that exists only once addressed, such as an empty spreadsheet cell. Null when the kind has no such children.</summary>
    public Node? ResolveVirtual(string kind, string key)
    {
        var child = ProjectVirtual(kind, key);
        if (child is null) return null;
        child.Parent = this;
        child.Path = ChildPath(kind, child.Key ?? key);
        return child;
    }

    protected virtual Node? ProjectVirtual(string kind, string key) => null;

    public virtual string GetRaw() =>
        throw new WriterException(ErrorCode.UnsupportedKind, $"{Kind} has no raw form", "Use get without --raw.");

    /// <summary>Writes one canonical value that the registry has already validated.</summary>
    public virtual void SetProp(string name, string value) => throw ReadOnly();

    /// <summary>Writes the props of one edit, in order. A node whose props share one record (a cell's style) writes those together.</summary>
    public virtual void SetProps(IEnumerable<KeyValuePair<string, string>> props)
    {
        foreach (var (name, value) in props) SetProp(name, value);
    }

    /// <summary>Creates a child at a 1-based position among this node's children (null = append) and returns it.</summary>
    public virtual Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index) => throw ReadOnly();

    public virtual void Remove() => throw ReadOnly();

    public virtual void MoveTo(Node newParent, int? index) => throw ReadOnly();

    public virtual void SetRaw(string raw) => throw ReadOnly();

    /// <summary>Creates a copy of this node under a parent at a 1-based position among its children (null = append) and returns it.</summary>
    public virtual Node CopyTo(Node newParent, int? index) =>
        throw new WriterException(ErrorCode.UnsupportedKind, $"A {Kind} cannot be copied in {Format} files", "copy works on .pptx slides and on the shapes, pictures and tables on them.");

    /// <summary>Creates a child from its raw XML, as 'get --raw' printed it, at a 1-based position (null = append) and returns it.</summary>
    public virtual Node AddRaw(string raw, int? index) =>
        throw new WriterException(ErrorCode.UnsupportedKind, $"A {Kind} in {Format} files takes no raw children", "add --raw works on .pptx slides: a shape, picture or table as 'get --raw' printed it.");

    WriterException ReadOnly() => new(ErrorCode.FormatReadonly, $"{Format} files cannot be edited ({Kind})",
        "Export to an editable format first: writer export <file> --to out.md");

    /// <summary>The current projection of the same native object among this node's children, or null.</summary>
    public Node? FindChild(object anchor) => Children.FirstOrDefault(c => ReferenceEquals(c.Anchor, anchor));

    /// <summary>Binary content for nodes that have it (images): content type and bytes. Null otherwise.</summary>
    public virtual (string ContentType, byte[] Data)? GetBinary() => null;

    IReadOnlyList<Node> Attach(IEnumerable<Node> children)
    {
        var list = children as List<Node> ?? children.ToList();
        var counts = new Dictionary<string, int>();
        foreach (var child in list)
        {
            var n = counts[child.Kind] = counts.GetValueOrDefault(child.Kind) + 1;
            child.Parent = this;
            child.Path = ChildPath(child.Kind, child.Key ?? n.ToString(CultureInfo.InvariantCulture));
        }
        return list;
    }

    string ChildPath(string kind, string key) =>
        Registry.Find(kind)?.Singleton == true
            ? $"{(Path == "/" ? "" : Path)}/{kind}"
            : $"{(Path == "/" ? "" : Path)}/{kind}[{key}]";
}
