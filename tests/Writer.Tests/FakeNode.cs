using Writer.Core;

namespace Writer.Tests;

/// <summary>In-memory node for Core tests. Children are the same instances every time, like a stable native tree.</summary>
public sealed class FakeNode(string kind, string? key = null, string? name = null, Dictionary<string, string>? props = null) : Node
{
    public List<FakeNode> Kids { get; } = [];
    public Func<string, string, Node?>? Virtual { get; set; }

    public FakeNode Add(params FakeNode[] children)
    {
        Kids.AddRange(children);
        return this;
    }

    public override string Kind => kind;
    public override string? Key => key;
    public override string? Name => name;
    public override object Anchor => this;
    public override string Format => Parent?.Format ?? "fake";
    protected override IEnumerable<Node> ProjectChildren() => Kids;
    public override IReadOnlyDictionary<string, string> GetProps() => props ?? new Dictionary<string, string>();
    protected override Node? ProjectVirtual(string k, string v) => Virtual?.Invoke(k, v);
    public override string GetRaw() => $"<{kind}/>";
}
