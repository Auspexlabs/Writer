using System.Globalization;

namespace Writer.Core;

/// <summary>Evaluates parsed paths against a node tree.</summary>
public static class PathResolver
{
    public static IReadOnlyList<Node> Query(Node root, string path) => Run(root, PathParser.Parse(path), out _, out _);

    /// <summary>Exactly one match, or PATH_NOT_FOUND / PATH_AMBIGUOUS with a hint.</summary>
    public static Node Single(Node root, string path)
    {
        var found = Run(root, PathParser.Parse(path), out var lastContext, out var failed);
        if (found.Count == 1) return found[0];
        if (found.Count == 0)
            throw new WriterException(ErrorCode.PathNotFound, $"Nothing matches {path}", NotFoundHint(lastContext, failed));
        throw new WriterException(ErrorCode.PathAmbiguous, $"{path} matches {found.Count} nodes",
            $"Use one of them, e.g. {found[0].Path} or {found[1].Path}, or add --all where the command supports it.");
    }

    static List<Node> Run(Node root, IReadOnlyList<Segment> segments, out List<Node> lastContext, out Segment? failed)
    {
        var current = new List<Node> { root };
        lastContext = current;
        failed = null;
        foreach (var segment in segments)
        {
            var next = new List<Node>();
            if (segment.Deep)
            {
                var seen = new HashSet<object>(ReferenceEqualityComparer.Instance);
                var all = current.SelectMany(Descendants).Where(n => Matches(n, segment.Kind) && seen.Add(n.Anchor)).ToList();
                next.AddRange(Apply(all, segment, null));
            }
            else
            {
                foreach (var context in current)
                    next.AddRange(Apply(context.Children.Where(n => Matches(n, segment.Kind)).ToList(), segment, context));
            }
            if (next.Count == 0)
            {
                lastContext = current;
                failed = segment;
                return next;
            }
            current = next;
        }
        return current;
    }

    static bool Matches(Node n, string kind) => kind == "*" || n.Kind == kind;

    static IEnumerable<Node> Descendants(Node node)
    {
        foreach (var child in node.Children)
        {
            yield return child;
            foreach (var grandchild in Descendants(child)) yield return grandchild;
        }
    }

    static List<Node> Apply(List<Node> nodes, Segment segment, Node? context)
    {
        foreach (var selector in segment.Selectors)
        {
            nodes = selector switch
            {
                AttrSelector a => nodes.Where(n => AttrMatches(n, a)).ToList(),
                IndexSelector ix when ix.Index > 0 && nodes.Count > 0 && nodes[0].Key is not null
                    => ByKey(nodes, ix.Index.ToString(CultureInfo.InvariantCulture), segment.Kind, context),
                IndexSelector ix => ByIndex(nodes, ix.Index, segment.Kind, context),
                KeySelector k => ByKey(nodes, k.Key, segment.Kind, context),
                _ => nodes,
            };
        }
        return nodes;
    }

    static List<Node> ByIndex(List<Node> nodes, int index, string kind, Node? context)
    {
        var i = index > 0 ? index - 1 : nodes.Count + index;
        if (i >= 0 && i < nodes.Count) return [nodes[i]];
        if (index > 0 && context is not null && kind != "*"
            && context.ResolveVirtual(kind, index.ToString(CultureInfo.InvariantCulture)) is { } v) return [v];
        return [];
    }

    static List<Node> ByKey(List<Node> nodes, string key, string kind, Node? context)
    {
        var hits = nodes.Where(n => Same(n.Key, key) || Same(n.Name, key)).ToList();
        if (hits.Count == 0 && context is not null && kind != "*" && context.ResolveVirtual(kind, key) is { } v) hits.Add(v);
        return hits;
    }

    static bool Same(string? a, string b) => a is not null && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    static bool AttrMatches(Node n, AttrSelector a)
    {
        if (!Registry.ToDisplay(n.Format, n.Kind, n.GetProps()).TryGetValue(a.Name, out var value)) return false;
        return a.Contains ? value.Contains(a.Value, StringComparison.OrdinalIgnoreCase) : value == a.Value;
    }

    static string NotFoundHint(List<Node> context, Segment? failed)
    {
        const string outline = "Run 'writer view <file> outline' to see every path.";
        if (context.Count == 0 || failed is null) return outline;
        var at = context[0];
        var children = at.Children;
        if (children.Count == 0) return $"{at.Path} has no children. {outline}";
        var shown = string.Join(", ", children.Take(10).Select(c => c.Path));
        return $"Under {at.Path}: {shown}{(children.Count > 10 ? ", ..." : "")}";
    }
}
