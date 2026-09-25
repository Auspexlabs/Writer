namespace Writer.Core;

/// <summary>Validated edits. Every method returns the node projected again, so its path and kind reflect the change.</summary>
public static class Mutations
{
    public static Node Set(Node node, IEnumerable<KeyValuePair<string, string>> props)
    {
        var canonical = Registry.Normalize(node.Format, node.Kind, props);
        node.SetProps(Ordered(canonical));
        return Refresh(node);
    }

    public static Node Add(Node parent, string kind, IEnumerable<KeyValuePair<string, string>> props, int? index)
    {
        Registry.CheckParent(parent.Format, kind, parent.Kind);
        var canonical = Registry.Normalize(parent.Format, kind, props);
        var created = parent.Add(kind, new Dictionary<string, string>(Ordered(canonical)), index);
        return parent.FindChild(created.Anchor) ?? created;
    }

    public static Node Move(Node node, Node newParent, int? index)
    {
        Registry.CheckParent(node.Format, node.Kind, newParent.Kind);
        node.MoveTo(newParent, index);
        return newParent.FindChild(node.Anchor) ?? node;
    }

    /// <summary>A copy of the node under newParent: the same markup, with what it refers to.</summary>
    public static Node Copy(Node node, Node newParent, int? index)
    {
        Registry.CheckParent(node.Format, node.Kind, newParent.Kind);
        var copy = node.CopyTo(newParent, index);
        return newParent.FindChild(copy.Anchor) ?? copy;
    }

    /// <summary>A child made from raw XML, as 'get --raw' printed it; the parent checks what the XML may be.</summary>
    public static Node AddRaw(Node parent, string raw, int? index)
    {
        var created = parent.AddRaw(raw, index);
        return parent.FindChild(created.Anchor) ?? created;
    }

    /// <summary>The same native object projected again from its parent. A paragraph restyled as a heading changes path and kind.</summary>
    public static Node Refresh(Node node) => node.Parent?.FindChild(node.Anchor) ?? node;

    /// <summary>list before level and value before type; a picture is reset, replaced or cut out before it is adjusted, and compressed
    /// last; a deck's palette before its fonts; text fitted once it and its box are set; everything else in the order given.</summary>
    static IEnumerable<KeyValuePair<string, string>> Ordered(Dictionary<string, string> props) =>
        props.OrderBy(p => p.Key switch { "reset" => -3, "src" => -2, "background" or "palette" => -1, "level" or "type" or "restart" => 1, "sectionBreak" or "caption" => -1, "bookmark" or "dropCap" => 1, "compress" or "fit" => 2, _ => 0 });
}
