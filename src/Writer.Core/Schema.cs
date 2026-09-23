namespace Writer.Core;

public enum PropType { String, Int, Bool, Length, Points, Color, Enum, Md, Html, Json }

/// <summary>One property of an element kind. Values are strings on the wire; Type says how to validate and display them.</summary>
public sealed record Prop(string Name, PropType Type, string Description)
{
    public string[] Aliases { get; init; } = [];
    /// <summary>Allowed values for <see cref="PropType.Enum"/>.</summary>
    public string[]? Values { get; init; }
    public int Min { get; init; } = int.MinValue;
    public int Max { get; init; } = int.MaxValue;
    public string? Example { get; init; }
    public bool ReadOnly { get; init; }
    public bool WriteOnly { get; init; }
    /// <summary>Formats this property applies to; null means every format of its kind.</summary>
    public string[]? Formats { get; init; }
}

/// <summary>One element kind of the unified tree.</summary>
public sealed record Kind(string Name, string Description)
{
    public string[] Formats { get; init; } = [];
    /// <summary>Kinds this one may be added to or moved under.</summary>
    public string[] Parents { get; init; } = [];
    public Prop[] Props { get; init; } = [];
    /// <summary>Character-level node (run): left out of outlines.</summary>
    public bool Inline { get; init; }
    /// <summary>Outlines show one line and do not descend (tables, sheets).</summary>
    public bool Summary { get; init; }
    /// <summary>At most one per parent, so its path carries no index: /body, not /body[1].</summary>
    public bool Singleton { get; init; }
}
