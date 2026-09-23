using System.Globalization;

namespace Writer.Core;

/// <summary>A piece of text with uniform inline formatting, the common currency between markdown, HTML and the Office formats.
/// Color and Highlight are RRGGBB; Size is points. Change is "inserted" or "deleted" for text under a tracked change,
/// with the Author and Date (ISO 8601) of that change when known.</summary>
public sealed record RunSpec(string Text, bool Bold = false, bool Italic = false, bool Strike = false, bool Code = false, string? Link = null,
    bool Underline = false, string? Color = null, string? Size = null, string? Font = null, string? Highlight = null,
    string? Change = null, string? Author = null, string? Date = null)
{
    /// <summary>The spec a run node's properties describe.</summary>
    public static RunSpec FromProps(IReadOnlyDictionary<string, string> p) => new(p.GetValueOrDefault("text") ?? "",
        Bold: p.GetValueOrDefault("bold") == "true",
        Italic: p.GetValueOrDefault("italic") == "true",
        Strike: p.GetValueOrDefault("strike") == "true",
        Code: p.GetValueOrDefault("code") == "true" || p.GetValueOrDefault("font") is "Consolas" or "Courier New",
        Link: p.GetValueOrDefault("link"),
        Underline: p.GetValueOrDefault("underline") == "true",
        Color: p.GetValueOrDefault("color") is { } c && c != "none" ? c : null,
        Size: p.GetValueOrDefault("size"),
        Font: p.GetValueOrDefault("font"),
        Highlight: p.GetValueOrDefault("highlight") is { } h && h != "none" ? h : null,
        Change: p.GetValueOrDefault("change"),
        Author: p.GetValueOrDefault("author"),
        Date: p.GetValueOrDefault("date"));

    public bool Deleted => Change == "deleted";

    /// <summary>For places that keep no revisions (slides, markdown, headers): an insertion shown underlined, a deletion struck through.</summary>
    public RunSpec Flattened() => Change is null ? this
        : this with { Underline = Underline || Change == "inserted", Strike = Strike || Deleted, Change = null, Author = null, Date = null };

    /// <summary>Half-points for docx, as the SDK wants them.</summary>
    public string? HalfPoints => Size is null ? null : ((int)Math.Round(double.Parse(Size, CultureInfo.InvariantCulture) * 2)).ToString(CultureInfo.InvariantCulture);
}
