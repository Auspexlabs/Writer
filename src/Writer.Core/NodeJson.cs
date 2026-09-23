using System.Buffers;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace Writer.Core;

/// <summary>The JSON shape every data verb prints.</summary>
public static class NodeJson
{
    static readonly JsonWriterOptions Options = new() { Indented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    /// <summary>A node with its properties and, up to <paramref name="depth"/> levels down, its children.
    /// The last level is summarised (kind, path, text preview).</summary>
    public static string Serialize(Node node, int depth) => Write(w => WriteNode(w, node, depth));

    /// <summary>Like <see cref="Serialize(Node, int)"/>, leaving out every node of the given kinds together with its subtree.
    /// A UI that renders paragraphs from their html property does not need thousands of run nodes.</summary>
    public static string Serialize(Node node, int depth, IReadOnlySet<string> skip) =>
        Write(w => WriteNode(w, node, depth, skip.Count == 0 ? null : n => !skip.Contains(n.Kind)));

    public static string Summaries(IEnumerable<Node> nodes) => Write(w =>
    {
        w.WriteStartArray();
        foreach (var n in nodes) WriteSummary(w, n);
        w.WriteEndArray();
    });

    public static string Write(Action<Utf8JsonWriter> body) => Write(body, Options);

    /// <summary>Single-line JSON, used for JSON-typed property values such as table data.</summary>
    public static string Compact(Action<Utf8JsonWriter> body) => Write(body, Options with { Indented = false });

    static string Write(Action<Utf8JsonWriter> body, JsonWriterOptions options)
    {
        var buffer = new ArrayBufferWriter<byte>();
        using (var w = new Utf8JsonWriter(buffer, options)) body(w);
        return Encoding.UTF8.GetString(buffer.WrittenSpan);
    }

    public static void WriteNode(Utf8JsonWriter w, Node node, int depth) => WriteNode(w, node, depth, null);

    static void WriteNode(Utf8JsonWriter w, Node node, int depth, Func<Node, bool>? include)
    {
        w.WriteStartObject();
        w.WriteString("kind", node.Kind);
        w.WriteString("path", node.Path);
        var props = node.GetProps();
        w.WriteStartObject("props");
        foreach (var (name, value) in Registry.ToDisplay(node.Format, node.Kind, props))
        {
            if (Registry.IsJson(node.Format, node.Kind, name))
            {
                w.WritePropertyName(name);
                w.WriteRawValue(value);
            }
            else w.WriteString(name, value);
        }
        w.WriteEndObject();
        if (node.GetComputed(props) is { Count: > 0 } computed)
        {
            w.WriteStartObject("computed");
            foreach (var (name, value) in Registry.ToDisplay(node.Format, node.Kind, computed)) w.WriteString(name, value);
            w.WriteEndObject();
        }
        if (depth > 0)
        {
            var children = include is null ? node.Children : node.Children.Where(include).ToList();
            if (children.Count > 0)
            {
                w.WriteStartArray("children");
                foreach (var child in children)
                {
                    if (depth > 1) WriteNode(w, child, depth - 1, include);
                    else WriteSummary(w, child);
                }
                w.WriteEndArray();
            }
        }
        w.WriteEndObject();
    }

    public static void WriteSummary(Utf8JsonWriter w, Node node)
    {
        w.WriteStartObject();
        w.WriteString("kind", node.Kind);
        w.WriteString("path", node.Path);
        if (node.Text is { Length: > 0 } text) w.WriteString("text", Preview(text, 80));
        w.WriteEndObject();
    }

    public static string Preview(string text, int max)
    {
        var flat = text.ReplaceLineEndings(" ");
        return flat.Length <= max ? flat : string.Concat(flat.AsSpan(0, max - 1), "…");
    }
}
