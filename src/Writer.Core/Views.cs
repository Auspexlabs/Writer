using System.Text;

namespace Writer.Core;

/// <summary>Whole-document text views.</summary>
public static class Views
{
    /// <summary>One line per block-level node: path, short properties, text preview.</summary>
    public static string Outline(Node root)
    {
        var sb = new StringBuilder();
        Walk(root, 0);
        return sb.ToString();

        void Walk(Node node, int depth)
        {
            var kind = Registry.Find(node.Kind);
            if (kind?.Inline == true) return;
            sb.Append(' ', depth * 2).Append(node.Path);
            var props = Registry.ToDisplay(node.Format, node.Kind, node.GetProps());
            foreach (var (name, value) in props)
                if (name is not ("text" or "html" or "id") && !Registry.IsJson(node.Format, node.Kind, name) && value.Length <= 40)
                    sb.Append("  ").Append(name).Append('=').Append(value.ReplaceLineEndings(" "));
            if (props.TryGetValue("text", out var text) && text.Length > 0)
                sb.Append("  \"").Append(NodeJson.Preview(text, 60)).Append('"');
            sb.Append('\n');
            foreach (var child in node.Children)
                if (kind?.Summary != true || child.Kind != "row") Walk(child, depth + 1);
        }
    }

    /// <summary>Plain text: blocks separated by blank lines, tables as tab-separated rows.</summary>
    public static string Text(Node root)
    {
        var sb = new StringBuilder();
        Walk(root);
        return sb.ToString().TrimEnd('\n') + "\n";

        void Walk(Node node)
        {
            if (node.Kind == "decor") return;
            if (node.Kind is "slide" or "page") sb.Append("--- ").Append(node.Path).Append(" ---\n");
            if (node.Kind == "topic")
            {
                Topic(node, 0);
                return;
            }
            if (node.Kind is "table" or "sheet")
            {
                foreach (var row in node.Children.Where(c => c.Kind == "row"))
                {
                    var cells = row.GetProps().TryGetValue("data", out var data) ? RowValues(data) : row.Children.Select(c => c.Text ?? "");
                    sb.AppendJoin('\t', cells.Select(c => c.ReplaceLineEndings(" "))).Append('\n');
                }
                sb.Append('\n');
                return;
            }
            var blocks = node.Children.Where(c => Registry.Find(c.Kind)?.Inline != true).ToList();
            if (blocks.Count == 0)
            {
                if (Spliced(node) is { Length: > 0 } text) sb.Append(text).Append("\n\n");
                return;
            }
            foreach (var block in blocks) Walk(block);
        }

        /// <summary>A block's text with its citations in it, each at its character offset (their text is not the paragraph's own); a
        /// contents or works-cited list under its title.</summary>
        static string Spliced(Node node)
        {
            var text = node.Text ?? "";
            if (node.Kind is "toc" or "bibliography" && node.GetProps().GetValueOrDefault("title") is { Length: > 0 } title) return title + "\n" + text; // a list under its title
            foreach (var p in node.Children.Where(c => c.Kind == "citation").Select(c => c.GetProps())
                .OrderByDescending(p => int.TryParse(p.GetValueOrDefault("at"), out var at) ? at : int.MaxValue))
                text = text.Insert(int.TryParse(p.GetValueOrDefault("at"), out var at) ? Math.Clamp(at, 0, text.Length) : text.Length, p.GetValueOrDefault("text") ?? "");
            return text;
        }

        /// <summary>Mind map: the root on its own line, every descendant as an indented bullet.</summary>
        void Topic(Node topic, int depth)
        {
            if (depth > 0) sb.Append(' ', (depth - 1) * 2).Append("- ");
            sb.Append(topic.Text).Append('\n');
            foreach (var child in topic.Children) Topic(child, depth + 1);
        }
    }

    static IEnumerable<string> RowValues(string json)
    {
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        return doc.RootElement.ValueKind == System.Text.Json.JsonValueKind.Array
            ? doc.RootElement.EnumerateArray().Select(e => e.ValueKind == System.Text.Json.JsonValueKind.String ? e.GetString() ?? "" : e.GetRawText()).ToList()
            : [];
    }
}
