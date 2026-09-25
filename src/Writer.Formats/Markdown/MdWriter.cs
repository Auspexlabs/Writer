using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Writer.Core;

namespace Writer.Formats.Markdown;

/// <summary>Renders model blocks back to markdown. Only edited blocks go through here.</summary>
static class MdWriter
{
    public static string Render(MdBlock block) => block switch
    {
        MdHeading h => new string('#', Math.Clamp(h.Level, 1, 6)) + " " + Inline(h.Runs).Replace("\n", " "),
        MdItems items => Items(items),
        MdCode c => Fence(c),
        MdImage i => $"![{Escape(i.Alt)}]({i.Src})",
        MdTable t => Table(t),
        _ => "",
    };

    static string Items(MdItems block)
    {
        var sb = new StringBuilder();
        var indentAt = new int[12];
        var counters = new int[12];
        var lastList = new string?[12];
        string? previous = null;
        for (var i = 0; i < block.Items.Count; i++)
        {
            var item = block.Items[i];
            if (item.List == "none")
            {
                if (i > 0) sb.Append("\n\n");
                sb.Append(Paragraph(item.Runs));
                previous = "none";
                Array.Clear(lastList);
                continue;
            }
            var level = Math.Clamp(item.Level, 0, 10);
            if (i > 0) sb.Append(previous == "none" ? "\n\n" : "\n");
            for (var deeper = level + 1; deeper < lastList.Length; deeper++) lastList[deeper] = null;
            var indent = level == 0 ? 0 : indentAt[level] > 0 ? indentAt[level] : 2 * level;
            string marker;
            if (item.List != "bullet") // number, and the Word kinds markdown has no marker for (outline, chinese)
            {
                counters[level] = lastList[level] == item.List ? counters[level] + 1 : 1;
                marker = counters[level].ToString(CultureInfo.InvariantCulture) + ". ";
            }
            else marker = "- ";
            lastList[level] = item.List;
            indentAt[level + 1] = indent + marker.Length;
            var continuation = "\\\n" + new string(' ', indent + marker.Length);
            sb.Append(' ', indent).Append(marker).Append(Inline(item.Runs).Replace("\n", continuation));
            previous = item.List;
        }
        return sb.ToString();
    }

    public static string Paragraph(IEnumerable<MdRun> runs) => Inline(runs).Replace("\n", "\\\n");

    static string Fence(MdCode code)
    {
        var fence = "```";
        while (code.Text.Contains(fence, StringComparison.Ordinal)) fence += "`";
        return fence + code.Lang + "\n" + code.Text + "\n" + fence;
    }

    static string Table(MdTable table)
    {
        var cols = table.Rows.Count == 0 ? 0 : table.Rows.Max(r => r.Cells.Count);
        if (cols == 0) return "";
        var lines = new List<string>();
        for (var i = 0; i < table.Rows.Count; i++)
        {
            var row = table.Rows[i];
            lines.Add("| " + string.Join(" | ", Enumerable.Range(0, cols).Select(c => c < row.Cells.Count ? Cell(row.Cells[c]) : "")) + " |");
            if (i == 0)
                lines.Add("| " + string.Join(" | ", Enumerable.Range(0, cols).Select(c => Separator(table.Aligns.ElementAtOrDefault(c)))) + " |");
        }
        return string.Join("\n", lines);
    }

    static string Cell(MdCell cell) => Inline(cell.Runs).Replace("|", "\\|").Replace("\n", "<br>");

    static string Separator(string? align) => align switch
    {
        "center" => ":---:",
        "right" => "---:",
        "left" => ":---",
        _ => "---",
    };

    public static string Inline(IEnumerable<MdRun> runs) => Inline(runs.Select(r => r.ToSpec()));

    static readonly Regex LineStartMarkers = new(@"(?m)^([ \t]*)([#>+\-|])");
    static readonly Regex LineStartNumbers = new(@"(?m)^([ \t]*)(\d+)([.)])(?=\s)");

    /// <summary>Runs → inline markdown, with characters that would start a block escaped at line starts.</summary>
    public static string Inline(IEnumerable<RunSpec> specs)
    {
        var text = string.Concat(specs.Select(Run));
        text = LineStartMarkers.Replace(text, "$1\\$2");
        return LineStartNumbers.Replace(text, "$1$2\\$3");
    }

    static string Run(RunSpec run)
    {
        if (run.Text.Length == 0) return "";
        var core = run.Text.Trim();
        if (core.Length == 0) return run.Text;
        var lead = run.Text[..run.Text.IndexOf(core, StringComparison.Ordinal)];
        var trail = run.Text[(lead.Length + core.Length)..];
        var s = run.Code ? CodeSpan(core) : Escape(core);
        if (run.Bold) s = "**" + s + "**";
        if (run.Italic) s = "*" + s + "*";
        if (run.Strike) s = "~~" + s + "~~";
        if (run.Link is not null) s = "[" + s + "](" + run.Link + ")";
        return lead + s + trail;
    }

    static string CodeSpan(string text)
    {
        var ticks = "`";
        while (text.Contains(ticks, StringComparison.Ordinal)) ticks += "`";
        var pad = text.StartsWith('`') || text.EndsWith('`') ? " " : "";
        return ticks + pad + text + pad + ticks;
    }

    public static string Escape(string text)
    {
        var sb = new StringBuilder(text.Length + 8);
        foreach (var ch in text)
        {
            if (ch is '\\' or '*' or '_' or '`' or '[' or ']' or '<' or '~') sb.Append('\\');
            sb.Append(ch);
        }
        return sb.ToString();
    }
}
