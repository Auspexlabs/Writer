using System.Globalization;

namespace Writer.Core;

public abstract record Selector;
/// <summary>[3] or [-1]: position among the candidates, 1-based, negative from the end.</summary>
public sealed record IndexSelector(int Index) : Selector;
/// <summary>[B3], [Sales], ["Q1 Sales"]: matches a node's Key or Name.</summary>
public sealed record KeySelector(string Key) : Selector;
/// <summary>[@text="x"] or [@text~="x"]: matches a displayed property exactly or by substring.</summary>
public sealed record AttrSelector(string Name, bool Contains, string Value) : Selector;

public sealed record Segment(bool Deep, string Kind, IReadOnlyList<Selector> Selectors);

/// <summary>Parses the simplified XPath used everywhere: /a[1]/b[@x="y"], //c[-1], /sheet[Sales]/cell[B3].</summary>
public static class PathParser
{
    public static IReadOnlyList<Segment> Parse(string path)
    {
        var p = path.Trim();
        if (p.Length == 0 || p[0] != '/') throw Bad(path, "Paths start with / (children) or // (any depth).");
        var segments = new List<Segment>();
        var i = 0;
        while (i < p.Length)
        {
            var deep = p.AsSpan(i).StartsWith("//");
            i += deep ? 2 : 1;
            if (i >= p.Length)
            {
                if (segments.Count == 0 && !deep) return segments;
                throw Bad(path, "Missing element name after /.");
            }
            var start = i;
            while (i < p.Length && (char.IsLetterOrDigit(p[i]) || p[i] is '*' or '_' or '-')) i++;
            if (i == start) throw Bad(path, $"Expected an element name at position {start + 1}.");
            var kind = p[start..i];
            var selectors = new List<Selector>();
            while (i < p.Length && p[i] == '[')
            {
                i++;
                selectors.Add(ReadSelector(p, ref i, path));
            }
            if (i < p.Length && p[i] != '/') throw Bad(path, $"Unexpected '{p[i]}' at position {i + 1}.");
            segments.Add(new Segment(deep, kind, selectors));
        }
        return segments;
    }

    static Selector ReadSelector(string p, ref int i, string path)
    {
        const string attrHelp = "Attribute selectors look like [@text=\"x\"] or [@text~=\"x\"].";
        if (i < p.Length && p[i] == '@')
        {
            i++;
            var s = i;
            while (i < p.Length && (char.IsLetterOrDigit(p[i]) || p[i] is '_' or '-')) i++;
            var name = p[s..i];
            if (name.Length == 0) throw Bad(path, attrHelp);
            var contains = i < p.Length && p[i] == '~';
            if (contains) i++;
            if (i >= p.Length || p[i] != '=') throw Bad(path, attrHelp);
            i++;
            var value = ReadValue(p, ref i, path, out _);
            Close(p, ref i, path);
            return new AttrSelector(name, contains, value);
        }
        var raw = ReadValue(p, ref i, path, out var quoted);
        Close(p, ref i, path);
        if (!quoted && int.TryParse(raw, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var n))
        {
            if (n == 0) throw Bad(path, "Positions start at 1; use [-1] for the last one.");
            return new IndexSelector(n);
        }
        if (raw.Length == 0) throw Bad(path, "Empty [] selector.");
        return new KeySelector(raw);
    }

    static string ReadValue(string p, ref int i, string path, out bool quoted)
    {
        quoted = false;
        if (i < p.Length && p[i] is '"' or '\'')
        {
            var quote = p[i++];
            var s = i;
            while (i < p.Length && p[i] != quote) i++;
            if (i >= p.Length) throw Bad(path, "Unterminated quote in selector.");
            quoted = true;
            return p[s..i++];
        }
        var start = i;
        while (i < p.Length && p[i] != ']') i++;
        return p[start..i].Trim();
    }

    static void Close(string p, ref int i, string path)
    {
        if (i >= p.Length || p[i] != ']') throw Bad(path, "Missing ] in selector.");
        i++;
    }

    static WriterException Bad(string path, string why) =>
        new(ErrorCode.Usage, $"Bad path '{path}': {why}",
            "Examples: /body/paragraph[2], //heading[3], //paragraph[@text~=\"Q4\"], /sheet[Sales]/cell[B3], /body/table[1]/row[-1]");
}
