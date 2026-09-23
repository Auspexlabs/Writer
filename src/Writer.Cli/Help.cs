using System.Text;
using Writer.Core;
using Writer.Formats;

namespace Writer.Cli;

/// <summary>Help text generated from the registry, so it can never drift from the code.</summary>
static class Help
{
    public static string Render(IReadOnlyList<string> args, bool json)
    {
        if (args.Count == 0) return json ? OverviewJson() : Overview();
        var format = Adapters.CanonicalFormat(args[0])
            ?? throw new WriterException(ErrorCode.UnknownFormat, $"Unknown format '{args[0]}'", $"Formats: {Formats()}. Example: writer help docx paragraph");
        if (args.Count == 1) return json ? FormatJson(format) : FormatText(format);
        var kind = Registry.Get(format, args[^1]);
        var verb = args.Count >= 3 ? args[1] : null;
        var props = Registry.PropsFor(kind, format).Where(p => verb switch
        {
            "set" or "add" => !p.ReadOnly,
            "get" or "query" => !p.WriteOnly,
            _ => true,
        }).ToList();
        return json ? KindJson(format, kind, props) : KindText(format, kind, props, verb);
    }

    static string Formats() => string.Join(", ", Adapters.All.Select(a => a.Format));

    static string Overview()
    {
        var sb = new StringBuilder();
        sb.Append($"writer {Runner.Version} — documents for AI agents\n\n");
        sb.Append("Usage: writer <command> <file> [path] [options]\n\nCommands:\n");
        foreach (var (verb, spec) in Args.Spec)
            if (verb != "version") sb.Append($"  {verb + " " + spec.Usage,-44} {spec.Description}\n");
        sb.Append($"\nFormats: {Formats()}   (writer help <format> lists its elements)\n\n");
        sb.Append("Paths:\n");
        sb.Append("  /body/paragraph[2]           2nd paragraph in the body\n");
        sb.Append("  //heading                    every heading\n");
        sb.Append("  //heading[3]                 3rd heading in the whole document\n");
        sb.Append("  //paragraph[@text~=\"Q4\"]     paragraphs containing Q4 (use = for an exact match)\n");
        sb.Append("  /body/table[1]/row[-1]       last row of the first table\n");
        sb.Append("  /body/*[3]                   3rd block of any kind\n\n");
        sb.Append("Data commands print JSON. Errors go to stderr as {\"error\":{\"code\",\"message\",\"hint\"}} with a non-zero exit code.\n");
        return sb.ToString();
    }

    static string OverviewJson() => NodeJson.Write(w =>
    {
        w.WriteStartObject();
        w.WriteString("version", Runner.Version);
        w.WriteStartArray("commands");
        foreach (var (verb, spec) in Args.Spec)
        {
            w.WriteStartObject();
            w.WriteString("name", verb);
            w.WriteString("usage", spec.Usage);
            w.WriteString("description", spec.Description);
            w.WriteEndObject();
        }
        w.WriteEndArray();
        w.WriteStartArray("formats");
        foreach (var a in Adapters.All) w.WriteStringValue(a.Format);
        w.WriteEndArray();
        w.WriteEndObject();
    }) + "\n";

    static string FormatText(string format)
    {
        var sb = new StringBuilder();
        sb.Append($"{format} elements (writer help {format} <element> shows properties):\n");
        foreach (var k in Registry.ForFormat(format))
        {
            sb.Append($"  {k.Name,-10} {k.Description}");
            if (k.Parents.Length > 0) sb.Append($"   under: {string.Join(", ", k.Parents)}");
            sb.Append('\n');
        }
        return sb.ToString();
    }

    static string FormatJson(string format) => NodeJson.Write(w =>
    {
        w.WriteStartObject();
        w.WriteString("format", format);
        w.WriteStartArray("elements");
        foreach (var k in Registry.ForFormat(format))
        {
            w.WriteStartObject();
            w.WriteString("name", k.Name);
            w.WriteString("description", k.Description);
            w.WriteStartArray("parents");
            foreach (var p in k.Parents) w.WriteStringValue(p);
            w.WriteEndArray();
            w.WriteEndObject();
        }
        w.WriteEndArray();
        w.WriteEndObject();
    }) + "\n";

    static string KindText(string format, Kind kind, List<Prop> props, string? verb)
    {
        var sb = new StringBuilder();
        sb.Append($"{kind.Name} ({format}): {kind.Description}\n");
        if (kind.Parents.Length > 0) sb.Append($"Goes under: {string.Join(", ", kind.Parents)}\n");
        sb.Append(props.Count == 0 ? "\nNo properties.\n" : verb is null ? "\nProperties:\n" : $"\nProperties for {verb}:\n");
        var width = props.Select(p => p.Name.Length).Append(11).Max();
        foreach (var p in props)
        {
            var flags = p.ReadOnly ? " (read-only)" : p.WriteOnly ? " (write-only)" : "";
            var aliases = p.Aliases.Length > 0 ? $" Alias: {string.Join(", ", p.Aliases)}." : "";
            sb.Append($"  {p.Name.PadRight(width)} {TypeLabel(p),-26} {p.Description}{flags}{aliases}\n");
            if (p.Example is not null && !p.ReadOnly) sb.Append($"  {"".PadRight(width)} {"",-26} e.g. --prop {p.Name}=\"{p.Example}\"\n");
        }
        if (kind.Parents.Length > 0)
        {
            var parent = kind.Parents.FirstOrDefault(p => Registry.Find(p)?.Formats.Contains(format) == true) ?? kind.Parents[0];
            var path = parent == "body" ? "/body" : $"//{parent}[1]";
            var example = props.FirstOrDefault(p => p.Name == "text") ?? props.FirstOrDefault(p => !p.ReadOnly && p.Example is not null);
            var prop = example is null ? "" : $" --prop {example.Name}=\"{example.Example}\"";
            sb.Append($"\nExample:\n  writer add file.{format} {path} --type {kind.Name}{prop}\n");
        }
        return sb.ToString();
    }

    static string KindJson(string format, Kind kind, List<Prop> props) => NodeJson.Write(w =>
    {
        w.WriteStartObject();
        w.WriteString("format", format);
        w.WriteString("element", kind.Name);
        w.WriteString("description", kind.Description);
        w.WriteStartArray("parents");
        foreach (var p in kind.Parents) w.WriteStringValue(p);
        w.WriteEndArray();
        w.WriteStartArray("props");
        foreach (var p in props)
        {
            w.WriteStartObject();
            w.WriteString("name", p.Name);
            w.WriteString("type", p.Type.ToString().ToLowerInvariant());
            w.WriteString("description", p.Description);
            if (p.Aliases.Length > 0)
            {
                w.WriteStartArray("aliases");
                foreach (var a in p.Aliases) w.WriteStringValue(a);
                w.WriteEndArray();
            }
            if (p.Values is not null)
            {
                w.WriteStartArray("values");
                foreach (var v in p.Values) w.WriteStringValue(v);
                w.WriteEndArray();
            }
            if (p.Min != int.MinValue) w.WriteNumber("min", p.Min);
            if (p.Max != int.MaxValue) w.WriteNumber("max", p.Max);
            if (p.Example is not null) w.WriteString("example", p.Example);
            if (p.ReadOnly) w.WriteBoolean("readOnly", true);
            if (p.WriteOnly) w.WriteBoolean("writeOnly", true);
            w.WriteEndObject();
        }
        w.WriteEndArray();
        w.WriteEndObject();
    }) + "\n";

    static string TypeLabel(Prop p) => p.Type switch
    {
        PropType.Enum => string.Join(" | ", p.Values!),
        PropType.Int when p.Min != int.MinValue && p.Max != int.MaxValue => $"int {p.Min}-{p.Max}",
        PropType.Int when p.Min != int.MinValue => $"int >= {p.Min}",
        PropType.Int => "int",
        PropType.Bool => "true | false",
        PropType.Length => "length (2cm, 1in, 12pt)",
        PropType.Points => "points (12pt)",
        PropType.Color => "color (RRGGBB or name)",
        PropType.Md => "inline markdown",
        PropType.Html => "inline html",
        PropType.Json => "json",
        _ => "string",
    };
}
