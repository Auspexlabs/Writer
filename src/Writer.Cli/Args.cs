using System.Globalization;
using Writer.Core;

namespace Writer.Cli;

/// <summary>Parsed command line: verb, positionals, --options with values, --flags.</summary>
public sealed class Args
{
    /// <summary>verb → (usage, description, options that take a value, boolean flags).</summary>
    internal static readonly Dictionary<string, (string Usage, string Description, string[] Values, string[] Flags)> Spec = new()
    {
        ["create"] = ("<file> [--from template]", "New blank document, or a copy of a template", ["from"], []),
        ["get"] = ("<file> [path] [--depth n] [--raw]", "One node as JSON (path defaults to /)", ["depth"], ["raw"]),
        ["query"] = ("<file> <path>", "Every node matching a path", [], []),
        ["add"] = ("<file> <parent> (--type kind [--prop k=v]... | --raw xml) [--index n | --after path | --before path]", "Add an element, or put one back from 'get --raw' output", ["type", "prop", "raw", "index", "after", "before"], []),
        ["set"] = ("<file> <path> [--prop k=v]... [--raw xml] [--all]", "Change properties, or replace the raw XML", ["prop", "raw"], ["all"]),
        ["remove"] = ("<file> <path> [--all]", "Delete elements", [], ["all"]),
        ["move"] = ("<file> <path> --to parent [--index n | --after path | --before path]", "Move an element under another parent", ["to", "index", "after", "before"], []),
        ["copy"] = ("<file> <path> --to parent [--index n | --after path | --before path]", "Copy an element under a parent (pptx: a slide, or a shape, picture or table)", ["to", "index", "after", "before"], []),
        ["view"] = ("<file> [outline|structure|text|html|json]", "The whole document; structure is the short form: headings, tables, sheets with their header row", [], []),
        ["search"] = ("<file> <text> [--ignore-case]", "Every block whose text contains the words, with its path", [], ["ignore-case"]),
        ["section"] = ("<file> <heading path> [--md markdown | --remove]", "A heading and everything up to the next heading of its level: print it, replace it with the markdown's blocks (heading included), or remove it", ["md"], ["remove"]),
        ["replace"] = ("<file> [path] --find text --with text [--preview] [--ignore-case]", "Find and replace across the document or under a path; Word paragraphs keep their formatting; --preview counts without writing", ["find", "with"], ["preview", "ignore-case"]),
        ["formula"] = ("<file> <cell or range path> <formula> [--check]", "xlsx: write a formula (a range fills like Excel, shifting relative references) and report what it refers to, with warnings; --check reports only", [], ["check"]),
        ["batch"] = ("<file> --run \"command\" [--run ...]", "Run the commands in order and save once; when one fails nothing is written", ["run"], []),
        ["export"] = ("<file> --to out.ext", "Write the document as another format: md, docx, html, json", ["to"], []),
        ["help"] = ("[format] [element] [--json]", "Elements and their properties", [], ["json"]),
        ["mcp"] = ("", "Serve the MCP protocol over stdio for AI agents", [], []),
        ["serve"] = ("[--dir folder] [--list-depth n] [--port n] [--allow-origin url]... [--no-token]", "Local HTTP API for a UI: POST /run, GET /files /file /html /json /outline /text /events", ["port", "allow-origin", "dir", "ui", "list-depth"], ["no-token"]),
        ["watch"] = ("<file> [--port n]", "Live preview in the browser, refreshed whenever the file changes", ["port", "allow-origin", "dir", "ui"], []),
        ["app"] = ("[--dir folder] [--list-depth n] [--port n] [--no-browser] [--drafts dir] [--grants-stdin]", "Open the bundled editor UI on a folder of documents", ["port", "allow-origin", "dir", "ui", "list-depth", "drafts"], ["no-token", "no-browser", "grants-stdin"]),
        ["version"] = ("", "Print the version", [], []),
    };

    public string Verb { get; init; } = "";
    public List<string> Positional { get; } = [];
    public Dictionary<string, List<string>> Options { get; } = new();
    public HashSet<string> Flags { get; } = new();

    public string? Opt(string name) => Options.TryGetValue(name, out var v) ? v[^1] : null;
    public IReadOnlyList<string> All(string name) => Options.TryGetValue(name, out var v) ? v : [];
    public bool Flag(string name) => Flags.Contains(name);

    public string Need(int index, string what) =>
        index < Positional.Count
            ? Positional[index]
            : throw new WriterException(ErrorCode.Usage, $"Missing {what}", $"Usage: writer {Verb} {Spec[Verb].Usage}");

    public int Int(string option, int fallback)
    {
        var v = Opt(option);
        if (v is null) return fallback;
        return int.TryParse(v, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var n)
            ? n
            : throw new WriterException(ErrorCode.Usage, $"--{option} must be a number", $"Got '{v}'.");
    }

    public static Args Parse(string[] argv)
    {
        if (argv.Length == 0) return new Args { Verb = "help" };
        var verb = argv[0] switch { "-h" or "--help" => "help", "-v" or "--version" => "version", var v => v };
        if (!Spec.TryGetValue(verb, out var spec))
            throw new WriterException(ErrorCode.Usage, $"Unknown command '{argv[0]}'",
                $"Commands: {string.Join(", ", Spec.Keys.Where(k => k != "version"))}. Run 'writer help'.");
        var args = new Args { Verb = verb };
        for (var i = 1; i < argv.Length; i++)
        {
            var token = argv[i];
            if (!token.StartsWith("--", StringComparison.Ordinal) || token.Length == 2)
            {
                args.Positional.Add(token);
                continue;
            }
            var name = token[2..];
            string? inline = null;
            var eq = name.IndexOf('=');
            if (eq >= 0)
            {
                inline = name[(eq + 1)..];
                name = name[..eq];
            }
            if (name == "debug" || spec.Flags.Contains(name))
            {
                args.Flags.Add(name);
                continue;
            }
            if (!spec.Values.Contains(name))
                throw new WriterException(ErrorCode.Usage, $"Unknown option --{name} for '{verb}'", $"Usage: writer {verb} {spec.Usage}");
            var value = inline ?? (i + 1 < argv.Length
                ? argv[++i]
                : throw new WriterException(ErrorCode.Usage, $"--{name} needs a value", $"Usage: writer {verb} {spec.Usage}"));
            if (!args.Options.TryGetValue(name, out var list)) args.Options[name] = list = [];
            list.Add(value);
        }
        return args;
    }
}
