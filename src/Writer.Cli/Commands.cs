using Writer.Core;
using Writer.Formats;
using Writer.Formats.Html;

namespace Writer.Cli;

/// <summary>One method per verb. Each returns the text to print.</summary>
static class Commands
{
    public static string Execute(Args a) => a.Verb switch
    {
        "help" => Help.Render(a.Positional, a.Flag("json")),
        "version" => Runner.Version + "\n",
        "create" => Create(a),
        "get" => Get(a),
        "query" => Query(a),
        "add" => Add(a),
        "set" => Set(a),
        "remove" => Remove(a),
        "move" => Move(a),
        "copy" => Copy(a),
        "view" => View(a),
        "export" => Export(a),
        "mcp" => ServeMcp(),
        "serve" => ServeHttp(a, null, app: false),
        "watch" => ServeHttp(a, a.Need(0, "file"), app: false),
        "app" => ServeHttp(a, null, app: true),
        _ => throw new WriterException(ErrorCode.Usage, $"Unknown command '{a.Verb}'", "Run 'writer help'."),
    };

    static string Create(Args a)
    {
        var file = a.Need(0, "file");
        var adapter = Adapters.ForPath(file);
        if (File.Exists(file))
            throw new WriterException(ErrorCode.Validation, $"{file} already exists", "Choose another name, or remove the file first.");
        using var doc = a.Opt("from") is { } template ? Files.Open(template) : adapter.Create();
        if (doc.Format != adapter.Format)
            throw new WriterException(ErrorCode.Validation, $"--from must be a .{adapter.Format} file", $"Both files must share a format; the template is .{doc.Format}.");
        Files.SaveAtomic(doc, file);
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("file", file);
            w.WriteString("format", doc.Format);
            w.WriteEndObject();
        }) + "\n";
    }

    static string Get(Args a)
    {
        using var doc = Files.Open(a.Need(0, "file"));
        var node = PathResolver.Single(doc.Root, a.Positional.ElementAtOrDefault(1) ?? "/");
        if (a.Flag("raw")) return node.GetRaw() + "\n";
        return NodeJson.Serialize(node, a.Int("depth", 1)) + "\n";
    }

    static string Query(Args a)
    {
        using var doc = Files.Open(a.Need(0, "file"));
        return NodeJson.Summaries(PathResolver.Query(doc.Root, a.Need(1, "path"))) + "\n";
    }

    static string Add(Args a)
    {
        var file = a.Need(0, "file");
        using var doc = Files.Open(file);
        Files.EnsureWritable(doc);
        var parent = PathResolver.Single(doc.Root, a.Need(1, "parent path"));
        var raw = a.Opt("raw");
        var kind = a.Opt("type") ?? (raw is not null ? null : throw new WriterException(ErrorCode.Usage, "--type is required",
            $"Elements that go under {parent.Kind}: {string.Join(", ", Registry.ForFormat(doc.Format).Where(k => k.Parents.Contains(parent.Kind)).Select(k => k.Name))}."));
        var node = raw is not null ? Mutations.AddRaw(parent, raw, Position(a, doc.Root, parent)) : Mutations.Add(parent, kind!, Props(a), Position(a, doc.Root, parent));
        Files.SaveAtomic(doc, file);
        return NodeJson.Serialize(node, 1) + "\n";
    }

    static string Set(Args a)
    {
        var file = a.Need(0, "file");
        using var doc = Files.Open(file);
        Files.EnsureWritable(doc);
        var path = a.Need(1, "path");
        var props = Props(a);
        var raw = a.Opt("raw");
        if (props.Count == 0 && raw is null)
            throw new WriterException(ErrorCode.Usage, "Nothing to set", "Add --prop name=value (see 'writer help <format> <element>') or --raw <xml>.");
        var results = new List<Node>();
        foreach (var target in Targets(doc.Root, path, a.Flag("all")))
        {
            var node = target;
            if (raw is not null)
            {
                node.SetRaw(raw);
                node = PathResolver.Single(doc.Root, target.Path);
            }
            results.Add(props.Count > 0 ? Mutations.Set(node, props) : Mutations.Refresh(node));
        }
        Files.SaveAtomic(doc, file);
        return (a.Flag("all") ? NodeJson.Summaries(results) : NodeJson.Serialize(results[0], 1)) + "\n";
    }

    static string Remove(Args a)
    {
        var file = a.Need(0, "file");
        using var doc = Files.Open(file);
        Files.EnsureWritable(doc);
        var targets = Targets(doc.Root, a.Need(1, "path"), a.Flag("all"));
        var removed = targets.Select(t => t.Path).ToList();
        foreach (var target in targets.Reverse()) target.Remove();
        Files.SaveAtomic(doc, file);
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteStartArray("removed");
            foreach (var p in removed) w.WriteStringValue(p);
            w.WriteEndArray();
            w.WriteEndObject();
        }) + "\n";
    }

    static string Move(Args a)
    {
        var file = a.Need(0, "file");
        using var doc = Files.Open(file);
        Files.EnsureWritable(doc);
        var node = PathResolver.Single(doc.Root, a.Need(1, "path"));
        var to = a.Opt("to") ?? throw new WriterException(ErrorCode.Usage, "--to is required", "Example: --to /body --index 1");
        var target = PathResolver.Single(doc.Root, to);
        var moved = Mutations.Move(node, target, Position(a, doc.Root, target));
        Files.SaveAtomic(doc, file);
        return NodeJson.Serialize(moved, 1) + "\n";
    }

    static string Copy(Args a)
    {
        var file = a.Need(0, "file");
        using var doc = Files.Open(file);
        Files.EnsureWritable(doc);
        var node = PathResolver.Single(doc.Root, a.Need(1, "path"));
        var to = a.Opt("to") ?? throw new WriterException(ErrorCode.Usage, "--to is required", "Example: --to /slide[2], or --to / --after /slide[1] for a slide");
        var target = PathResolver.Single(doc.Root, to);
        var copy = Mutations.Copy(node, target, Position(a, doc.Root, target));
        Files.SaveAtomic(doc, file);
        return NodeJson.Serialize(copy, 1) + "\n";
    }

    static string View(Args a)
    {
        using var doc = Files.Open(a.Need(0, "file"));
        var mode = a.Positional.ElementAtOrDefault(1) ?? "outline";
        return mode switch
        {
            "outline" => Views.Outline(doc.Root),
            "text" => Views.Text(doc.Root),
            "html" => HtmlWriter.Render(doc),
            "json" => NodeJson.Serialize(doc.Root, int.MaxValue) + "\n",
            _ => throw new WriterException(ErrorCode.Usage, $"Unknown view '{mode}'", "Views: outline, text, html, json."),
        };
    }

    static string ServeMcp()
    {
        Mcp.ServeAsync().GetAwaiter().GetResult();
        return "";
    }

    static string ServeHttp(Args a, string? watchFile, bool app)
    {
        if (watchFile is not null) using (Files.Open(watchFile)) { }
        var ui = a.Opt("ui") ?? Serve.FindUiDir();
        if (app && ui is null)
            throw new WriterException(ErrorCode.FileNotFound, "The editor UI was not found", "Set WRITER_UI to the ui folder, or pass --ui <folder>.");
        var ai = new AiStore(AiConfig.DefaultFile());
        var model = ai.Load();
        var server = new Serve(a.Int("port", watchFile is null && !app ? 26315 : 0), requireToken: !a.Flag("no-token"), a.All("allow-origin"), a.Opt("dir"), ui, listDepth: a.Int("list-depth", 3), ai: ai, drafts: a.Opt("drafts"));
        server.Start();
        if (a.Flag("grants-stdin")) new Thread(() => Serve.ReadGrants(Console.In, server.ApplyGrantLine, server.Dispose)) { IsBackground = true, Name = "grants" }.Start();
        var url = server.Url;
        if (watchFile is not null || app)
        {
            var link = app ? url + "/app/index.dc.html" : url + "/?file=" + Uri.EscapeDataString(Path.GetFullPath(watchFile!));
            var launchLink = server.Token is null ? link : link + "?auth=" + server.Bootstrap();
            if (!a.Flag("no-browser")) Serve.OpenBrowser(launchLink);
            // A desktop host needs the one-time bootstrap URL to authenticate its embedded webview.
            // The browser mode keeps the token out of the printed URL as before.
            url = a.Flag("no-browser") ? launchLink : link;
        }
        Console.Error.WriteLine(NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("url", url);
            w.WriteString("workspace", server.Workspace);
            if (server.Token is not null && watchFile is null && !app) w.WriteString("token", server.Token);
            w.WriteString("assistant", model.Usable ? $"{model.Model} ({model.Provider}, from {model.Source})"
                : "off (set it up in Settings › AI, or set WRITER_AI_PROVIDER, WRITER_AI_KEY and WRITER_AI_MODEL)");
            w.WriteEndObject();
        }));
        server.RunAsync().GetAwaiter().GetResult();
        return "";
    }

    static string Export(Args a)
    {
        var file = a.Need(0, "file");
        var to = a.Opt("to") ?? throw new WriterException(ErrorCode.Usage, "--to is required", "Example: --to report.md");
        using var doc = Files.Open(file);
        var ext = Path.GetExtension(to).ToLowerInvariant();
        var warnings = new List<string>();
        switch (ext)
        {
            case ".html" or ".htm":
                Files.WriteTextAtomic(to, HtmlWriter.Render(doc));
                break;
            case ".json":
                Files.WriteTextAtomic(to, NodeJson.Serialize(doc.Root, int.MaxValue) + "\n");
                break;
            default:
                var (target, found) = Exporter.Export(doc, Adapters.ForPath(to), to);
                warnings = found;
                using (target) Files.SaveAtomic(target, to);
                break;
        }
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("file", to);
            w.WriteString("format", ext.TrimStart('.'));
            w.WriteStartArray("warnings");
            foreach (var warning in warnings) w.WriteStringValue(warning);
            w.WriteEndArray();
            w.WriteEndObject();
        }) + "\n";
    }

    static List<KeyValuePair<string, string>> Props(Args a)
    {
        var list = new List<KeyValuePair<string, string>>();
        foreach (var raw in a.All("prop"))
        {
            var eq = raw.IndexOf('=');
            if (eq <= 0)
                throw new WriterException(ErrorCode.Usage, $"--prop expects name=value, got '{raw}'", "Example: --prop text=\"Hello\"");
            list.Add(new(raw[..eq].Trim(), raw[(eq + 1)..]));
        }
        return list;
    }

    static IReadOnlyList<Node> Targets(Node root, string path, bool all)
    {
        if (!all) return [PathResolver.Single(root, path)];
        var nodes = PathResolver.Query(root, path);
        if (nodes.Count == 0)
            throw new WriterException(ErrorCode.PathNotFound, $"Nothing matches {path}", "Run 'writer view <file> outline' to see every path.");
        return nodes;
    }

    /// <summary>--index, --after or --before as a 1-based position among the parent's children; null appends.</summary>
    static int? Position(Args a, Node root, Node parent)
    {
        if (a.Opt("index") is not null) return a.Int("index", 0);
        if (a.Opt("after") is { } after) return SiblingIndex(root, parent, after) + 1;
        if (a.Opt("before") is { } before) return SiblingIndex(root, parent, before);
        return null;
    }

    static int SiblingIndex(Node root, Node parent, string path)
    {
        var sibling = PathResolver.Single(root, path);
        var children = parent.Children;
        for (var i = 0; i < children.Count; i++)
            if (ReferenceEquals(children[i].Anchor, sibling.Anchor)) return i + 1;
        throw new WriterException(ErrorCode.Validation, $"{path} is not a child of {parent.Path}", "--after and --before must name a direct child of the parent.");
    }
}
