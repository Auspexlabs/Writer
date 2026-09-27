using System.Globalization;
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
using Writer.Cli;
using Writer.Core;
using Writer.Formats;
using Writer.Formats.Html;

namespace Writer.Browser;

/// <summary>The engine inside a web page: what `writer serve` answers over HTTP (Serve.cs), as calls from JavaScript
/// (ui/embed/server.js). The documents live in the page's in-memory file system under <see cref="Workspace"/>; the page puts
/// their bytes in and takes them out. A failure throws an exception whose message is the JSON the server would send, with
/// the HTTP status it would have: {"status":404,"error":{"code","message","hint"}}.</summary>
[System.Text.Json.Serialization.JsonSerializable(typeof(List<string>))]
internal sealed partial class BrowserJson : System.Text.Json.Serialization.JsonSerializerContext;

[SupportedOSPlatform("browser")]
public static partial class Engine
{
    public const string Workspace = "/work";

    public static void Main()
    {
        Directory.CreateDirectory(Workspace);
        Directory.SetCurrentDirectory(Workspace); // a relative --to (export) lands in the workspace too
    }

    /// <summary>POST /run: {"command":"..."} or {"argv":[...]} → {"code":0,"output":"..."} or {"code":n,"error":{...}}.</summary>
    [JSExport]
    public static string Run(string body) => Guard(() =>
    {
        string[] argv;
        try
        {
            using var parsed = JsonDocument.Parse(body);
            var root = parsed.RootElement;
            if (root.TryGetProperty("argv", out var array) && array.ValueKind == JsonValueKind.Array)
                argv = array.EnumerateArray().Select(e => e.GetString() ?? "").ToArray();
            else if (root.TryGetProperty("command", out var command) && command.ValueKind == JsonValueKind.String)
                argv = Args.Tokenize(command.GetString()!);
            else throw new WriterException(ErrorCode.Usage, "Body must have \"command\" or \"argv\"", "Example: {\"command\":\"view report.docx outline\"}");
        }
        catch (JsonException ex)
        {
            throw new WriterException(ErrorCode.Usage, $"Body is not JSON: {ex.Message}", "Example: {\"command\":\"view report.docx outline\"}");
        }
        var (code, text) = RunArgv(argv);
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteNumber("code", code);
            if (code == 0) w.WriteString("output", text);
            else
            {
                w.WritePropertyName("error");
                w.WriteRawValue(ErrorPart(text));
            }
            w.WriteEndObject();
        });
    });

    /// <summary>One command line against the workspace, as Serve.RunArgv runs it: a relative file argument is a document of the
    /// workspace, server verbs are refused. The exit code, and what the command printed (its JSON error when it failed).</summary>
    static (int Code, string Text) RunArgv(string[] argv)
    {
        if (argv.Length > 0 && argv[0] == "writer") argv = argv[1..];
        if (argv.Length > 1 && argv[0] is not "help" && !Path.IsPathRooted(argv[1]) && !argv[1].StartsWith('-'))
            argv[1] = Path.Combine(Workspace, argv[1]);
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        Forget(); // a command reads the file as it is now, and may write it
        var code = Runner.Run(argv, stdout, stderr);
        if (code == 0) Wrote(argv);
        return (code, code == 0 ? stdout.ToString() : stderr.ToString());
    }

    // ---- the assistant (ui/embed/ai.js runs its loop; the site answers the model calls) ----

    /// <summary>The system prompt for the open file (the desktop assistant's: rules, the command reference, the outline, which
    /// outline=false leaves out when the user asked about a selection). file may be empty: no document open.</summary>
    [JSExport]
    public static string ChatSystem(string file, bool outline) => Guard(() =>
        Assistant.SystemPrompt(string.IsNullOrEmpty(file) ? null : Inside(file), Workspace, outline, web: false));

    /// <summary>The editing tools: [{name, description, input_schema}] (writer, batch, plan).</summary>
    [JSExport]
    public static string ChatTools() => Guard(() =>
    {
        var tools = new System.Text.Json.Nodes.JsonArray();
        foreach (var (name, description, schema) in Assistant.EditingTools())
            tools.Add((System.Text.Json.Nodes.JsonNode)new System.Text.Json.Nodes.JsonObject { ["name"] = name, ["description"] = description, ["input_schema"] = schema });
        return tools.ToJsonString();
    });

    /// <summary>Runs one tool call: {display, code, output, wrote}.</summary>
    [JSExport]
    public static string ChatTool(string name, string input) => Guard(() =>
    {
        System.Text.Json.Nodes.JsonObject? args = null;
        try { args = System.Text.Json.Nodes.JsonNode.Parse(string.IsNullOrWhiteSpace(input) ? "{}" : input) as System.Text.Json.Nodes.JsonObject; }
        catch (JsonException) { }
        var (display, code, output, wrote) = Assistant.RunEditingTool(name, args, RunArgv);
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("display", display);
            w.WriteNumber("code", code);
            w.WriteString("output", output.Length > 6000 ? output[..6000] + "\n…(truncated)" : output);
            w.WriteBoolean("wrote", wrote);
            w.WriteEndObject();
        });
    });

    /// <summary>AI 自动补全 in the page: the request the site's model answers, {system, user} (Assistant.CompleteMessage).</summary>
    [JSExport]
    public static string CompletePrompt(string before, string after, string hint) => NodeJson.Write(w =>
    {
        w.WriteStartObject();
        w.WriteString("system", Assistant.CompleteSystem);
        w.WriteString("user", Assistant.CompleteMessage(before, after, hint));
        w.WriteEndObject();
    });

    /// <summary>The model's reply as the editor shows it (Assistant.CleanCompletion).</summary>
    [JSExport]
    public static string CompleteClean(string text, string before) => Assistant.CleanCompletion(text, before);

    /// <summary>GET /files: the documents in the workspace, newest first, and every extension the engine opens.</summary>
    [JSExport]
    public static string Files() => Guard(() =>
    {
        var extensions = Adapters.All.SelectMany(a => a.Extensions).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var entries = Directory.EnumerateFiles(Workspace, "*", SearchOption.AllDirectories)
            .Select(f => new FileInfo(f))
            .Where(i => !i.Name.StartsWith('.') && !i.Name.StartsWith("~$") && extensions.Contains(i.Extension))
            .OrderByDescending(i => i.LastWriteTimeUtc)
            .ToList();
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("workspace", Workspace);
            w.WriteString("version", Runner.Version);
            w.WriteBoolean("chat", false);
            w.WriteStartObject("open");
            foreach (var (ext, editor) in Adapters.OpensAs.OrderBy(x => x.Key, StringComparer.Ordinal)) w.WriteString(ext, editor);
            w.WriteEndObject();
            w.WriteStartArray("files");
            foreach (var info in entries)
            {
                w.WriteStartObject();
                w.WriteString("path", Report(info.FullName));
                w.WriteString("name", info.Name);
                w.WriteString("format", Adapters.CanonicalFormat(info.Extension) ?? info.Extension.TrimStart('.'));
                w.WriteNumber("size", info.Length);
                w.WriteString("modified", info.LastWriteTimeUtc.ToString("o", CultureInfo.InvariantCulture));
                w.WriteEndObject();
            }
            w.WriteEndArray();
            w.WriteEndObject();
        });
    });

    /// <summary>GET /stat?file=: {path, mtime, size}.</summary>
    [JSExport]
    public static string Stat(string file) => Guard(() =>
    {
        var full = Existing(file);
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("path", Report(full));
            w.WriteNumber("mtime", MtimeMs(full));
            w.WriteNumber("size", new FileInfo(full).Length);
            w.WriteEndObject();
        });
    });

    /// <summary>GET /file?file=: the bytes.</summary>
    [JSExport]
    public static byte[] Read(string file) => Guard(() => File.ReadAllBytes(Existing(file)));

    /// <summary>PUT /file?file=: writes the bytes (refusing when ifMtime, the mtime the page read, is no longer the file's).</summary>
    [JSExport]
    public static string Write(string file, byte[] bytes, string? ifMtime) => Guard(() =>
    {
        var full = Inside(file);
        if (!string.IsNullOrEmpty(ifMtime) && File.Exists(full) && (!long.TryParse(ifMtime, out var want) || MtimeMs(full) != want))
            throw new WriterException(ErrorCode.Conflict, $"{Report(full)} changed since it was read", "GET /stat?file=<path> for the current version, then reload before saving again.");
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        Forget(full);
        File.WriteAllBytes(full, bytes);
        Bump(full);
        return Saved(full);
    });

    /// <summary>PUT /file?file=to&amp;from=from: moves the file there, or copies it with keep (另存为).</summary>
    [JSExport]
    public static string Move(string from, string to, bool keep) => Guard(() =>
    {
        var source = Existing(from);
        var target = Inside(to);
        if (File.Exists(target) && !string.Equals(source, target, StringComparison.Ordinal))
            throw new WriterException(ErrorCode.Validation, $"{Report(target)} already exists", "Pick another name.");
        Directory.CreateDirectory(Path.GetDirectoryName(target)!);
        Forget(source); Forget(target);
        if (keep) File.Copy(source, target, overwrite: true);
        else { File.Move(source, target, overwrite: true); Versions.Remove(source); if (!Changed.Contains(source)) Changed.Add(source); }
        Bump(target);
        return Saved(target);
    });

    /// <summary>DELETE /file?file=.</summary>
    [JSExport]
    public static string Delete(string file) => Guard(() =>
    {
        var full = Existing(file);
        Forget(full);
        File.Delete(full);
        Versions.Remove(full);
        if (!Changed.Contains(full)) Changed.Add(full);
        return NodeJson.Write(w => { w.WriteStartObject(); w.WriteString("deleted", Report(full)); w.WriteEndObject(); });
    });

    /// <summary>GET /json?file=&amp;skip=: the document tree (skip: node kinds left out, comma-separated).</summary>
    [JSExport]
    public static string Json(string file, string skip) => Guard(() =>
    {
        var kinds = skip.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToHashSet();
        return NodeJson.SerializeCompact(Doc(Existing(file)).Root, int.MaxValue, kinds);
    });

    /// <summary>GET /html, /outline, /text ?file=: the other views.</summary>
    [JSExport]
    public static string View(string file, string mode) => Guard(() =>
    {
        var doc = Doc(Existing(file));
        return mode switch
        {
            "html" => HtmlWriter.Render(doc),
            "outline" => Views.Outline(doc.Root),
            "text" => Views.Text(doc.Root),
            _ => throw new WriterException(ErrorCode.Usage, $"Unknown view '{mode}'", "Views: html, outline, text."),
        };
    });

    static string? _binaryType;

    /// <summary>GET /binary?file=&amp;path=: the bytes of an image node; <see cref="BinaryType"/> is then their content type.</summary>
    [JSExport]
    public static byte[] Binary(string file, string path) => Guard(() =>
    {
        var node = PathResolver.Single(Doc(Existing(file)).Root, path);
        var binary = node.GetBinary() ?? throw new WriterException(ErrorCode.Validation, $"{path} has no binary content", "Ask for an image node.");
        _binaryType = binary.ContentType;
        return binary.Data;
    });

    [JSExport]
    public static string BinaryType() => _binaryType ?? "application/octet-stream";

    /// <summary>The page's paths: relative inside the workspace, '/' separated.</summary>
    static string Report(string full) => (full.StartsWith(Workspace + "/", StringComparison.Ordinal) ? full[(Workspace.Length + 1)..] : full).Replace('\\', '/');

    /// <summary>A path of the workspace: relative ones resolve against it, and nothing may leave it.</summary>
    static string Inside(string? file)
    {
        if (string.IsNullOrEmpty(file)) throw new WriterException(ErrorCode.Usage, "file is required", "Add ?file=<path>.");
        var full = Path.GetFullPath(Path.IsPathRooted(file) ? file : Path.Combine(Workspace, file));
        if (!full.StartsWith(Workspace + "/", StringComparison.Ordinal))
            throw new WriterException(ErrorCode.Validation, "The file must be inside the workspace", "Use a path relative to the workspace.");
        return full;
    }

    static string Existing(string? file)
    {
        var full = Inside(file);
        if (!File.Exists(full)) throw new WriterException(ErrorCode.FileNotFound, $"{Report(full)} not found", "Check the path.");
        return full;
    }

    // ---- versions: the page's file system keeps modification times to the second, so two saves in one second would look
    // like none. Every write through this class gives the file a new time in milliseconds, later than any before. ----
    static readonly Dictionary<string, long> Versions = new(StringComparer.Ordinal);
    static long _last;

    static long MtimeMs(string file) => Versions.TryGetValue(file, out var v) ? v : new DateTimeOffset(File.GetLastWriteTimeUtc(file)).ToUnixTimeMilliseconds();

    static void Bump(string full)
    {
        Versions[full] = _last = Math.Max(DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(), _last + 1);
        if (!Changed.Contains(full)) Changed.Add(full);
    }

    static readonly List<string> Changed = [];

    /// <summary>The files written since the last call, as the page names them (a JSON array): the page tells the editor and
    /// the embedding site.</summary>
    [JSExport]
    public static string Drain()
    {
        var list = Changed.Select(Report).ToList();
        Changed.Clear();
        return JsonSerializer.Serialize(list, BrowserJson.Default.ListString);
    }

    /// <summary>After a command succeeded: the document it names has a new version unless the verb only reads, and so does
    /// what `export --to` wrote.</summary>
    static void Wrote(string[] argv)
    {
        if (argv.Length < 2 || argv[0] is "help" or "version" or "get" or "query" or "view" or "search") return;
        if (argv[0] != "export") { if (File.Exists(argv[1])) Bump(Path.GetFullPath(argv[1])); return; }
        var to = Array.IndexOf(argv, "--to");
        if (to > 0 && to + 1 < argv.Length && Path.GetFullPath(Path.IsPathRooted(argv[to + 1]) ? argv[to + 1] : Path.Combine(Workspace, argv[to + 1])) is var target && File.Exists(target)) Bump(target);
    }

    // ---- open documents: a Word file with twenty pictures asks for each picture's bytes, and each ask would parse it again.
    // The last few opened stay open until a write through this class, or any command, changes what is on disk. ----
    static readonly List<(string Path, long Version, Document Doc)> Opened = [];

    static Document Doc(string full)
    {
        var version = MtimeMs(full);
        var i = Opened.FindIndex(o => o.Path == full);
        if (i >= 0)
        {
            var hit = Opened[i];
            Opened.RemoveAt(i);
            if (hit.Version == version) { Opened.Add(hit); return hit.Doc; }
            hit.Doc.Dispose();
        }
        var doc = Writer.Cli.Files.Open(full);
        Opened.Add((full, version, doc));
        if (Opened.Count > 4) { Opened[0].Doc.Dispose(); Opened.RemoveAt(0); }
        return doc;
    }

    /// <summary>Closes the open copy of a file, or of every file.</summary>
    static void Forget(string? full = null)
    {
        for (var i = Opened.Count - 1; i >= 0; i--)
            if (full is null || Opened[i].Path == full) { Opened[i].Doc.Dispose(); Opened.RemoveAt(i); }
    }

    static string Saved(string full) => NodeJson.Write(w =>
    {
        w.WriteStartObject();
        w.WriteString("path", Report(full));
        w.WriteNumber("size", new FileInfo(full).Length);
        w.WriteNumber("mtime", MtimeMs(full));
        w.WriteEndObject();
    });

    /// <summary>The error object out of the JSON a failed command printed ({"error":{...}}), or the text as its message.</summary>
    static string ErrorPart(string text)
    {
        try
        {
            using var parsed = JsonDocument.Parse(text);
            if (parsed.RootElement.TryGetProperty("error", out var error)) return error.GetRawText();
        }
        catch (JsonException) { }
        return NodeJson.Write(w => { w.WriteStartObject(); w.WriteString("code", "INTERNAL"); w.WriteString("message", text.Trim()); w.WriteEndObject(); });
    }

    /// <summary>Runs one call; a failure becomes an exception carrying the server's JSON answer and status.</summary>
    static T Guard<T>(Func<T> call)
    {
        try
        {
            return call();
        }
        catch (WriterException ex)
        {
            var status = ex.Code switch { ErrorCode.FileNotFound => 404, ErrorCode.Conflict => 409, _ => 400 };
            throw new InvalidOperationException("{\"status\":" + status + "," + Runner.ErrorJson(ex)[1..]);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new InvalidOperationException("{\"status\":500," + Runner.ErrorJson(new WriterException(ErrorCode.Io, ex.Message, "The page's file store refused it."))[1..]);
        }
    }
}
