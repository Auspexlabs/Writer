using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Writer.Core;
using Writer.Formats;
using Writer.Formats.Html;

namespace Writer.Cli;

/// <summary>`writer serve`: a local HTTP API for a UI; `writer watch`: a live preview on top of it; `writer app`: the bundled UI.
/// Listens on 127.0.0.1 only. Requests carry the session token as a bearer header or an HttpOnly cookie;
/// browsers on other origins are refused unless the origin was allow-listed at start-up.</summary>
public sealed class Serve : IDisposable
{
    // One cookie per port: cookies are shared across ports on 127.0.0.1, and a desktop host may run one engine per folder.
    string CookieName => $"writer_token_{Port}";

    static readonly Dictionary<string, string> ContentTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        [".html"] = "text/html; charset=utf-8", [".js"] = "text/javascript; charset=utf-8", [".mjs"] = "text/javascript; charset=utf-8",
        [".css"] = "text/css; charset=utf-8", [".json"] = "application/json; charset=utf-8", [".md"] = "text/markdown; charset=utf-8",
        [".txt"] = "text/plain; charset=utf-8", [".csv"] = "text/csv; charset=utf-8", [".svg"] = "image/svg+xml", [".png"] = "image/png",
        [".jpg"] = "image/jpeg", [".jpeg"] = "image/jpeg", [".gif"] = "image/gif", [".webp"] = "image/webp", [".ico"] = "image/x-icon",
        [".woff"] = "font/woff", [".woff2"] = "font/woff2", [".ttf"] = "font/ttf", [".pdf"] = "application/pdf",
        [".docx"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        [".xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        [".pptx"] = "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        [".mm"] = "text/xml; charset=utf-8",
    };

    readonly HttpListener _listener = new();
    readonly string? _token;
    readonly HashSet<string> _allowedOrigins;
    readonly CancellationTokenSource _stop = new();
    readonly AiStore? _ai;
    readonly Lock _aiGate = new();
    volatile Live _live;
    readonly int _listDepth;
    string? _bootstrap;

    public int Port { get; }
    public string Url => $"http://127.0.0.1:{Port}";
    public string? Token => _token;
    /// <summary>Folder that relative file paths resolve against and that /files lists.</summary>
    public string Workspace { get; }
    /// <summary>Folder served under /app/, or null when no UI is available.</summary>
    public string? UiDir { get; }
    /// <summary>The desktop apps' drafts folder: new documents live here until their first save. Null in the browser version.</summary>
    public string? Drafts { get; }
    readonly HashSet<string> _granted = new(PathComparer);
    /// <summary>Files opened in this window from outside (Finder, 打开…, 最近使用): listed even when their folder cannot be read.</summary>
    readonly HashSet<string> _listed = new(PathComparer);
    readonly Lock _grantGate = new();
    static readonly StringComparer PathComparer = OperatingSystem.IsLinux() ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
    static readonly StringComparison PathComparison = OperatingSystem.IsLinux() ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    /// <param name="listDepth">How many folder levels below the workspace /files walks (0 = the workspace folder only).</param>
    /// <param name="ai">The model settings behind GET/PUT /ai; without them the assistant is <paramref name="chat"/> alone.</param>
    /// <param name="drafts">The desktop apps' drafts folder (created on demand); null in the browser version.</param>
    public Serve(int port, bool requireToken, IEnumerable<string>? allowedOrigins = null, string? workspace = null, string? uiDir = null, Chat? chat = null, int listDepth = 3, AiStore? ai = null, string? drafts = null)
    {
        _ai = ai;
        _live = ai is null ? new(AiConfig.None, chat, default) : Current(ai);
        _listDepth = Math.Max(0, listDepth);
        Port = port == 0 ? FreePort() : port;
        _token = requireToken ? NewSecret() : null;
        _allowedOrigins = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { $"http://127.0.0.1:{Port}", $"http://localhost:{Port}" };
        foreach (var origin in allowedOrigins ?? []) _allowedOrigins.Add(origin.TrimEnd('/'));
        Workspace = Path.GetFullPath(workspace ?? Directory.GetCurrentDirectory());
        UiDir = uiDir is null ? null : Path.GetFullPath(uiDir);
        Drafts = drafts is null ? null : Path.GetFullPath(drafts);
        if (Drafts is not null) Directory.CreateDirectory(Drafts);
        _listener.Prefixes.Add($"http://127.0.0.1:{Port}/");
    }

    /// <summary>A file the desktop shell let the user choose (its Save dialog): raw reads and writes may reach it.</summary>
    public void Grant(string path) { var full = Path.GetFullPath(path); lock (_grantGate) _granted.Add(full); }
    bool Granted(string full) { lock (_grantGate) return _granted.Contains(full); }

    /// <summary>One line of the parent's grants channel (--grants-stdin). "allow &lt;absolute path&gt;" grants a file the user chose
    /// in the Save dialog. "list &lt;absolute path&gt;" names a file opened in this window: /files lists it even when its folder
    /// cannot be read (the App Store sandbox grants the file, not its folder), and access still follows the workspace rules,
    /// so it may not be replaced the way a Save dialog target may. Anything else is ignored.</summary>
    public bool ApplyGrantLine(string line)
    {
        var verb = line.StartsWith("allow ", StringComparison.Ordinal) ? "allow" : line.StartsWith("list ", StringComparison.Ordinal) ? "list" : null;
        if (verb is null) return false;
        var path = line[(verb.Length + 1)..].Trim();
        if (path.Length == 0 || !Path.IsPathRooted(path)) return false;
        if (verb == "allow") Grant(path);
        else lock (_grantGate) _listed.Add(Path.GetFullPath(path));
        return true;
    }

    /// <summary>The --grants-stdin loop: every line goes to apply; the end of the stream (the desktop shell went away) calls stop.</summary>
    public static void ReadGrants(TextReader input, Func<string, bool> apply, Action stop)
    {
        for (var line = input.ReadLine(); line is not null; line = input.ReadLine()) apply(line);
        stop();
    }

    static string NewSecret() => Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant();

    static int FreePort()
    {
        var socket = new TcpListener(IPAddress.Loopback, 0);
        socket.Start();
        var port = ((IPEndPoint)socket.LocalEndpoint).Port;
        socket.Stop();
        return port;
    }

    /// <summary>The bundled UI: WRITER_UI, a ui folder beside the executable, or the repository's ui folder when running from source.</summary>
    public static string? FindUiDir()
    {
        var candidates = new List<string>();
        if (Environment.GetEnvironmentVariable("WRITER_UI") is { Length: > 0 } env) candidates.Add(env);
        var dir = AppContext.BaseDirectory;
        for (var i = 0; i < 6 && dir is not null; i++)
        {
            candidates.Add(Path.Combine(dir, "ui"));
            dir = Path.GetDirectoryName(dir.TrimEnd(Path.DirectorySeparatorChar));
        }
        candidates.Add(Path.Combine(Directory.GetCurrentDirectory(), "ui"));
        return candidates.FirstOrDefault(c => File.Exists(Path.Combine(c, "index.dc.html")));
    }

    /// <summary>A one-time code that the preview URL carries; the first visit swaps it for a cookie and drops it from the URL.</summary>
    public string Bootstrap() => _bootstrap = NewSecret();

    public void Start()
    {
        try
        {
            _listener.Start();
        }
        catch (HttpListenerException ex)
        {
            throw new WriterException(ErrorCode.Io, $"Cannot listen on port {Port}: {ex.Message}", "Pick another port with --port.");
        }
        _ = AcceptLoop();
    }

    public void Dispose()
    {
        _stop.Cancel();
        _listener.Close();
    }

    /// <summary>Runs until the process is stopped.</summary>
    public Task RunAsync() => Task.Delay(Timeout.Infinite, _stop.Token).ContinueWith(_ => { });

    async Task AcceptLoop()
    {
        while (_listener.IsListening)
        {
            HttpListenerContext context;
            try
            {
                context = await _listener.GetContextAsync();
            }
            catch (Exception) when (_stop.IsCancellationRequested || !_listener.IsListening)
            {
                return;
            }
            _ = Task.Run(() => Handle(context));
        }
    }

    async Task Handle(HttpListenerContext context)
    {
        var request = context.Request;
        var response = context.Response;
        try
        {
            if (!LocalHost(request))
            {
                await Json(response, 403, Error("FORBIDDEN", "Requests must address 127.0.0.1 or localhost", "Use the URL printed at start-up."));
                return;
            }
            var origin = request.Headers["Origin"];
            if (origin is not null)
            {
                if (!_allowedOrigins.Contains(origin.TrimEnd('/')))
                {
                    await Json(response, 403, Error("FORBIDDEN", $"Origin {origin} is not allowed", "Start the server with --allow-origin <origin> to let a web UI on that origin call it."));
                    return;
                }
                response.Headers["Access-Control-Allow-Origin"] = origin;
                response.Headers["Access-Control-Allow-Credentials"] = "true";
                response.Headers["Access-Control-Allow-Headers"] = "Authorization, Content-Type";
                response.Headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, OPTIONS";
                response.Headers["Vary"] = "Origin";
            }
            if (request.HttpMethod == "OPTIONS")
            {
                response.StatusCode = 204;
                return;
            }
            var path = request.Url?.AbsolutePath ?? "/";
            if (await Bootstrapped(request, response, path)) return;
            if (path == "/app" || path.StartsWith("/app/", StringComparison.Ordinal))
            {
                await Static(response, path);
                return;
            }
            var file = Resolve(request.QueryString["file"]);
            if (!Authorized(request, path))
            {
                if (path == "/") await Text(response, 401, "text/html; charset=utf-8", "<!doctype html><p>Open the preview with <code>writer watch &lt;file&gt;</code> or the app with <code>writer app</code>, which sign this browser in.</p>");
                else await Json(response, 401, Error("UNAUTHORIZED", "Missing or wrong token", "Send Authorization: Bearer <token> (printed at start-up), or open the app via writer app."));
                return;
            }
            switch (path)
            {
                case "/":
                    await Text(response, 200, "text/html; charset=utf-8", PreviewPage(request.QueryString["file"] ?? ""));
                    break;
                case "/run" when request.HttpMethod == "POST":
                    if (request.ContentType?.StartsWith("application/json", StringComparison.OrdinalIgnoreCase) != true)
                    {
                        await Json(response, 415, Error("UNSUPPORTED_MEDIA_TYPE", "POST /run needs Content-Type: application/json", "Send {\"command\":\"...\"} as JSON."));
                        break;
                    }
                    await Json(response, 200, RunCommand(await new StreamReader(request.InputStream, Encoding.UTF8).ReadToEndAsync(), Workspace));
                    break;
                case "/files":
                    await Json(response, 200, ListFiles(request.QueryString["drafts"] == "1"));
                    break;
                case "/chat" when request.HttpMethod == "POST":
                    if (request.ContentType?.StartsWith("application/json", StringComparison.OrdinalIgnoreCase) != true)
                    {
                        await Json(response, 415, Error("UNSUPPORTED_MEDIA_TYPE", "POST /chat needs Content-Type: application/json", "Send {\"file\":\"...\",\"messages\":[...]} as JSON."));
                        break;
                    }
                    await ChatStream(response, await new StreamReader(request.InputStream, Encoding.UTF8).ReadToEndAsync());
                    break;
                case "/file" when request.HttpMethod == "GET":
                    await RawFile(response, file);
                    break;
                case "/stat":
                    await Json(response, 200, StatFile(file));
                    break;
                case "/file" when request.HttpMethod == "PUT":
                    await Json(response, 200, await WriteFile(request, file));
                    break;
                case "/file" when request.HttpMethod == "DELETE":
                    await Json(response, 200, DeleteFile(file));
                    break;
                case "/binary":
                    await Binary(response, file, request.QueryString["path"]);
                    break;
                case "/html":
                    await Text(response, 200, "text/html; charset=utf-8", WithDoc(file, doc => HtmlWriter.Render(doc)));
                    break;
                case "/json":
                    var skip = (request.QueryString["skip"] ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToHashSet();
                    await Json(response, 200, WithDoc(file, doc => NodeJson.Serialize(doc.Root, int.MaxValue, skip)));
                    break;
                case "/outline":
                    await Text(response, 200, "text/plain; charset=utf-8", WithDoc(file, doc => Views.Outline(doc.Root)));
                    break;
                case "/text":
                    await Text(response, 200, "text/plain; charset=utf-8", WithDoc(file, doc => Views.Text(doc.Root)));
                    break;
                case "/events":
                    await Events(response, file ?? throw new WriterException(ErrorCode.Usage, "file is required", "GET /events?file=<path>"));
                    break;
                case "/ai" when request.HttpMethod == "GET":
                    AiSettings();
                    await Json(response, 200, AiJson(Ai().Config));
                    break;
                case "/ai" when request.HttpMethod == "PUT":
                    if (await JsonBody(request, response, AiExample) is { } settings) await Json(response, 200, AiJson(SaveAi(settings)));
                    break;
                case "/ai/test" when request.HttpMethod == "POST":
                    if (await JsonBody(request, response, AiExample) is { } trial) await Json(response, 200, await TestAi(trial));
                    break;
                default:
                    await Json(response, 404, Error("NOT_FOUND", "No such endpoint", "Endpoints: POST /run, POST /chat, GET /files, /file, /stat, /binary, /html, /json, /outline, /text, /events, PUT /file, DELETE /file, GET/PUT /ai, POST /ai/test, /app/."));
                    break;
            }
        }
        catch (WriterException ex)
        {
            await Json(response, ex.Code == ErrorCode.FileNotFound ? 404 : ex.Code == ErrorCode.Conflict ? 409 : 400, Runner.ErrorJson(ex));
        }
        catch (Exception ex)
        {
            try
            {
                await Json(response, 500, Runner.ErrorJson(new WriterException(ErrorCode.Internal, ex.Message, "This is a bug in writer.")));
            }
            catch (Exception)
            {
                // the client went away
            }
        }
        finally
        {
            try { response.Close(); } catch (Exception) { }
        }
    }

    /// <summary>Relative paths live in the workspace; absolute paths are taken as given.</summary>
    string? Resolve(string? file)
    {
        if (string.IsNullOrEmpty(file)) return null;
        return Path.IsPathRooted(file) ? file : Path.GetFullPath(Path.Combine(Workspace, file));
    }

    static bool Under(string full, string root) => full.StartsWith(root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, PathComparison);
    /// <summary>How the page sees a path: relative inside the workspace, absolute elsewhere; '/' separators everywhere.</summary>
    string Report(string full) => (Under(full, Workspace) ? Path.GetRelativePath(Workspace, full) : full).Replace('\\', '/');

    /// <summary>Raw file I/O may reach the workspace, the drafts folder and granted files. Below the workspace or drafts root
    /// no path component may be a symbolic link, and the file itself never.</summary>
    string Allowed(string? file, string usage)
    {
        if (file is null) throw new WriterException(ErrorCode.Usage, "file is required", usage);
        file = Path.GetFullPath(file);
        var root = Under(file, Workspace) ? Workspace : Drafts is not null && Under(file, Drafts) ? Drafts : null;
        if (root is null && !Granted(file))
            throw new WriterException(ErrorCode.Validation, "The file must be inside the workspace", $"Use a path relative to {Workspace}.");
        for (var dir = Path.GetDirectoryName(file); root is not null && dir is not null && dir.Length > root.Length; dir = Path.GetDirectoryName(dir))
            if (Directory.Exists(dir) && new DirectoryInfo(dir).LinkTarget is not null)
                throw new WriterException(ErrorCode.Validation, "The path goes through a symbolic link", "Use the real folder.");
        if (File.Exists(file) && new FileInfo(file).LinkTarget is not null)
            throw new WriterException(ErrorCode.Validation, "The file is a symbolic link", "Use the real file.");
        return file;
    }

    /// <summary>Defeats DNS rebinding: the Host header must name the loopback address this server listens on.</summary>
    bool LocalHost(HttpListenerRequest request)
    {
        var host = request.Url?.Host ?? "";
        var port = request.Url?.Port ?? 0;
        return port == Port && host is "127.0.0.1" or "localhost" or "::1";
    }

    /// <summary>Header first; the cookie set by the bootstrap second; a query token only where headers are impossible (event streams).</summary>
    bool Authorized(HttpListenerRequest request, string path)
    {
        if (_token is null) return true;
        var header = request.Headers["Authorization"];
        if (header is not null && header.StartsWith("Bearer ", StringComparison.Ordinal) && Matches(header[7..].Trim())) return true;
        if (request.Cookies[CookieName]?.Value is { } cookie && Matches(cookie)) return true;
        return path == "/events" && request.QueryString["token"] is { } query && Matches(query);
    }

    bool Matches(string candidate) =>
        _token is not null && CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(candidate), Encoding.UTF8.GetBytes(_token));

    /// <summary>A URL carrying the one-time code: swap it for an HttpOnly cookie and redirect to the same URL without it.</summary>
    async Task<bool> Bootstrapped(HttpListenerRequest request, HttpListenerResponse response, string path)
    {
        var code = request.QueryString["auth"];
        if (code is null || _token is null) return false;
        var ok = _bootstrap is not null && CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(code), Encoding.UTF8.GetBytes(_bootstrap));
        _bootstrap = null;
        if (!ok)
        {
            await Json(response, 401, Error("UNAUTHORIZED", "The link has already been used", "Run writer watch or writer app again for a fresh link."));
            return true;
        }
        response.Headers["Set-Cookie"] = $"{CookieName}={_token}; Path=/; HttpOnly; SameSite=Strict";
        var query = string.Join("&", request.QueryString.AllKeys.Where(k => k is not null && k != "auth")
            .Select(k => Uri.EscapeDataString(k!) + "=" + Uri.EscapeDataString(request.QueryString[k] ?? "")));
        response.StatusCode = 302;
        response.RedirectLocation = path + (query.Length > 0 ? "?" + query : "");
        return true;
    }

    async Task Static(HttpListenerResponse response, string path)
    {
        if (UiDir is null)
        {
            await Json(response, 404, Error("NOT_FOUND", "No UI is bundled with this build", "Set WRITER_UI to the ui folder or run from the repository."));
            return;
        }
        var relative = path.Length <= 5 ? "index.dc.html" : Uri.UnescapeDataString(path[5..]);
        if (relative.Length == 0) relative = "index.dc.html";
        var full = Path.GetFullPath(Path.Combine(UiDir, relative));
        if (!full.StartsWith(UiDir + Path.DirectorySeparatorChar, StringComparison.Ordinal) && full != UiDir)
        {
            await Json(response, 403, Error("FORBIDDEN", "Path escapes the UI folder", "Request files under /app/ only."));
            return;
        }
        if (Directory.Exists(full)) full = Path.Combine(full, "index.dc.html");
        if (!File.Exists(full))
        {
            await Json(response, 404, Error("NOT_FOUND", $"{relative} is not part of the UI", "Check the path."));
            return;
        }
        await Bytes(response, 200, ContentTypes.GetValueOrDefault(Path.GetExtension(full), "application/octet-stream"), await File.ReadAllBytesAsync(full));
    }

    const int MaxListed = 500;

    /// <summary>GET /files (the workspace) or GET /files?drafts=1 (the drafts folder, depth 0). A folder that is missing or
    /// unreadable (the App Store sandbox grants a chosen file, not its folder) is skipped rather than failing the request.
    /// The workspace listing also carries every granted file elsewhere, and every file opened in the workspace, that still exists.</summary>
    string ListFiles(bool drafts = false)
    {
        if (drafts && Drafts is null) throw new WriterException(ErrorCode.Usage, "This server has no drafts folder", "Start it with writer app --drafts <dir>.");
        var extensions = Adapters.All.SelectMany(a => a.Extensions).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var entries = new List<(string Full, FileInfo Info)>();
        void Walk(string dir, int depth)
        {
            if (depth > (drafts ? 0 : _listDepth)) return;
            try
            {
                foreach (var f in Directory.EnumerateFiles(dir))
                {
                    var info = new FileInfo(f);
                    if (info.Name.StartsWith('.') || info.Name.StartsWith("~$") || !extensions.Contains(info.Extension)) continue;
                    entries.Add((f, info));
                }
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException or IOException) { /* unreadable folder: skip it */ }
            try
            {
                foreach (var d in Directory.EnumerateDirectories(dir))
                {
                    var name = Path.GetFileName(d);
                    if (name.StartsWith('.') || name.EndsWith("_files") || name is "node_modules" or "bin" or "obj" or "dist") continue;
                    Walk(d, depth + 1);
                }
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException or IOException) { /* unreadable folder: skip it */ }
        }
        Walk(drafts ? Drafts! : Workspace, 0);
        if (!drafts)
        {
            var seen = new HashSet<string>(entries.Select(e => e.Full), PathComparer);
            lock (_grantGate)
                foreach (var g in _granted.Concat(_listed.Where(f => Under(f, Workspace))))
                    if (seen.Add(g) && extensions.Contains(Path.GetExtension(g)) && File.Exists(g))
                        entries.Add((g, new FileInfo(g)));
        }
        entries.Sort((a, b) => b.Info.LastWriteTimeUtc.CompareTo(a.Info.LastWriteTimeUtc));
        if (entries.Count > MaxListed) entries.RemoveRange(MaxListed, entries.Count - MaxListed);
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("workspace", Workspace);
            w.WriteString("version", Runner.Version);
            var chat = Ai().Chat;
            w.WriteBoolean("chat", chat is not null);
            if (chat is not null) w.WriteString("model", chat.Model);
            if (Drafts is not null) w.WriteString("drafts", Drafts.Replace('\\', '/'));
            w.WriteStartArray("files");
            foreach (var (full, info) in entries)
            {
                w.WriteStartObject();
                w.WriteString("path", Report(full));
                w.WriteString("name", info.Name);
                w.WriteString("format", Adapters.CanonicalFormat(info.Extension) ?? info.Extension.TrimStart('.'));
                w.WriteNumber("size", info.Length);
                w.WriteString("modified", info.LastWriteTimeUtc.ToString("o", CultureInfo.InvariantCulture));
                w.WriteEndObject();
            }
            w.WriteEndArray();
            w.WriteEndObject();
        });
    }

    async Task RawFile(HttpListenerResponse response, string? file)
    {
        file = Allowed(file, "GET /file?file=<relative path>");
        if (!File.Exists(file)) throw new WriterException(ErrorCode.FileNotFound, $"{file} not found", "Check the path.");
        await Download(response, Path.GetFileName(file), ContentTypes.GetValueOrDefault(Path.GetExtension(file), "application/octet-stream"), await File.ReadAllBytesAsync(file));
    }

    /// <summary>Milliseconds since the epoch, the unit the page tracks a file's mtime in and compares after a poll or a save.</summary>
    static long MtimeMs(string file) => new DateTimeOffset(File.GetLastWriteTimeUtc(file)).ToUnixTimeMilliseconds();

    /// <summary>GET /stat?file=path: {path, mtime, size} — cheap enough to poll every couple of seconds to notice another
    /// program (the MCP server, an agent's CLI calls) changing the file outside this window.</summary>
    string StatFile(string? file)
    {
        file = Allowed(file, "GET /stat?file=<path>");
        if (!File.Exists(file)) throw new WriterException(ErrorCode.FileNotFound, $"{file} not found", "Check the path.");
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("path", Report(file));
            w.WriteNumber("mtime", MtimeMs(file));
            w.WriteNumber("size", new FileInfo(file).Length);
            w.WriteEndObject();
        });
    }

    /// <summary>Bytes the browser must never interpret as a page on this origin: opaque unless a raster image, sandboxed, no sniffing.</summary>
    static async Task Download(HttpListenerResponse response, string name, string contentType, byte[] bytes)
    {
        var media = contentType.Split(';')[0].Trim().ToLowerInvariant();
        if (media is not ("image/png" or "image/jpeg" or "image/gif" or "image/webp" or "image/bmp" or "application/pdf" or "text/markdown" or "text/plain" or "text/csv"
            or "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            or "application/vnd.openxmlformats-officedocument.presentationml.presentation"))
            contentType = "application/octet-stream";
        var ascii = new string(name.Select(c => c is > ' ' and < (char)127 && c is not '"' and not '\\' ? c : '_').ToArray());
        response.Headers["Content-Disposition"] = "attachment; filename=\"" + ascii + "\"; filename*=UTF-8''" + Uri.EscapeDataString(name);
        response.Headers["X-Content-Type-Options"] = "nosniff";
        response.Headers["Content-Security-Policy"] = "sandbox; default-src 'none'";
        await Bytes(response, 200, contentType, bytes);
    }

    /// <summary>PUT /file?file=path writes the body; with &amp;from=old it moves that file there, or copies it with &amp;keep=1 (另存为).
    /// A target outside the workspace and drafts must be granted, or sit beside a granted source with the same extension.</summary>
    async Task<string> WriteFile(HttpListenerRequest request, string? file)
    {
        long size;
        if (Resolve(request.QueryString["from"]) is { } rawFrom)
        {
            var from = Allowed(rawFrom, "PUT /file?file=<new path>&from=<old path>");
            var to = Path.GetFullPath(file ?? throw new WriterException(ErrorCode.Usage, "file is required", "PUT /file?file=<new path>&from=<old path>"));
            // a new name beside a granted file (renaming it) is fine; an existing file there is someone else's
            var sibling = !File.Exists(to) && Granted(from) && PathComparer.Equals(Path.GetDirectoryName(to), Path.GetDirectoryName(from))
                && PathComparer.Equals(Path.GetExtension(to), Path.GetExtension(from));
            if (sibling) Grant(to);
            to = Allowed(to, "PUT /file?file=<new path>&from=<old path>");
            if (!File.Exists(from)) throw new WriterException(ErrorCode.FileNotFound, $"{from} not found", "Check the path.");
            var replace = File.Exists(to);
            if (replace && !Granted(to)) throw new WriterException(ErrorCode.Validation, $"{to} already exists", "Pick another name.");
            Directory.CreateDirectory(Path.GetDirectoryName(to)!);
            if (request.QueryString["keep"] == "1") Files.ReplaceAtomic(to, s => { using var src = File.OpenRead(from); src.CopyTo(s); });
            else if (replace) { Files.ReplaceAtomic(to, s => { using var src = File.OpenRead(from); src.CopyTo(s); }); File.Delete(from); }
            else File.Move(from, to);
            file = to;
            size = new FileInfo(to).Length;
        }
        else
        {
            file = Allowed(file, "PUT /file?file=<path>");
            // the real guard against overwriting someone else's edit: a save that read the file at one mtime refuses to
            // land on top of a different one, so the page can reload and merge instead of silently clobbering it
            if (request.QueryString["ifMtime"] is { } ifMtime && File.Exists(file)
                && (!long.TryParse(ifMtime, out var want) || MtimeMs(file) != want))
                throw new WriterException(ErrorCode.Conflict, $"{Report(file)} changed on disk since it was read",
                    "GET /stat?file=<path> for the current version and mtime, then reload before saving again.");
            Directory.CreateDirectory(Path.GetDirectoryName(file)!);
            var ms = new MemoryStream();
            await request.InputStream.CopyToAsync(ms);
            Files.ReplaceAtomic(file, s => s.Write(ms.GetBuffer(), 0, (int)ms.Length));
            size = ms.Length;
        }
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("path", Report(file));
            w.WriteNumber("size", size);
            w.WriteNumber("mtime", MtimeMs(file));
            w.WriteEndObject();
        });
    }

    /// <summary>DELETE /file?file=path: 不存储 on a draft. Only the drafts folder.</summary>
    string DeleteFile(string? file)
    {
        file = Allowed(file, "DELETE /file?file=<draft path>");
        if (Drafts is null || !Under(file, Drafts)) throw new WriterException(ErrorCode.Validation, "Only drafts can be deleted", "Delete files in Finder or Explorer.");
        File.Delete(file);
        return NodeJson.Write(w => { w.WriteStartObject(); w.WriteString("deleted", Report(file)); w.WriteEndObject(); });
    }

    static async Task Binary(HttpListenerResponse response, string? file, string? nodePath)
    {
        if (file is null || nodePath is null) throw new WriterException(ErrorCode.Usage, "file and path are required", "GET /binary?file=<path>&path=/body/image[1]");
        using var doc = Files.Open(file);
        var node = PathResolver.Single(doc.Root, nodePath);
        var binary = node.GetBinary() ?? throw new WriterException(ErrorCode.Validation, $"{nodePath} has no binary content", "Ask for an image node.");
        await Download(response, Path.GetFileName(node.GetProps().GetValueOrDefault("src") ?? "image"), binary.ContentType, binary.Data);
    }

    static string WithDoc(string? file, Func<Document, string> render)
    {
        if (string.IsNullOrEmpty(file)) throw new WriterException(ErrorCode.Usage, "file is required", "Add ?file=<path>.");
        using var doc = Files.Open(file);
        return render(doc);
    }

    /// <summary>Body: {"command": "..."} or {"argv": ["view", "a.docx"]}. Returns {"code", "output"} or {"code", "error"}.
    /// Relative file arguments are resolved against the workspace.</summary>
    public static string RunCommand(string body, string? workspace = null)
    {
        string[] argv;
        try
        {
            using var parsed = JsonDocument.Parse(body);
            var root = parsed.RootElement;
            if (root.TryGetProperty("argv", out var array) && array.ValueKind == JsonValueKind.Array)
                argv = array.EnumerateArray().Select(e => e.GetString() ?? "").ToArray();
            else if (root.TryGetProperty("command", out var command) && command.ValueKind == JsonValueKind.String)
                argv = Mcp.Tokenize(command.GetString()!);
            else throw new WriterException(ErrorCode.Usage, "Body must have \"command\" or \"argv\"", "Example: {\"command\":\"view report.docx outline\"}");
        }
        catch (JsonException ex)
        {
            throw new WriterException(ErrorCode.Usage, $"Body is not JSON: {ex.Message}", "Example: {\"command\":\"view report.docx outline\"}");
        }
        var (code, text) = RunArgv(argv, workspace);
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
    }

    /// <summary>Runs one command in-process. A relative file argument is resolved against the workspace; server verbs are refused.
    /// Returns the exit code and what the command printed (stdout on success, the JSON error on failure).</summary>
    public static (int Code, string Text) RunArgv(string[] argv, string? workspace)
    {
        if (argv.Length > 0 && argv[0] == "writer") argv = argv[1..];
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        int code;
        try
        {
            if (argv.Length > 0 && argv[0] is "mcp" or "serve" or "watch" or "app")
                throw new WriterException(ErrorCode.Usage, $"'{argv[0]}' cannot run inside the server", "Run a document command.");
            if (workspace is not null && argv.Length > 1 && argv[0] is not "help" && !Path.IsPathRooted(argv[1]) && !argv[1].StartsWith('-'))
                argv[1] = Path.Combine(workspace, argv[1]);
            code = Runner.Run(argv, stdout, stderr);
        }
        catch (WriterException ex)
        {
            stderr.WriteLine(Runner.ErrorJson(ex));
            code = ex.ExitCode;
        }
        return (code, code == 0 ? stdout.ToString() : stderr.ToString());
    }

    /// <summary>POST /chat: {"file": "a.docx", "messages": [{"role","content"}...], "instructions"?: "...", "selection"?: "..."}
    /// → server-sent events text, tool, done, error. instructions are the user's standing preferences (appended to the system prompt);
    /// selection means "only this part of the document": it is quoted before the new message and the outline is left out.</summary>
    static string? Text(JsonElement root, string name) => root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    async Task ChatStream(HttpListenerResponse response, string body)
    {
        var chat = Ai().Chat;
        if (chat is null)
        {
            await Json(response, 503, Error("NO_MODEL", "No model is set up", "Set one up in Settings › AI (PUT /ai), or start the server with WRITER_AI_PROVIDER, WRITER_AI_KEY and WRITER_AI_MODEL (or ANTHROPIC_API_KEY)."));
            return;
        }
        string? file, instructions, selection;
        JsonElement messages;
        bool web;
        try
        {
            using var parsed = JsonDocument.Parse(body);
            var root = parsed.RootElement;
            file = root.TryGetProperty("file", out var f) && f.ValueKind == JsonValueKind.String ? Resolve(f.GetString()) : null;
            if (!root.TryGetProperty("messages", out var m) || m.ValueKind != JsonValueKind.Array)
                throw new WriterException(ErrorCode.Usage, "Body needs a messages array", "{\"file\":\"a.docx\",\"messages\":[{\"role\":\"user\",\"content\":\"...\"}]}");
            messages = m.Clone();
            instructions = Text(root, "instructions");
            selection = Text(root, "selection");
            // 联网搜索: on unless the settings page sends web:false.
            web = !(root.TryGetProperty("web", out var w) && w.ValueKind == JsonValueKind.False);
        }
        catch (JsonException ex)
        {
            throw new WriterException(ErrorCode.Usage, $"Body is not JSON: {ex.Message}", "Send {\"file\":\"a.docx\",\"messages\":[...]}.");
        }
        if (file is not null && !File.Exists(file)) throw new WriterException(ErrorCode.FileNotFound, $"{file} not found", "Check the path.");
        response.StatusCode = 200;
        response.ContentType = "text/event-stream";
        response.Headers["Cache-Control"] = "no-cache";
        response.SendChunked = true;
        var stream = response.OutputStream;
        var gate = new SemaphoreSlim(1, 1);
        // HttpListener has no "client disconnected" event: a write to an abandoned SSE stream throws, which is the only signal
        // that the browser's AbortController fired. That failure cancels this token, which the Chat loop already checks between
        // steps and tool calls, so a client going away stops further model calls and tool commands (already-run ones stand).
        using var aborted = CancellationTokenSource.CreateLinkedTokenSource(_stop.Token);
        async Task Write(string chunk)
        {
            await gate.WaitAsync(_stop.Token);
            try
            {
                await stream.WriteAsync(Encoding.UTF8.GetBytes(chunk), _stop.Token);
                await stream.FlushAsync(_stop.Token);
            }
            catch (Exception) when (!_stop.IsCancellationRequested)
            {
                aborted.Cancel();
            }
            finally
            {
                gate.Release();
            }
        }
        Task Emit(string name, string json) => Write("event: " + name + "\ndata: " + json.ReplaceLineEndings(" ") + "\n\n");
        // Nothing is written while the model thinks, so an SSE comment every second is what notices 停止 in time: its failed
        // write cancels the model call before the reply can run a tool command.
        var ping = Task.Run(async () =>
        {
            try { while (true) { await Task.Delay(PingEvery, aborted.Token); await Write(": ping\n\n"); } }
            catch (OperationCanceledException) { }
        });
        try
        {
            await chat.RunAsync(file, Workspace, messages, Emit, aborted.Token, instructions, selection, web);
        }
        catch (WriterException ex)
        {
            await Emit("error", NodeJson.Compact(w =>
            {
                w.WriteStartObject();
                w.WriteString("message", ex.Message);
                w.WriteString("hint", ex.Hint);
                w.WriteEndObject();
            }));
        }
        catch (OperationCanceledException)
        {
            // the client went away (or the server is stopping): nothing left to notify
        }
        finally
        {
            aborted.Cancel();
            await ping;
        }
    }

    static readonly TimeSpan PingEvery = TimeSpan.FromSeconds(1);

    const string AiExample = "{\"provider\":\"deepseek\",\"baseUrl\":\"https://api.deepseek.com\",\"model\":\"deepseek-flash\",\"apiKey\":\"sk-...\"}";

    AiStore AiSettings() => _ai ?? throw new WriterException(ErrorCode.Usage, "This server keeps no model settings", "Start it with writer serve or writer app.");

    /// <summary>The settings in force, their assistant, and the settings file's write time when they were read.</summary>
    sealed record Live(AiConfig Config, Chat? Chat, DateTime Stamp);

    static Live Current(AiStore store)
    {
        var stamp = store.Stamp();
        var config = store.Load();
        return new(config, config.Usable ? store.Chat(config) : null, stamp);
    }

    /// <summary>The settings in force. The Mac app runs one engine per folder window, all on the one per-user settings file:
    /// when another engine has saved it since, its settings apply here too.</summary>
    Live Ai()
    {
        var live = _live;
        if (_ai is not { } store || store.Stamp() == live.Stamp) return live;
        lock (_aiGate)
        {
            if (store.Stamp() != _live.Stamp) _live = Current(store);
            return _live;
        }
    }

    /// <summary>The request body, or null after answering 415: like /run and /chat, these routes take JSON only.</summary>
    static async Task<string?> JsonBody(HttpListenerRequest request, HttpListenerResponse response, string example)
    {
        if (request.ContentType?.StartsWith("application/json", StringComparison.OrdinalIgnoreCase) == true)
            return await new StreamReader(request.InputStream, Encoding.UTF8).ReadToEndAsync();
        await Json(response, 415, Error("UNSUPPORTED_MEDIA_TYPE", $"{request.HttpMethod} {request.Url?.AbsolutePath} needs Content-Type: application/json", "Send " + example + " as JSON."));
        return null;
    }

    /// <summary>GET /ai: {provider, baseUrl, model, hasKey, source}; source is env, file or none. Never the key.</summary>
    static string AiJson(AiConfig c) => NodeJson.Write(w =>
    {
        w.WriteStartObject();
        w.WriteString("provider", c.Provider);
        w.WriteString("baseUrl", c.BaseUrl);
        w.WriteString("model", c.Model);
        w.WriteBoolean("hasKey", c.HasKey);
        w.WriteString("source", c.Source);
        w.WriteEndObject();
    });

    /// <summary>{provider, baseUrl, model, apiKey?} checked. Without apiKey the key of <paramref name="keep"/> is used, but only while the
    /// address stays on its server, so a stored key never travels to another one; "" clears it.</summary>
    static AiConfig ParseAi(JsonElement root, AiConfig? keep)
    {
        if (root.ValueKind != JsonValueKind.Object) throw new WriterException(ErrorCode.Usage, "Body must be a JSON object", AiExample);
        var provider = Text(root, "provider")?.Trim() ?? "";
        if (!AiConfig.Providers.ContainsKey(provider))
            throw new WriterException(ErrorCode.Validation, $"不认识的服务商「{provider}」", "可选：" + string.Join("、", AiConfig.Providers.Keys));
        var baseUrl = (Text(root, "baseUrl") ?? "").Trim().TrimEnd('/');
        if (!AiConfig.IsHttpUrl(baseUrl)) throw new WriterException(ErrorCode.Validation, "接口地址要以 http:// 或 https:// 开头", "例如 https://api.deepseek.com");
        var model = (Text(root, "model") ?? "").Trim();
        if (model.Length == 0) throw new WriterException(ErrorCode.Validation, "请填写模型", "填服务商文档里的模型名。");
        var key = Text(root, "apiKey") is { } typed ? typed.Trim() : keep is not null && AiConfig.SameOrigin(keep.BaseUrl, baseUrl) ? keep.Key : "";
        // a key travels in a request header: a space, line break or full-width character pasted along would break the request
        if (key.Any(c => c is < '!' or > '~')) throw new WriterException(ErrorCode.Validation, "API Key 里有空格、换行或中文字符", "重新复制服务商给的 Key 再粘贴。");
        return new AiConfig(provider, baseUrl, model, key, "file");
    }

    static JsonElement ParseBody(string body)
    {
        try
        {
            using var parsed = JsonDocument.Parse(body);
            return parsed.RootElement.Clone();
        }
        catch (JsonException ex)
        {
            throw new WriterException(ErrorCode.Usage, $"Body is not JSON: {ex.Message}", "Send " + AiExample + ".");
        }
    }

    /// <summary>PUT /ai: writes the settings file and swaps the live assistant without a restart. Environment variables still win,
    /// so the answer (the settings that apply now) can still say source "env".</summary>
    AiConfig SaveAi(string body)
    {
        var store = AiSettings();
        var root = ParseBody(body);
        lock (_aiGate)
        {
            store.Save(ParseAi(root, store.Stored()));
            return (_live = Current(store)).Config;
        }
    }

    /// <summary>POST /ai/test: one small request with the posted settings (the live ones when the body names no provider) → {ok, error?}.</summary>
    async Task<string> TestAi(string body)
    {
        var store = AiSettings();
        var root = ParseBody(body);
        var live = Ai().Config;
        var config = root.ValueKind == JsonValueKind.Object && !root.TryGetProperty("provider", out _) ? live : ParseAi(root, live);
        var error = config.Usable ? await store.Chat(config).TestAsync(_stop.Token)
            : config.Model.Length == 0 ? "请填写模型" : !AiConfig.IsHttpUrl(config.BaseUrl) ? "接口地址不对" : "请填写 API Key";
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteBoolean("ok", error is null);
            if (error is not null) w.WriteString("error", error);
            w.WriteEndObject();
        });
    }

    static string ErrorPart(string stderr)
    {
        try
        {
            using var parsed = JsonDocument.Parse(stderr);
            return parsed.RootElement.TryGetProperty("error", out var e) ? e.GetRawText() : "{}";
        }
        catch (JsonException)
        {
            return "{}";
        }
    }

    async Task Events(HttpListenerResponse response, string file)
    {
        var full = Path.GetFullPath(file);
        if (!File.Exists(full)) throw new WriterException(ErrorCode.FileNotFound, $"{file} not found", "Check the path.");
        response.StatusCode = 200;
        response.ContentType = "text/event-stream";
        response.Headers["Cache-Control"] = "no-cache";
        response.SendChunked = true;
        var changed = new SemaphoreSlim(0);
        using var watcher = new FileSystemWatcher(Path.GetDirectoryName(full)!, Path.GetFileName(full))
        {
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName | NotifyFilters.CreationTime,
            EnableRaisingEvents = true,
        };
        FileSystemEventHandler onChange = (_, _) => changed.Release();
        watcher.Changed += onChange;
        watcher.Created += onChange;
        watcher.Renamed += (_, _) => changed.Release();
        var stream = response.OutputStream;
        await stream.WriteAsync(Encoding.UTF8.GetBytes("event: ready\ndata: " + JsonEncoded(full) + "\n\n"));
        await stream.FlushAsync();
        while (!_stop.IsCancellationRequested)
        {
            var fired = await changed.WaitAsync(TimeSpan.FromSeconds(15), _stop.Token);
            if (fired)
            {
                await Task.Delay(150, _stop.Token);
                while (changed.CurrentCount > 0) await changed.WaitAsync(_stop.Token);
            }
            var message = fired ? "event: change\ndata: " + JsonEncoded(full) + "\n\n" : ": keepalive\n\n";
            await stream.WriteAsync(Encoding.UTF8.GetBytes(message), _stop.Token);
            await stream.FlushAsync(_stop.Token);
        }
    }

    static string JsonEncoded(string s) => NodeJson.Compact(w => w.WriteStringValue(s));

    static string Error(string code, string message, string hint) =>
        Runner.ErrorJson(new WriterException(ErrorCode.Usage, message, hint)).Replace("\"USAGE\"", "\"" + code + "\"");

    static string PreviewPage(string file) =>
        $$"""
        <!doctype html>
        <html><head><meta charset="utf-8"><title>writer preview</title>
        <style>body{margin:0;font-family:-apple-system,"Segoe UI",Helvetica,Arial,sans-serif;background:#ddd}
        header{padding:8px 16px;background:#222;color:#eee;font-size:13px;display:flex;gap:16px;align-items:center}
        header span.ok{color:#7c6}iframe{border:0;width:100vw;height:calc(100vh - 36px);background:#fff}</style></head>
        <body><header><strong>writer</strong><span id="file"></span><span id="status" class="ok">connecting…</span></header>
        <iframe id="view" sandbox=""></iframe>
        <script>
        const file = new URLSearchParams(location.search).get('file') || {{JsonEncoded(file)}};
        const q = '?file=' + encodeURIComponent(file);
        document.getElementById('file').textContent = file;
        async function load() {
          const r = await fetch('/html' + q, { credentials: 'same-origin' });
          document.getElementById('view').srcdoc = await r.text();
          document.getElementById('status').textContent = r.ok ? 'live · ' + new Date().toLocaleTimeString() : 'error ' + r.status;
        }
        load();
        const es = new EventSource('/events' + q);
        es.addEventListener('change', load);
        es.onerror = () => { document.getElementById('status').textContent = 'disconnected'; };
        </script></body></html>
        """;

    static async Task Json(HttpListenerResponse response, int status, string json) => await Text(response, status, "application/json; charset=utf-8", json);

    static async Task Text(HttpListenerResponse response, int status, string contentType, string text) =>
        await Bytes(response, status, contentType, Encoding.UTF8.GetBytes(text));

    static async Task Bytes(HttpListenerResponse response, int status, string contentType, byte[] bytes)
    {
        response.StatusCode = status;
        response.ContentType = contentType;
        response.ContentLength64 = bytes.Length;
        await response.OutputStream.WriteAsync(bytes);
    }

    /// <summary>Opens a URL in the default browser; failures are silent because the URL is printed anyway.</summary>
    public static void OpenBrowser(string url)
    {
        try
        {
            if (OperatingSystem.IsMacOS()) Process.Start("open", url);
            else if (OperatingSystem.IsWindows()) Process.Start(new ProcessStartInfo("cmd", $"/c start \"\" \"{url}\"") { CreateNoWindow = true });
            else Process.Start("xdg-open", url);
        }
        catch (Exception)
        {
            // no browser available
        }
    }
}
