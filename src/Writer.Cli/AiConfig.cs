using System.Text.Json;
using System.Text.Json.Nodes;
using Writer.Core;

namespace Writer.Cli;

/// <summary>The assistant's model settings. The provider picks the wire format: "anthropic" speaks the Anthropic Messages API,
/// "chatgpt" (用 ChatGPT 登录) the Codex backend's Responses API (ChatCodex.cs), every other provider an OpenAI-compatible
/// /chat/completions. The key goes nowhere but the settings file and requests to BaseUrl; the ChatGPT sign-in's tokens nowhere but
/// the settings file, the sign-in service and the Codex backend.</summary>
public sealed record AiConfig(string Provider, string BaseUrl, string Model, string Key, string Source)
{
    /// <summary>Known providers and their base URLs as each provider's API docs give them (checked 2026-09-23 at
    /// platform.claude.com/docs/en/get-started, developers.openai.com/api/reference, api-docs.deepseek.com,
    /// help.aliyun.com/zh/model-studio/compatibility-of-openai-with-dashscope (the shared dashscope.aliyuncs.com domain
    /// "仍可正常使用" beside per-workspace ones), platform.kimi.com/docs/guide/start-using-kimi-api,
    /// docs.bigmodel.cn/cn/guide/develop/openai/introduction, docs.volcengine.com/docs/82379/1330626,
    /// docs.ollama.com/api/openai-compatibility, lmstudio.ai/docs/developer/openai-compat).
    /// AI_PROVIDERS in ui/prefs.js offers the same list; ui/tests/ai.test.mjs keeps the two in step.</summary>
    public static readonly IReadOnlyDictionary<string, string> Providers = new Dictionary<string, string>
    {
        ["anthropic"] = "https://api.anthropic.com",
        ["openai"] = "https://api.openai.com/v1",
        // not an API: the backend the Codex CLI talks to with a ChatGPT sign-in (ChatGpt.cs); WRITER_CHATGPT_BASE_URL moves it
        ["chatgpt"] = "https://chatgpt.com/backend-api/codex",
        ["deepseek"] = "https://api.deepseek.com",
        ["qwen"] = "https://dashscope.aliyuncs.com/compatible-mode/v1",
        ["kimi"] = "https://api.moonshot.cn/v1",
        ["glm"] = "https://open.bigmodel.cn/api/paas/v4",
        ["doubao"] = "https://ark.cn-beijing.volces.com/api/v3",
        ["ollama"] = "http://localhost:11434/v1",
        ["lmstudio"] = "http://localhost:1234/v1",
        ["custom"] = "",
    };

    public static readonly AiConfig None = new("anthropic", Providers["anthropic"], "", "", "none");

    /// <summary>The model the autocomplete (POST /complete) uses: a fast one of the same provider; "" means the assistant's.</summary>
    public string CompleteModel { get; init; } = "";

    /// <summary>The ChatGPT sign-in (用 ChatGPT 登录), kept whichever provider is in use so that switching back needs no new sign-in.</summary>
    public ChatGptTokens? Account { get; init; }

    public bool OpenAi => Provider != "anthropic";
    /// <summary>用 ChatGPT 登录: the sign-in's tokens instead of a key, the Codex backend's wire format.</summary>
    public bool Codex => Provider == "chatgpt";
    public bool HasKey => Key.Length > 0;
    /// <summary>A ChatGPT sign-in has tokens instead; local servers and custom endpoints may run without a key.</summary>
    public bool NeedsKey => Provider is not ("chatgpt" or "ollama" or "lmstudio" or "custom");
    public bool Usable => Model.Length > 0 && IsHttpUrl(BaseUrl) && (Codex ? Account is not null : HasKey || !NeedsKey);

    /// <summary>A record prints every member; this one must never print the key or the tokens.</summary>
    public override string ToString() => $"{Provider} {Model} at {BaseUrl} ({Source})";

    public static bool IsHttpUrl(string url) =>
        Uri.TryCreate(url, UriKind.Absolute, out var u) && (u.Scheme == "http" || u.Scheme == "https") && u.Host.Length > 0
        && u.UserInfo.Length == 0 && u.Query.Length == 0 && u.Fragment.Length == 0;

    /// <summary>Same scheme, host and port: a stored key may follow the address only within one server.</summary>
    public static bool SameOrigin(string a, string b) =>
        Uri.TryCreate(a, UriKind.Absolute, out var x) && Uri.TryCreate(b, UriKind.Absolute, out var y)
        && Uri.Compare(x, y, UriComponents.SchemeAndServer, UriFormat.Unescaped, StringComparison.OrdinalIgnoreCase) == 0;

    /// <summary>WRITER_AI_PROVIDER, WRITER_AI_BASE_URL, WRITER_AI_KEY, WRITER_AI_MODEL and WRITER_AI_COMPLETE_MODEL, or the older
    /// ANTHROPIC_API_KEY, ANTHROPIC_BASE_URL and WRITER_MODEL. Null when none of them names a provider, an address or a key.</summary>
    public static AiConfig? FromEnvironment(Func<string, string?> env)
    {
        string? Var(string name) => env(name)?.Trim() is { Length: > 0 } v ? v : null;
        var provider = Var("WRITER_AI_PROVIDER")?.ToLowerInvariant();
        var baseUrl = Var("WRITER_AI_BASE_URL");
        var key = Var("WRITER_AI_KEY");
        if (provider is null && baseUrl is null && key is null && Var("ANTHROPIC_API_KEY") is null) return null;
        provider = provider is null ? "anthropic" : Providers.ContainsKey(provider) ? provider : "custom";
        var anthropic = provider == "anthropic";
        if (provider == "chatgpt") // the sign-in comes from the settings file (AiStore.Load); its tokens go to the Codex backend only
            return new AiConfig(provider, ChatGpt.BaseUrl, Var("WRITER_AI_MODEL") ?? ChatGpt.DefaultModel, "", "env") { CompleteModel = Var("WRITER_AI_COMPLETE_MODEL") ?? "" };
        return new AiConfig(provider,
            (baseUrl ?? (anthropic ? Var("ANTHROPIC_BASE_URL") : null) ?? Providers[provider]).TrimEnd('/'),
            Var("WRITER_AI_MODEL") ?? Var("WRITER_MODEL") ?? (anthropic ? Chat.DefaultModel : ""),
            key ?? (anthropic ? Var("ANTHROPIC_API_KEY") : null) ?? "",
            "env") { CompleteModel = Var("WRITER_AI_COMPLETE_MODEL") ?? "" };
    }

    /// <summary>~/Library/Application Support/Writer/ai.json on macOS, %APPDATA%\Writer\ai.json on Windows,
    /// $XDG_CONFIG_HOME/writer/ai.json (default ~/.config) elsewhere.</summary>
    public static string DefaultFile()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (OperatingSystem.IsMacOS()) return Path.Combine(home, "Library", "Application Support", "Writer", "ai.json");
        if (OperatingSystem.IsWindows()) return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Writer", "ai.json");
        var xdg = Environment.GetEnvironmentVariable("XDG_CONFIG_HOME");
        return Path.Combine(string.IsNullOrWhiteSpace(xdg) ? Path.Combine(home, ".config") : xdg, "writer", "ai.json");
    }

    /// <summary>The settings file, or null when there is none or it cannot be read.</summary>
    public static AiConfig? Read(string file)
    {
        try
        {
            if (JsonNode.Parse(File.ReadAllText(file)) is not JsonObject o) return null;
            string S(string name) => o[name] is JsonValue v && v.TryGetValue<string>(out var s) ? s.Trim() : "";
            var provider = S("provider");
            provider = Providers.ContainsKey(provider) ? provider : "custom";
            // the sign-in's tokens only ever go to the Codex backend, whatever address the file names
            var baseUrl = provider == "chatgpt" ? ChatGpt.BaseUrl : S("baseUrl").TrimEnd('/');
            return new AiConfig(provider, baseUrl, S("model"), provider == "chatgpt" ? "" : S("apiKey"), "file")
            {
                CompleteModel = S("completeModel"),
                Account = ChatGptTokens.FromJson(o["chatgpt"]),
            };
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException or JsonException)
        {
            return null;
        }
    }

    /// <summary>Replaces the settings file atomically. On Unix the folder is created 0700 and the file exists only as 0600.</summary>
    public void Write(string file)
    {
        var dir = Path.GetDirectoryName(Path.GetFullPath(file))!;
        var o = new JsonObject { ["provider"] = Provider, ["baseUrl"] = BaseUrl, ["model"] = Model, ["completeModel"] = CompleteModel, ["apiKey"] = Key };
        if (Account is not null) o["chatgpt"] = Account.ToJson();
        var json = o.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
        var tmp = file + "." + Guid.NewGuid().ToString("N") + ".tmp";
        if (OperatingSystem.IsWindows())
        {
            Directory.CreateDirectory(dir);
            File.WriteAllText(tmp, json);
        }
        else
        {
            Directory.CreateDirectory(dir, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            var options = new FileStreamOptions { Mode = FileMode.CreateNew, Access = FileAccess.Write, UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite };
            using var writer = new StreamWriter(new FileStream(tmp, options));
            writer.Write(json);
        }
        File.Move(tmp, file, overwrite: true);
    }
}

/// <summary>Where a server keeps its model settings: the per-user file, under whatever the environment sets.
/// Tests pass their own file, environment and HTTP handler.</summary>
public sealed class AiStore(string file, Func<string, string?>? env = null, HttpMessageHandler? handler = null)
{
    readonly SemaphoreSlim _refreshing = new(1, 1);

    /// <summary>The port the ChatGPT sign-in waits on for the browser (the Codex CLI's 1455; tests take any free one with 0).</summary>
    public int LoginPort { get; init; } = ChatGpt.Port;

    /// <summary>Opens the sign-in page; false when there is no browser to open it (tests never open one).</summary>
    public Func<string, bool> Browser { get; init; } = ChatGpt.OpenBrowser;

    /// <summary>The HTTP handler behind every model and sign-in request (tests answer them; null is the network).</summary>
    public HttpMessageHandler? Handler => handler;

    /// <summary>What applies: the environment's settings when it has any, else the file's. The ChatGPT sign-in is the file's either way.</summary>
    public AiConfig Load() => AiConfig.FromEnvironment(env ?? Environment.GetEnvironmentVariable) is { } e ? e with { Account = Stored()?.Account } : Stored() ?? AiConfig.None;

    /// <summary>The file's settings alone, or null.</summary>
    public AiConfig? Stored() => AiConfig.Read(file);

    public void Save(AiConfig config) => config.Write(file);

    /// <summary>When the file was last written (a date in 1601 when there is none): another engine may have saved it.</summary>
    public DateTime Stamp() => File.GetLastWriteTimeUtc(file);

    public Chat Chat(AiConfig config) => config.Codex
        ? new("", config.Model, ChatGpt.BaseUrl, handler, openAi: true, config.Provider) { CompleteModel = config.CompleteModel, Account = ChatGptAccess }
        : new(config.Key, config.Model, config.BaseUrl, handler, config.OpenAi, config.Provider) { CompleteModel = config.CompleteModel };

    /// <summary>The ChatGPT sign-in's tokens for the next request, read from the file (the one place they live) and refreshed first
    /// when they lapse within five minutes or when the backend refused <paramref name="refused"/>, the access token it was sent. A refresh
    /// token works once, so a refresh holds a lock that the engines of other windows (the Mac app runs one per folder window) honour
    /// too, and whoever comes second finds the new tokens already saved.</summary>
    public async Task<ChatGptTokens> ChatGptAccess(string? refused, CancellationToken ct)
    {
        var tokens = Stored()?.Account ?? throw SignedOut();
        if (!tokens.Stale && (refused is null || refused != tokens.Access)) return tokens;
        await _refreshing.WaitAsync(ct);
        try
        {
            using var held = await Lock(ct);
            var stored = Stored();
            tokens = stored?.Account ?? throw SignedOut();
            if (!tokens.Stale && (refused is null || refused != tokens.Access)) return tokens; // another engine refreshed them meanwhile
            using var http = new HttpClient(handler ?? new SocketsHttpHandler(), disposeHandler: handler is null) { Timeout = TimeSpan.FromSeconds(30) };
            var fresh = await ChatGpt.Refresh(http, tokens, ct);
            Save(stored! with { Account = fresh });
            return fresh;
        }
        finally
        {
            _refreshing.Release();
        }
    }

    static WriterException SignedOut() => new(ErrorCode.Io, "请先用 ChatGPT 登录", "设置 › AI › 用 ChatGPT 登录");

    /// <summary>ai.lock beside the settings file, held open exclusively (an advisory lock between engines on macOS and Linux).</summary>
    async Task<FileStream> Lock(CancellationToken ct)
    {
        var path = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(file))!, "ai.lock");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        for (var tries = 0; ; tries++)
        {
            try
            {
                return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException) when (tries < 150)
            {
                await Task.Delay(100, ct); // another engine is refreshing: at most a few seconds
            }
        }
    }
}
