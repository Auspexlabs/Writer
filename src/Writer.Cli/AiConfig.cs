using System.Text.Json;
using System.Text.Json.Nodes;

namespace Writer.Cli;

/// <summary>The assistant's model settings. The provider picks the wire format: "anthropic" speaks the Anthropic Messages API,
/// every other provider an OpenAI-compatible /chat/completions. The key goes nowhere but the settings file and requests to BaseUrl.</summary>
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

    public bool OpenAi => Provider != "anthropic";
    public bool HasKey => Key.Length > 0;
    /// <summary>Local servers and custom endpoints may run without a key.</summary>
    public bool NeedsKey => Provider is not ("ollama" or "lmstudio" or "custom");
    public bool Usable => Model.Length > 0 && IsHttpUrl(BaseUrl) && (HasKey || !NeedsKey);

    /// <summary>A record prints every member; this one must never print the key.</summary>
    public override string ToString() => $"{Provider} {Model} at {BaseUrl} ({Source})";

    public static bool IsHttpUrl(string url) =>
        Uri.TryCreate(url, UriKind.Absolute, out var u) && (u.Scheme == "http" || u.Scheme == "https") && u.Host.Length > 0
        && u.UserInfo.Length == 0 && u.Query.Length == 0 && u.Fragment.Length == 0;

    /// <summary>Same scheme, host and port: a stored key may follow the address only within one server.</summary>
    public static bool SameOrigin(string a, string b) =>
        Uri.TryCreate(a, UriKind.Absolute, out var x) && Uri.TryCreate(b, UriKind.Absolute, out var y)
        && Uri.Compare(x, y, UriComponents.SchemeAndServer, UriFormat.Unescaped, StringComparison.OrdinalIgnoreCase) == 0;

    /// <summary>WRITER_AI_PROVIDER, WRITER_AI_BASE_URL, WRITER_AI_KEY and WRITER_AI_MODEL, or the older ANTHROPIC_API_KEY,
    /// ANTHROPIC_BASE_URL and WRITER_MODEL. Null when none of them names a provider, an address or a key.</summary>
    public static AiConfig? FromEnvironment(Func<string, string?> env)
    {
        string? Var(string name) => env(name)?.Trim() is { Length: > 0 } v ? v : null;
        var provider = Var("WRITER_AI_PROVIDER")?.ToLowerInvariant();
        var baseUrl = Var("WRITER_AI_BASE_URL");
        var key = Var("WRITER_AI_KEY");
        if (provider is null && baseUrl is null && key is null && Var("ANTHROPIC_API_KEY") is null) return null;
        provider = provider is null ? "anthropic" : Providers.ContainsKey(provider) ? provider : "custom";
        var anthropic = provider == "anthropic";
        return new AiConfig(provider,
            (baseUrl ?? (anthropic ? Var("ANTHROPIC_BASE_URL") : null) ?? Providers[provider]).TrimEnd('/'),
            Var("WRITER_AI_MODEL") ?? Var("WRITER_MODEL") ?? (anthropic ? Chat.DefaultModel : ""),
            key ?? (anthropic ? Var("ANTHROPIC_API_KEY") : null) ?? "",
            "env");
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
            return new AiConfig(Providers.ContainsKey(provider) ? provider : "custom", S("baseUrl").TrimEnd('/'), S("model"), S("apiKey"), "file");
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
        var json = new JsonObject { ["provider"] = Provider, ["baseUrl"] = BaseUrl, ["model"] = Model, ["apiKey"] = Key }
            .ToJsonString(new JsonSerializerOptions { WriteIndented = true });
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
    /// <summary>What applies: the environment's settings when it has any, else the file's.</summary>
    public AiConfig Load() => AiConfig.FromEnvironment(env ?? Environment.GetEnvironmentVariable) ?? Stored() ?? AiConfig.None;

    /// <summary>The file's settings alone, or null.</summary>
    public AiConfig? Stored() => AiConfig.Read(file);

    public void Save(AiConfig config) => config.Write(file);

    /// <summary>When the file was last written (a date in 1601 when there is none): another engine may have saved it.</summary>
    public DateTime Stamp() => File.GetLastWriteTimeUtc(file);

    public Chat Chat(AiConfig config) => new(config.Key, config.Model, config.BaseUrl, handler, config.OpenAi);
}
