using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Writer.Cli;
using static Writer.Tests.ChatTests;

namespace Writer.Tests;

/// <summary>The model settings: GET/PUT /ai, POST /ai/test, the settings file and the environment variables that win over it.</summary>
public class AiTests : IDisposable
{
    const string Key = "sk-secret-0123456789";
    readonly string _dir = Directory.CreateTempSubdirectory("writer-ai").FullName;
    readonly Dictionary<string, string> _env = [];
    readonly List<Serve> _servers = [];

    string SettingsFile => Path.Combine(_dir, "config", "Writer", "ai.json");

    (Serve Server, HttpClient Client) Start(HttpMessageHandler? api = null)
    {
        var server = new Serve(0, requireToken: true, workspace: _dir, ai: new AiStore(SettingsFile, name => _env.GetValueOrDefault(name), api));
        server.Start();
        _servers.Add(server);
        var client = new HttpClient { BaseAddress = new Uri(server.Url) };
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", server.Token);
        return (server, client);
    }

    static StringContent Body(object o) => new(JsonSerializer.Serialize(o), Encoding.UTF8, "application/json");

    static async Task<JsonElement> Read(HttpResponseMessage response)
    {
        var text = await response.Content.ReadAsStringAsync();
        Assert.DoesNotContain(Key, text);
        return JsonDocument.Parse(text).RootElement.Clone();
    }

    [Fact]
    public async Task Put_saves_the_settings_file_privately_and_swaps_the_assistant_while_get_never_shows_the_key()
    {
        var junk = Path.Combine(_dir, "junk.json");
        File.WriteAllText(junk, "not json");
        Assert.Null(AiConfig.Read(junk));
        var api = new FakeOpenAi((HttpStatusCode.OK, Sse(Delta("""{"content":"Hi from DeepSeek"}"""), Delta("{}", "stop"))));
        var (_, client) = Start(api);

        var none = await Read(await client.GetAsync("/ai"));
        Assert.Equal(("anthropic", "https://api.anthropic.com", "", false, "none"), (none.GetProperty("provider").GetString(), none.GetProperty("baseUrl").GetString(), none.GetProperty("model").GetString(), none.GetProperty("hasKey").GetBoolean(), none.GetProperty("source").GetString()));
        Assert.False((await Read(await client.GetAsync("/files"))).GetProperty("chat").GetBoolean());

        var saved = await Read(await client.PutAsync("/ai", Body(new { provider = "deepseek", baseUrl = "https://api.deepseek.com/", model = "deepseek-flash", apiKey = Key })));
        Assert.Equal(("https://api.deepseek.com", true, "file"), (saved.GetProperty("baseUrl").GetString(), saved.GetProperty("hasKey").GetBoolean(), saved.GetProperty("source").GetString()));
        Assert.False(saved.TryGetProperty("apiKey", out _));
        Assert.Contains(Key, File.ReadAllText(SettingsFile));
        if (!OperatingSystem.IsWindows())
        {
            Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(SettingsFile));
            Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute, File.GetUnixFileMode(Path.GetDirectoryName(SettingsFile)!));
        }
        Assert.True((await Read(await client.GetAsync("/ai"))).GetProperty("hasKey").GetBoolean());
        var files = await Read(await client.GetAsync("/files"));
        Assert.Equal((true, "deepseek-flash"), (files.GetProperty("chat").GetBoolean(), files.GetProperty("model").GetString()));

        var chat = await client.PostAsync("/chat", Body(new { messages = new[] { new { role = "user", content = "hi" } } }));
        var stream = await chat.Content.ReadAsStringAsync();
        Assert.Contains("event: text\ndata: {\"text\":\"Hi from DeepSeek\"}", stream);
        Assert.DoesNotContain(Key, stream);
        Assert.Equal(("https://api.deepseek.com/chat/completions", "Bearer " + Key), (api.Requests[0].Url, api.Requests[0].Auth));

        var kept = await Read(await client.PutAsync("/ai", Body(new { provider = "deepseek", baseUrl = "https://api.deepseek.com", model = "deepseek-v4-pro" })));
        Assert.True(kept.GetProperty("hasKey").GetBoolean(), "leaving apiKey out keeps the stored key");
        var moved = await Read(await client.PutAsync("/ai", Body(new { provider = "kimi", baseUrl = "https://api.moonshot.cn/v1", model = "kimi-k3" })));
        Assert.False(moved.GetProperty("hasKey").GetBoolean(), "the stored key never follows the settings to another server");
        Assert.DoesNotContain(Key, File.ReadAllText(SettingsFile));
        Assert.False((await Read(await client.GetAsync("/files"))).GetProperty("chat").GetBoolean());
        await client.PutAsync("/ai", Body(new { provider = "kimi", baseUrl = "https://api.moonshot.cn/v1", model = "kimi-k3", apiKey = Key }));
        var cleared = await Read(await client.PutAsync("/ai", Body(new { provider = "kimi", baseUrl = "https://api.moonshot.cn/v1", model = "kimi-k3", apiKey = "" })));
        Assert.False(cleared.GetProperty("hasKey").GetBoolean());
        var local = await Read(await client.PutAsync("/ai", Body(new { provider = "ollama", baseUrl = "http://localhost:11434/v1", model = "qwen3:8b" })));
        Assert.False(local.GetProperty("hasKey").GetBoolean());
        Assert.True((await Read(await client.GetAsync("/files"))).GetProperty("chat").GetBoolean(), "a local server needs no key");

        foreach (var bad in new object[] { new { provider = "nope", baseUrl = "https://x.test", model = "m" }, new { provider = "openai", baseUrl = "ftp://x.test", model = "m" }, new { provider = "openai", baseUrl = "https://x.test", model = " " } })
            Assert.Equal(HttpStatusCode.BadRequest, (await client.PutAsync("/ai", Body(bad))).StatusCode);
        foreach (var pasted in new[] { "sk-abc def", "sk-abc\ndef", "sk-中文", "sk-abc，" })
        {
            var refused = await client.PutAsync("/ai", Body(new { provider = "openai", baseUrl = "https://x.test", model = "m", apiKey = pasted }));
            Assert.Equal(HttpStatusCode.BadRequest, refused.StatusCode);
            Assert.Equal("API Key 里有空格、换行或中文字符", (await Read(refused)).GetProperty("error").GetProperty("message").GetString());
        }
        Assert.Equal(HttpStatusCode.UnsupportedMediaType, (await client.PutAsync("/ai", new StringContent("{}", Encoding.UTF8, "text/plain"))).StatusCode);
        var anonymous = new HttpRequestMessage(HttpMethod.Get, "/ai");
        anonymous.Headers.Authorization = null;
        using var bare = new HttpClient { BaseAddress = client.BaseAddress };
        Assert.Equal(HttpStatusCode.Unauthorized, (await bare.SendAsync(anonymous)).StatusCode);
    }

    [Fact]
    public async Task A_save_in_one_engine_applies_in_the_others_that_share_the_settings_file()
    {
        var api = new FakeOpenAi((HttpStatusCode.OK, Sse(Delta("""{"content":"ok"}"""), Delta("{}", "stop"))));
        var (_, here) = Start(api);
        var (_, other) = Start(api); // the Mac app runs one engine per folder window
        Assert.False((await Read(await other.GetAsync("/files"))).GetProperty("chat").GetBoolean());

        await here.PutAsync("/ai", Body(new { provider = "deepseek", baseUrl = "https://api.deepseek.com", model = "deepseek-flash", apiKey = Key }));
        var files = await Read(await other.GetAsync("/files"));
        Assert.Equal((true, "deepseek-flash"), (files.GetProperty("chat").GetBoolean(), files.GetProperty("model").GetString()));
        Assert.True((await Read(await other.GetAsync("/ai"))).GetProperty("hasKey").GetBoolean());
        var chat = await other.PostAsync("/chat", Body(new { messages = new[] { new { role = "user", content = "hi" } } }));
        Assert.Contains("event: text\ndata: {\"text\":\"ok\"}", await chat.Content.ReadAsStringAsync());

        await here.PutAsync("/ai", Body(new { provider = "deepseek", baseUrl = "https://api.deepseek.com", model = "deepseek-flash", apiKey = "" }));
        Assert.False((await Read(await other.GetAsync("/files"))).GetProperty("chat").GetBoolean(), "a cleared key is gone everywhere");
    }

    [Fact]
    public async Task Environment_variables_win_over_the_file()
    {
        _env["WRITER_AI_PROVIDER"] = "openai";
        _env["WRITER_AI_KEY"] = Key;
        _env["WRITER_AI_MODEL"] = "gpt-6-luna";
        var (_, client) = Start();

        var env = await Read(await client.GetAsync("/ai"));
        Assert.Equal(("openai", "https://api.openai.com/v1", "gpt-6-luna", true, "env"), (env.GetProperty("provider").GetString(), env.GetProperty("baseUrl").GetString(), env.GetProperty("model").GetString(), env.GetProperty("hasKey").GetBoolean(), env.GetProperty("source").GetString()));
        var saved = await Read(await client.PutAsync("/ai", Body(new { provider = "deepseek", baseUrl = "https://api.deepseek.com", model = "deepseek-flash", apiKey = "sk-file" })));
        Assert.Equal(("env", "gpt-6-luna"), (saved.GetProperty("source").GetString(), saved.GetProperty("model").GetString()));
        Assert.Contains("deepseek-flash", File.ReadAllText(SettingsFile));
        Assert.DoesNotContain(Key, File.ReadAllText(SettingsFile), StringComparison.Ordinal);
        Assert.Equal("gpt-6-luna", (await Read(await client.GetAsync("/files"))).GetProperty("model").GetString());

        string? Legacy(string name) => name switch { "ANTHROPIC_API_KEY" => "sk-ant", "ANTHROPIC_BASE_URL" => "https://proxy.test/", _ => null };
        Assert.Equal(new AiConfig("anthropic", "https://proxy.test", "claude-sonnet-5", "sk-ant", "env"), AiConfig.FromEnvironment(Legacy));
        var other = AiConfig.FromEnvironment(name => name == "WRITER_AI_PROVIDER" ? "DeepSeek" : Legacy(name))!;
        Assert.Equal(("deepseek", "https://api.deepseek.com", ""), (other.Provider, other.BaseUrl, other.Key)); // the Anthropic key stays with Anthropic
        Assert.Null(AiConfig.FromEnvironment(name => name == "WRITER_MODEL" ? "m" : null));
        Assert.DoesNotContain("sk-ant", AiConfig.FromEnvironment(Legacy)!.ToString());
    }

    [Fact]
    public async Task Test_endpoint_makes_one_small_request_and_names_the_problem()
    {
        var replies = new Queue<(HttpStatusCode, string)>([
            (HttpStatusCode.OK, Sse(Delta("""{"content":"ok"}"""), Delta("{}", "stop"))),
            (HttpStatusCode.Unauthorized, """{"error":{"message":"Authentication Fails, Your api key: ****6789 is invalid"}}"""),
            (HttpStatusCode.NotFound, """{"error":{"message":"Not Found"}}"""),
        ]);
        var api = new FakeOpenAi((_, _) => replies.Dequeue());
        var (_, client) = Start(api);
        await client.PutAsync("/ai", Body(new { provider = "deepseek", baseUrl = "https://api.deepseek.com", model = "deepseek-flash", apiKey = Key }));
        async Task<JsonElement> Test(object body) => await Read(await client.PostAsync("/ai/test", Body(body)));

        Assert.True((await Test(new { })).GetProperty("ok").GetBoolean());
        var sent = api.Requests.Single();
        Assert.Equal(("https://api.deepseek.com/chat/completions", "Bearer " + Key), (sent.Url, sent.Auth));
        Assert.Equal("writer", sent.Body.GetProperty("tools")[0].GetProperty("function").GetProperty("name").GetString());

        var form = new { provider = "deepseek", baseUrl = "https://api.deepseek.com", model = "deepseek-flash", apiKey = "sk-wrong-key" };
        var denied = await Test(form);
        Assert.Equal((false, "Key 无效"), (denied.GetProperty("ok").GetBoolean(), denied.GetProperty("error").GetString()));
        Assert.Equal("Bearer sk-wrong-key", api.Requests[^1].Auth);
        Assert.Equal("地址或模型不对", (await Test(new { provider = "deepseek", baseUrl = "https://api.deepseek.com", model = "nope" })).GetProperty("error").GetString());
        Assert.Equal("Bearer " + Key, api.Requests[^1].Auth); // no apiKey: the stored one, same server

        var elsewhere = await Test(new { provider = "openai", baseUrl = "https://api.openai.com/v1", model = "gpt-6-luna" });
        Assert.Equal("请填写 API Key", elsewhere.GetProperty("error").GetString());
        Assert.Equal(3, api.Requests.Count); // the stored key was not sent to another server

        var (_, offline) = Start(new Unreachable());
        await offline.PutAsync("/ai", Body(new { provider = "ollama", baseUrl = "http://localhost:11434/v1", model = "qwen3:8b" }));
        Assert.Equal("连不上", (await Read(await offline.PostAsync("/ai/test", Body(new { })))).GetProperty("error").GetString());
    }

    [Fact]
    public void Settings_live_in_the_per_user_folder_of_each_system()
    {
        var file = AiConfig.DefaultFile();
        if (OperatingSystem.IsMacOS()) Assert.Equal(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library", "Application Support", "Writer", "ai.json"), file);
        else if (OperatingSystem.IsWindows()) Assert.Equal(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Writer", "ai.json"), file);
        else Assert.EndsWith(Path.Combine("writer", "ai.json"), file);
    }

    public void Dispose()
    {
        foreach (var server in _servers) server.Dispose();
        Directory.Delete(_dir, true);
    }
}
