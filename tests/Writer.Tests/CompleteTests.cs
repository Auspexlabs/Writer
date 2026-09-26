using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Writer.Cli;

namespace Writer.Tests;

/// <summary>The autocomplete (POST /complete) and the model list (POST /ai/models): the requests they make of each kind of
/// provider, and what the editor gets back.</summary>
public class CompleteTests : IDisposable
{
    const string Key = "sk-secret-0123456789";
    readonly string _dir = Directory.CreateTempSubdirectory("writer-complete").FullName;
    readonly List<Serve> _servers = [];

    /// <summary>A model API: answers each request with reply(n, url, body) and records (url, auth, body; "" for a GET).</summary>
    sealed class Api(Func<int, string, string, (HttpStatusCode Status, string Body)> reply) : HttpMessageHandler
    {
        public List<(string Url, string? Auth, string? AnthropicKey, string Body)> Requests { get; } = [];

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            var body = request.Content is null ? "" : await request.Content.ReadAsStringAsync(ct);
            var url = request.RequestUri!.ToString();
            Requests.Add((url, request.Headers.Authorization?.ToString(), request.Headers.TryGetValues("x-api-key", out var k) ? k.Single() : null, body));
            var (status, text) = reply(Requests.Count - 1, url, body);
            return new HttpResponseMessage(status) { Content = new StringContent(text, Encoding.UTF8, "application/json") };
        }
    }

    HttpClient Start(Api api, object settings)
    {
        var server = new Serve(0, requireToken: true, workspace: _dir, ai: new AiStore(Path.Combine(_dir, "ai.json"), _ => null, api));
        server.Start();
        _servers.Add(server);
        var client = new HttpClient { BaseAddress = new Uri(server.Url) };
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", server.Token);
        Assert.Equal(HttpStatusCode.OK, client.PutAsync("/ai", Body(settings)).Result.StatusCode);
        return client;
    }

    static StringContent Body(object o) => new(JsonSerializer.Serialize(o), Encoding.UTF8, "application/json");

    static async Task<JsonElement> Read(HttpResponseMessage response)
    {
        var text = await response.Content.ReadAsStringAsync();
        Assert.DoesNotContain(Key, text);
        return JsonDocument.Parse(text).RootElement.Clone();
    }

    static string Choice(string content) => JsonSerializer.Serialize(new { choices = new[] { new { index = 0, message = new { role = "assistant", content } } } });

    [Fact]
    public async Task A_completion_asks_the_fast_model_with_thinking_off_and_answers_one_clean_line()
    {
        var api = new Api((_, _, _) => (HttpStatusCode.OK, Choice(" 我们下周再讨论细节。\n第二段不要")));
        var client = Start(api, new { provider = "deepseek", baseUrl = "https://api.deepseek.com", model = "deepseek-v4-pro", completeModel = "deepseek-flash", apiKey = Key });

        var ai = await Read(await client.GetAsync("/ai"));
        Assert.Equal("deepseek-flash", ai.GetProperty("completeModel").GetString());
        Assert.Contains("\"completeModel\": \"deepseek-flash\"", File.ReadAllText(Path.Combine(_dir, "ai.json")));

        var done = await Read(await client.PostAsync("/complete", Body(new { before = "会议的结论是：先上线基础版，", after = "下一段", hint = "the body of a Word document" })));
        Assert.Equal("我们下周再讨论细节。", done.GetProperty("text").GetString());
        var (url, auth, _, sent) = api.Requests.Single();
        Assert.Equal(("https://api.deepseek.com/chat/completions", "Bearer " + Key), (url, auth));
        var body = JsonDocument.Parse(sent).RootElement;
        Assert.Equal("deepseek-flash", body.GetProperty("model").GetString());
        Assert.Equal("disabled", body.GetProperty("thinking").GetProperty("type").GetString());
        Assert.False(body.GetProperty("stream").GetBoolean());
        Assert.False(body.TryGetProperty("tools", out _), "a completion runs no tools");
        Assert.Equal(Assistant.CompleteSystem, body.GetProperty("messages")[0].GetProperty("content").GetString());
        var user = body.GetProperty("messages")[1].GetProperty("content").GetString()!;
        Assert.StartsWith("<where>the body of a Word document</where>\n<text_before_cursor>\n", user);
        Assert.Contains("会议的结论是：先上线基础版，\n</text_before_cursor>", user);
        Assert.Contains("<text_after_cursor>\n下一段", user);
    }

    [Fact]
    public async Task A_provider_that_refuses_the_thinking_switch_gets_the_plain_request_and_OpenAi_its_own_parameters()
    {
        var api = new Api((n, _, _) => n == 0
            ? (HttpStatusCode.BadRequest, """{"error":{"message":"Unrecognized request argument supplied: thinking"}}""")
            : (HttpStatusCode.OK, Choice("the plan")));
        var client = Start(api, new { provider = "kimi", baseUrl = "https://api.moonshot.cn/v1", model = "kimi-k3", apiKey = Key });
        Assert.Equal("the plan", (await Read(await client.PostAsync("/complete", Body(new { before = "Next week we will present " })))).GetProperty("text").GetString());
        Assert.Equal(2, api.Requests.Count);
        var plain = JsonDocument.Parse(api.Requests[1].Body).RootElement;
        Assert.Equal(("kimi-k3", 400), (plain.GetProperty("model").GetString(), plain.GetProperty("max_tokens").GetInt32()));
        Assert.False(plain.TryGetProperty("thinking", out _));

        var openai = new Api((_, _, _) => (HttpStatusCode.OK, Choice("ok")));
        var c2 = Start(openai, new { provider = "openai", baseUrl = "https://api.openai.com/v1", model = "gpt-6-sol", completeModel = "gpt-6-luna", apiKey = Key });
        await c2.PostAsync("/complete", Body(new { before = "Dear team, " }));
        var o = JsonDocument.Parse(openai.Requests.Single().Body).RootElement;
        Assert.Equal(("gpt-6-luna", "none", 400), (o.GetProperty("model").GetString(), o.GetProperty("reasoning_effort").GetString(), o.GetProperty("max_completion_tokens").GetInt32()));
        Assert.False(o.TryGetProperty("max_tokens", out _));

        var qwen = new Api((_, _, _) => (HttpStatusCode.OK, Choice("ok")));
        var c3 = Start(qwen, new { provider = "qwen", baseUrl = "https://dashscope.aliyuncs.com/compatible-mode/v1", model = "qwen3.7-plus", apiKey = Key });
        await c3.PostAsync("/complete", Body(new { before = "天气预报说" }));
        var q = JsonDocument.Parse(qwen.Requests.Single().Body).RootElement;
        Assert.Equal(("qwen3.7-plus", false), (q.GetProperty("model").GetString(), q.GetProperty("enable_thinking").GetBoolean()));
    }

    [Fact]
    public async Task Anthropic_completions_read_the_text_blocks_and_no_model_means_503()
    {
        var api = new Api((_, _, _) => (HttpStatusCode.OK, """{"content":[{"type":"thinking","thinking":"hmm"},{"type":"text","text":"quarterly report"}],"stop_reason":"end_turn"}"""));
        var client = Start(api, new { provider = "anthropic", baseUrl = "https://api.anthropic.com", model = "claude-sonnet-5", completeModel = "claude-haiku-4-5", apiKey = Key });
        Assert.Equal("quarterly report", (await Read(await client.PostAsync("/complete", Body(new { before = "Attached is the " })))).GetProperty("text").GetString());
        var (url, _, key, sent) = api.Requests.Single();
        Assert.Equal(("https://api.anthropic.com/v1/messages", Key), (url, key));
        var body = JsonDocument.Parse(sent).RootElement;
        Assert.Equal(("claude-haiku-4-5", Assistant.CompleteSystem), (body.GetProperty("model").GetString(), body.GetProperty("system").GetString()));
        Assert.False(body.TryGetProperty("tools", out _));

        var none = new Serve(0, requireToken: false, workspace: _dir, ai: new AiStore(Path.Combine(_dir, "none.json"), _ => null, api));
        none.Start();
        _servers.Add(none);
        using var bare = new HttpClient { BaseAddress = new Uri(none.Url) };
        var refused = await bare.PostAsync("/complete", Body(new { before = "x" }));
        Assert.Equal(HttpStatusCode.ServiceUnavailable, refused.StatusCode);
        Assert.Equal(HttpStatusCode.BadRequest, (await client.PostAsync("/complete", Body(new { after = "x" }))).StatusCode);
    }

    [Fact]
    public async Task The_model_list_comes_from_the_provider_with_the_stored_key_and_leaves_out_what_cannot_chat()
    {
        var api = new Api((_, url, _) => url.EndsWith("/v1/models?limit=100", StringComparison.Ordinal)
            ? (HttpStatusCode.OK, """{"data":[{"id":"claude-opus-5-5","display_name":"Claude Opus 5.5"},{"id":"claude-haiku-4-5","display_name":"Claude Haiku 4.5"}],"has_more":false}""")
            : (HttpStatusCode.OK, """{"object":"list","data":[{"id":"gpt-6-sol"},{"id":"text-embedding-4"},{"id":"gpt-6-luna"},{"id":"gpt-image-2"},{"id":"tts-2"},{"id":"gpt-6-sol"}]}"""));
        var client = Start(api, new { provider = "openai", baseUrl = "https://api.openai.com/v1", model = "gpt-6-sol", apiKey = Key });

        var live = await Read(await client.PostAsync("/ai/models", Body(new { })));
        Assert.Equal(["gpt-6-sol", "gpt-6-luna"], live.GetProperty("models").EnumerateArray().Select(m => m.GetProperty("id").GetString()!).ToArray());
        Assert.Equal(("https://api.openai.com/v1/models", "Bearer " + Key, ""), (api.Requests[^1].Url, api.Requests[^1].Auth, api.Requests[^1].Body));

        var typed = await Read(await client.PostAsync("/ai/models", Body(new { provider = "anthropic", baseUrl = "https://api.anthropic.com", apiKey = "sk-ant-typed" })));
        var names = typed.GetProperty("models").EnumerateArray().Select(m => (m.GetProperty("id").GetString(), m.GetProperty("name").GetString())).ToArray();
        Assert.Equal([("claude-opus-5-5", "Claude Opus 5.5"), ("claude-haiku-4-5", "Claude Haiku 4.5")], names);
        Assert.Equal("sk-ant-typed", api.Requests[^1].AnthropicKey);

        var elsewhere = await client.PostAsync("/ai/models", Body(new { provider = "deepseek", baseUrl = "https://api.deepseek.com" }));
        Assert.Equal(HttpStatusCode.BadRequest, elsewhere.StatusCode);
        Assert.Equal("请填写 API Key", (await Read(elsewhere)).GetProperty("error").GetProperty("message").GetString());
        Assert.Equal(2, api.Requests.Count); // the OpenAI key never went to DeepSeek
    }

    [Theory]
    [InlineData("Hello wor", "ld, how are you?", "ld, how are you?")]
    [InlineData("I think we should", "we should go now", " go now")]
    [InlineData("I think we should ", " go now", "go now")]
    [InlineData("会议的结论是", " 先上线", "先上线")]
    [InlineData("He said ", "\"we will ship on Friday\"", "we will ship on Friday")]
    [InlineData("Plans: ", "```\nship it\n```", "ship it")]
    [InlineData("The end.", "\nA new paragraph", "")]
    [InlineData("Almost", "   ", "")]
    [InlineData("Almost", null, "")]
    public void CleanCompletion_keeps_one_line_without_echoes_quotes_or_a_second_space(string before, string? reply, string shown) =>
        Assert.Equal(shown, Assistant.CleanCompletion(reply, before));

    public void Dispose()
    {
        foreach (var server in _servers) server.Dispose();
        Directory.Delete(_dir, true);
    }
}
