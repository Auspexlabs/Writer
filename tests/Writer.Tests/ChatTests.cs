using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Writer.Cli;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class ChatTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("writer-chat").FullName;

    /// <summary>Plays canned model replies in order and records what was sent.</summary>
    sealed class FakeApi(params (HttpStatusCode Status, string Body)[] replies) : HttpMessageHandler
    {
        public List<string> Requests { get; } = [];
        int _n;

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            Assert.Equal("https://fake.test/v1/messages", request.RequestUri!.ToString());
            Assert.Equal("k-1", request.Headers.GetValues("x-api-key").Single());
            Requests.Add(await request.Content!.ReadAsStringAsync(ct));
            var (status, body) = replies[Math.Min(_n++, replies.Length - 1)];
            return new HttpResponseMessage(status) { Content = new StringContent(body, Encoding.UTF8, "application/json") };
        }
    }

    static string Messages(params (string Role, string Content)[] turns) =>
        JsonSerializer.Serialize(turns.Select(t => new { role = t.Role, content = t.Content }));

    [Fact]
    public async Task Runs_tool_calls_until_the_model_stops_and_streams_every_step()
    {
        var file = Path.Combine(_dir, "doc.docx");
        File.WriteAllBytes(file, Docx(P("Old text")));
        var api = new FakeApi(
            (HttpStatusCode.OK, """{"content":[{"type":"text","text":"Let me look."},{"type":"tool_use","id":"t1","name":"writer","input":{"command":"set doc.docx /body/paragraph[1] --prop text=Hi"}}],"stop_reason":"tool_use"}"""),
            (HttpStatusCode.OK, """{"content":[{"type":"text","text":"Done."}],"stop_reason":"end_turn"}"""));
        var chat = new Chat("k-1", "model-x", "https://fake.test/", api);
        var events = new List<(string Name, JsonElement Data)>();
        using var history = JsonDocument.Parse(Messages(("user", "Change the first paragraph to Hi")));

        await chat.RunAsync(file, _dir, history.RootElement, (name, json) => { events.Add((name, JsonDocument.Parse(json).RootElement)); return Task.CompletedTask; }, CancellationToken.None);

        Assert.Equal(["text", "tool", "text", "done"], events.Select(e => e.Name));
        Assert.Equal("Let me look.", events[0].Data.GetProperty("text").GetString());
        Assert.Equal(0, events[1].Data.GetProperty("code").GetInt32());
        Assert.Contains("/body/paragraph[1]", events[1].Data.GetProperty("output").GetString());
        Assert.Equal(2, events[3].Data.GetProperty("steps").GetInt32());
        using var doc = OpenDocx(File.ReadAllBytes(file));
        Assert.Equal("Hi", doc.Root.Children[0].Children[0].Text);

        Assert.Equal(2, api.Requests.Count);
        var first = JsonDocument.Parse(api.Requests[0]).RootElement;
        Assert.Equal("model-x", first.GetProperty("model").GetString());
        Assert.Equal("writer", first.GetProperty("tools")[0].GetProperty("name").GetString());
        var system = first.GetProperty("system").GetString()!;
        Assert.Contains("Open file: doc.docx", system);
        Assert.Contains("/body/paragraph[1]", system);
        Assert.Contains("paragraph (docx)", system);
        var second = JsonDocument.Parse(api.Requests[1]).RootElement.GetProperty("messages");
        Assert.Equal(3, second.GetArrayLength());
        Assert.Equal("tool_use", second[1].GetProperty("content")[1].GetProperty("type").GetString());
        var result = second[2].GetProperty("content")[0];
        Assert.Equal(("tool_result", "t1", false), (result.GetProperty("type").GetString(), result.GetProperty("tool_use_id").GetString(), result.GetProperty("is_error").GetBoolean()));
    }

    [Fact]
    public async Task Api_errors_and_bad_transcripts_become_error_events()
    {
        var api = new FakeApi((HttpStatusCode.Unauthorized, """{"type":"error","error":{"type":"authentication_error","message":"invalid x-api-key"}}"""));
        var chat = new Chat("k-1", "model-x", "https://fake.test", api);
        var events = new List<(string Name, string Json)>();
        Task Emit(string name, string json) { events.Add((name, json)); return Task.CompletedTask; }

        using var history = JsonDocument.Parse(Messages(("user", "hello")));
        await chat.RunAsync(null, _dir, history.RootElement, Emit, CancellationToken.None);
        Assert.Equal("error", events.Single().Name);
        Assert.Contains("invalid x-api-key", events[0].Json);

        events.Clear();
        using var empty = JsonDocument.Parse("[]");
        await chat.RunAsync(null, _dir, empty.RootElement, Emit, CancellationToken.None);
        Assert.Equal("error", events.Single().Name);
        Assert.Empty(api.Requests.Skip(1));
    }

    [Fact]
    public async Task Serve_streams_chat_events_and_reports_when_no_key_is_set()
    {
        var file = Path.Combine(_dir, "a.md");
        File.WriteAllText(file, "# Hi\n");
        using var off = new Serve(0, requireToken: false, workspace: _dir);
        off.Start();
        using var client = new HttpClient { BaseAddress = new Uri(off.Url) };
        var body = new StringContent("""{"file":"a.md","messages":[{"role":"user","content":"hi"}]}""", Encoding.UTF8, "application/json");
        var refused = await client.PostAsync("/chat", body);
        Assert.Equal(HttpStatusCode.ServiceUnavailable, refused.StatusCode);
        var files = JsonDocument.Parse(await client.GetStringAsync("/files")).RootElement;
        Assert.False(files.GetProperty("chat").GetBoolean());

        var api = new FakeApi((HttpStatusCode.OK, """{"content":[{"type":"text","text":"Hello!"}],"stop_reason":"end_turn"}"""));
        using var on = new Serve(0, requireToken: false, workspace: _dir, chat: new Chat("k-1", "model-x", "https://fake.test", api));
        on.Start();
        using var client2 = new HttpClient { BaseAddress = new Uri(on.Url) };
        Assert.Equal("model-x", JsonDocument.Parse(await client2.GetStringAsync("/files")).RootElement.GetProperty("model").GetString());
        var response = await client2.PostAsync("/chat", new StringContent("""{"file":"a.md","messages":[{"role":"user","content":"hi"}]}""", Encoding.UTF8, "application/json"));
        Assert.Equal("text/event-stream", response.Content.Headers.ContentType?.MediaType);
        var text = await response.Content.ReadAsStringAsync();
        Assert.Contains("event: text\ndata: {\"text\":\"Hello!\"}", text);
        Assert.Contains("event: done", text);
        var missing = await client2.PostAsync("/chat", new StringContent("""{"file":"nope.md","messages":[]}""", Encoding.UTF8, "application/json"));
        Assert.Equal(HttpStatusCode.NotFound, missing.StatusCode);
    }

    [Fact]
    public async Task Instructions_join_the_system_prompt_and_a_selection_replaces_the_outline()
    {
        File.WriteAllBytes(Path.Combine(_dir, "doc.docx"), Docx(P("Quarterly revenue grew"), P("Second paragraph")));
        var api = new FakeApi((HttpStatusCode.OK, """{"content":[{"type":"text","text":"Done."}],"stop_reason":"end_turn"}"""));
        using var serve = new Serve(0, requireToken: false, workspace: _dir, chat: new Chat("k-1", "model-x", "https://fake.test", api));
        serve.Start();
        using var client = new HttpClient { BaseAddress = new Uri(serve.Url) };
        async Task<JsonElement> Send(object body)
        {
            var response = await client.PostAsync("/chat", new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json"));
            Assert.Contains("event: done", await response.Content.ReadAsStringAsync());
            return JsonDocument.Parse(api.Requests[^1]).RootElement;
        }

        var plain = await Send(new { file = "doc.docx", messages = new[] { new { role = "user", content = "Shorten it" } } });
        Assert.Contains("## Current outline", plain.GetProperty("system").GetString());
        Assert.DoesNotContain("preferences", plain.GetProperty("system").GetString());

        var tuned = await Send(new { file = "doc.docx", messages = new[] { new { role = "user", content = "Shorten it" } }, instructions = "Use formal language.", selection = "Quarterly revenue grew" });
        var system = tuned.GetProperty("system").GetString()!;
        Assert.EndsWith("## The user's preferences\n\nUse formal language.\n", system);
        Assert.DoesNotContain("## Current outline", system);
        Assert.Contains("Open file: doc.docx", system);
        var content = tuned.GetProperty("messages")[0].GetProperty("content").GetString()!;
        Assert.StartsWith("The user selected this part of the document:\n<selection>\nQuarterly revenue grew\n</selection>\n\n", content);
        Assert.EndsWith("Shorten it", content);

        var nothingSelected = await Send(new { file = "doc.docx", messages = new[] { new { role = "user", content = "Hi" } }, selection = "" });
        Assert.DoesNotContain("## Current outline", nothingSelected.GetProperty("system").GetString());
        Assert.Equal("Hi", nothingSelected.GetProperty("messages")[0].GetProperty("content").GetString());
    }

    /// <summary>An OpenAI-compatible server: answers each request with reply(n, request body) and records the requests.</summary>
    internal sealed class FakeOpenAi(Func<int, string, (HttpStatusCode Status, string Body)> reply) : HttpMessageHandler
    {
        public FakeOpenAi(params (HttpStatusCode Status, string Body)[] replies) : this((n, _) => replies[Math.Min(n, replies.Length - 1)]) { }

        public List<(string Url, string? Auth, JsonElement Body)> Requests { get; } = [];

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            var body = await request.Content!.ReadAsStringAsync(ct);
            Requests.Add((request.RequestUri!.ToString(), request.Headers.Authorization?.ToString(), JsonDocument.Parse(body).RootElement.Clone()));
            var (status, text) = reply(Requests.Count - 1, body);
            return new HttpResponseMessage(status) { Content = new StringContent(text, Encoding.UTF8, status == HttpStatusCode.OK ? "text/event-stream" : "application/json") };
        }
    }

    /// <summary>A server-sent event stream of chat.completion.chunk deltas, then [DONE].</summary>
    internal static string Sse(params string[] chunks) => string.Concat(chunks.Select(c => "data: " + c + "\n\n")) + "data: [DONE]\n\n";

    internal static string Delta(string delta, string? finish = null) =>
        "{\"choices\":[{\"index\":0,\"delta\":" + delta + ",\"finish_reason\":" + (finish is null ? "null" : "\"" + finish + "\"") + "}]}";

    static async Task<List<(string Name, JsonElement Data)>> Run(Chat chat, string? file, string dir, params (string Role, string Content)[] turns)
    {
        var events = new List<(string Name, JsonElement Data)>();
        using var history = JsonDocument.Parse(Messages(turns));
        await chat.RunAsync(file, dir, history.RootElement, (name, json) => { events.Add((name, JsonDocument.Parse(json).RootElement.Clone())); return Task.CompletedTask; }, CancellationToken.None);
        return events;
    }

    [Fact]
    public async Task OpenAi_streams_text_deltas_into_one_text_event()
    {
        var api = new FakeOpenAi((HttpStatusCode.OK, Sse(Delta("""{"role":"assistant","content":"Hel"}"""), Delta("""{"content":"lo"}"""), Delta("{}", "stop"))));
        var chat = new Chat("k-1", "gpt-x", "https://fake.test/v1/", api, openAi: true);

        var events = await Run(chat, null, _dir, ("user", "hi"));

        Assert.Equal(["text", "done"], events.Select(e => e.Name));
        Assert.Equal("Hello", events[0].Data.GetProperty("text").GetString());
        var (url, auth, body) = api.Requests.Single();
        Assert.Equal("https://fake.test/v1/chat/completions", url);
        Assert.Equal("Bearer k-1", auth);
        Assert.Equal("gpt-x", body.GetProperty("model").GetString());
        Assert.True(body.GetProperty("stream").GetBoolean());
        Assert.Equal(["system", "user"], body.GetProperty("messages").EnumerateArray().Select(m => m.GetProperty("role").GetString()));
        var tool = body.GetProperty("tools")[0];
        Assert.Equal("function", tool.GetProperty("type").GetString());
        Assert.Equal("writer", tool.GetProperty("function").GetProperty("name").GetString());
        Assert.Equal("command", tool.GetProperty("function").GetProperty("parameters").GetProperty("required")[0].GetString());
    }

    [Fact]
    public async Task OpenAi_tool_call_arguments_arrive_in_fragments_and_the_result_goes_back_as_a_tool_message()
    {
        var file = Path.Combine(_dir, "doc.docx");
        File.WriteAllBytes(file, Docx(P("Old text")));
        var api = new FakeOpenAi(
            (HttpStatusCode.OK, Sse(
                Delta("""{"role":"assistant","content":null,"reasoning_content":"Paragraph one","tool_calls":[{"index":0,"id":"c1","type":"function","function":{"name":"writer","arguments":""}}]}"""),
                Delta("""{"tool_calls":[{"index":0,"function":{"arguments":"{\"command\":\"set doc.docx /body/para"}}]}"""),
                Delta("""{"tool_calls":[{"index":0,"function":{"arguments":"graph[1] --prop text=Hi\"}"}}]}"""),
                Delta("{}", "tool_calls"))),
            (HttpStatusCode.OK, Sse(Delta("""{"content":"Done."}"""), Delta("{}", "stop"))));
        var chat = new Chat("k-1", "m", "https://fake.test", api, openAi: true);

        var events = await Run(chat, file, _dir, ("user", "Change the first paragraph to Hi"));

        Assert.Equal(["tool", "text", "done"], events.Select(e => e.Name));
        Assert.Equal("set doc.docx /body/paragraph[1] --prop text=Hi", events[0].Data.GetProperty("command").GetString());
        Assert.Equal(0, events[0].Data.GetProperty("code").GetInt32());
        Assert.Equal("Done.", events[1].Data.GetProperty("text").GetString());
        using (var doc = OpenDocx(File.ReadAllBytes(file))) Assert.Equal("Hi", doc.Root.Children[0].Children[0].Text);

        var second = api.Requests[1].Body.GetProperty("messages");
        Assert.Equal(["system", "user", "assistant", "tool"], second.EnumerateArray().Select(m => m.GetProperty("role").GetString()));
        var asked = second[2];
        Assert.Equal(JsonValueKind.Null, asked.GetProperty("content").ValueKind);
        Assert.Equal("Paragraph one", asked.GetProperty("reasoning_content").GetString());
        var call = asked.GetProperty("tool_calls")[0];
        Assert.Equal(("c1", "writer"), (call.GetProperty("id").GetString(), call.GetProperty("function").GetProperty("name").GetString()));
        Assert.Equal("""{"command":"set doc.docx /body/paragraph[1] --prop text=Hi"}""", call.GetProperty("function").GetProperty("arguments").GetString());
        Assert.Equal("c1", second[3].GetProperty("tool_call_id").GetString());
        Assert.Contains("/body/paragraph[1]", second[3].GetProperty("content").GetString());
    }

    [Fact]
    public async Task OpenAi_parallel_tool_calls_run_in_order_and_each_gets_its_result()
    {
        File.WriteAllBytes(Path.Combine(_dir, "doc.docx"), Docx(P("Hello")));
        var api = new FakeOpenAi(
            (HttpStatusCode.OK, Sse(
                Delta("""{"content":"Looking.","tool_calls":[{"index":0,"id":"a","type":"function","function":{"name":"writer","arguments":"{\"command\":"}}]}"""),
                Delta("""{"tool_calls":[{"index":1,"id":"b","type":"function","function":{"name":"writer","arguments":"{\"command\":"}}]}"""),
                Delta("""{"tool_calls":[{"index":0,"function":{"arguments":"\"view doc.docx text\"}"}},{"index":1,"function":{"arguments":"\"view doc.docx outline\"}"}}]}"""),
                Delta("{}", "tool_calls"))),
            (HttpStatusCode.OK, Sse(Delta("""{"content":"It says Hello."}"""), Delta("{}", "stop"))));
        var chat = new Chat("", "m", "http://localhost:11434/v1", api, openAi: true);

        var events = await Run(chat, null, _dir, ("user", "What does it say?"));

        Assert.Equal(["text", "tool", "tool", "text", "done"], events.Select(e => e.Name));
        Assert.Equal(["view doc.docx text", "view doc.docx outline"], events.Where(e => e.Name == "tool").Select(e => e.Data.GetProperty("command").GetString()));
        Assert.Null(api.Requests[0].Auth); // no key, no Authorization header
        var second = api.Requests[1].Body.GetProperty("messages");
        Assert.Equal(["a", "b"], second[2].GetProperty("tool_calls").EnumerateArray().Select(c => c.GetProperty("id").GetString()));
        Assert.Equal(("tool", "a"), (second[3].GetProperty("role").GetString(), second[3].GetProperty("tool_call_id").GetString()));
        Assert.Equal(("tool", "b"), (second[4].GetProperty("role").GetString(), second[4].GetProperty("tool_call_id").GetString()));
        Assert.Equal("Hello\n", second[3].GetProperty("content").GetString());
    }

    [Fact]
    public async Task Api_failures_become_short_chinese_errors_without_the_key()
    {
        const string key = "sk-secret-0123456789";
        async Task<JsonElement> Failing(HttpMessageHandler handler, bool openAi = true)
        {
            var events = await Run(new Chat(key, "m", "https://fake.test", handler, openAi), null, _dir, ("user", "hi"));
            Assert.Equal("error", events.Single().Name);
            Assert.DoesNotContain(key, events[0].Data.GetRawText());
            return events[0].Data;
        }

        var unauthorized = await Failing(new FakeOpenAi((HttpStatusCode.Unauthorized, $$$"""{"error":{"message":"Incorrect API key provided: {{{key}}}","type":"invalid_request_error"}}""")));
        Assert.Equal("Key 无效", unauthorized.GetProperty("message").GetString());
        Assert.Contains("Incorrect API key provided: ***", unauthorized.GetProperty("hint").GetString());
        Assert.Equal("地址或模型不对", (await Failing(new FakeOpenAi((HttpStatusCode.NotFound, """{"error":{"message":"The model `m` does not exist"}}""")))).GetProperty("message").GetString());
        Assert.Equal("这个模型不支持工具调用，请换一个支持 function calling 的模型",
            (await Failing(new FakeOpenAi((HttpStatusCode.BadRequest, """{"error":{"message":"registry.ollama.ai/library/gemma:2b does not support tools","type":"api_error"}}""")))).GetProperty("message").GetString());
        Assert.Equal("这个模型不支持工具调用，请换一个支持 function calling 的模型",
            (await Failing(new FakeOpenAi((HttpStatusCode.BadRequest, """{"error":"deepseek-reasoner does not support Function Calling"}""")))).GetProperty("message").GetString());
        Assert.Equal("模型服务出错（429）：Rate limit reached", (await Failing(new FakeOpenAi((HttpStatusCode.TooManyRequests, """{"error":{"message":"Rate limit reached"}}""")))).GetProperty("message").GetString());
        var offline = await Failing(new Unreachable());
        Assert.Equal("连不上", offline.GetProperty("message").GetString());
        Assert.Contains("Connection refused", offline.GetProperty("hint").GetString());
        Assert.Equal("连不上", (await Failing(new Unreachable(), openAi: false)).GetProperty("message").GetString());
        Assert.Equal("接口地址不对：那里返回的不是模型接口的数据", (await Failing(new FakeOpenAi((HttpStatusCode.OK, "<!doctype html><p>Welcome</p>")))).GetProperty("message").GetString());
        Assert.Equal("模型服务出错：Internal error", (await Failing(new FakeOpenAi((HttpStatusCode.OK, Sse(Delta("""{"content":"a"}"""), """{"error":{"message":"Internal error"}}"""))))).GetProperty("message").GetString());
    }

    internal sealed class Unreachable : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct) =>
            throw new HttpRequestException("Connection refused (fake.test:443)");
    }

    [Fact]
    public async Task Earlier_turns_are_folded_into_the_new_message_when_the_provider_wants_reasoning_back()
    {
        const string refusal = """{"error":{"message":"The `reasoning_content` in the thinking mode must be passed back to the API.","type":"invalid_request_error"}}""";
        var api = new FakeOpenAi((n, body) => JsonDocument.Parse(body).RootElement.GetProperty("messages").EnumerateArray()
            .Any(m => m.GetProperty("role").GetString() == "assistant" && !m.TryGetProperty("reasoning_content", out _))
            ? (HttpStatusCode.BadRequest, refusal) : (HttpStatusCode.OK, Sse(Delta("""{"reasoning_content":"…","content":"ok"}"""), Delta("{}", "stop"))));
        var chat = new Chat("k-1", "deepseek-flash", "https://fake.test", api, openAi: true);

        var events = await Run(chat, null, _dir, ("user", "Summarise it"), ("assistant", "It is about Q3."), ("user", "Shorter"));

        Assert.Equal(["text", "done"], events.Select(e => e.Name));
        Assert.Equal(2, api.Requests.Count);
        var retried = api.Requests[1].Body.GetProperty("messages");
        Assert.Equal(["system", "user"], retried.EnumerateArray().Select(m => m.GetProperty("role").GetString()));
        var folded = retried[1].GetProperty("content").GetString()!;
        Assert.StartsWith("Earlier in this conversation:\n<conversation>\nuser: Summarise it\n\nassistant: It is about Q3.\n\n</conversation>\n\n", folded);
        Assert.EndsWith("Shorter", folded);

        await Run(chat, null, _dir, ("user", "a"), ("assistant", "b"), ("user", "c"));
        Assert.Equal(3, api.Requests.Count); // folded up front from then on
    }

    [Fact]
    public async Task OpenAi_reasoning_models_get_reasoning_effort_none_once_they_refuse_tools_with_reasoning()
    {
        // gpt-5.4 and later default to reasoning, and /chat/completions takes function tools only with reasoning_effort "none"
        const string refusal = """{"error":{"message":"Function tools with reasoning_effort are not supported for gpt-6-luna in /v1/chat/completions. Please use /v1/responses instead.","type":"invalid_request_error","param":"reasoning_effort"}}""";
        var api = new FakeOpenAi((_, body) => JsonDocument.Parse(body).RootElement.TryGetProperty("reasoning_effort", out var effort) && effort.GetString() == "none"
            ? (HttpStatusCode.OK, Sse(Delta("""{"content":"ok"}"""), Delta("{}", "stop"))) : (HttpStatusCode.BadRequest, refusal));
        var chat = new Chat("k-1", "gpt-6-luna", "https://fake.test/v1", api, openAi: true);

        var events = await Run(chat, null, _dir, ("user", "hi"));

        Assert.Equal(["text", "done"], events.Select(e => e.Name));
        Assert.Equal(2, api.Requests.Count);
        Assert.False(api.Requests[0].Body.TryGetProperty("reasoning_effort", out _), "older models refuse the parameter, so it is not sent up front");
        Assert.Equal("none", api.Requests[1].Body.GetProperty("reasoning_effort").GetString());
        Assert.Null(await chat.TestAsync(CancellationToken.None));
        Assert.Equal(3, api.Requests.Count); // from then on it goes with the first request
    }

    public void Dispose() => Directory.Delete(_dir, true);
}
