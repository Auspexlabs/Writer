using System.Net;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Web;
using Writer.Cli;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>用 ChatGPT 登录: the Codex CLI's sign-in (its authorize page, the browser's return to the loopback port, the code swapped
/// for tokens), the tokens kept in the settings file and refreshed, and the assistant, the autocomplete and the model list on the
/// Codex backend's Responses API. OpenAI's sign-in service and backend are stand-ins answering by path.</summary>
public class ChatGptTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("writer-chatgpt").FullName;
    readonly List<Serve> _servers = [];
    string SettingsFile => Path.Combine(_dir, "config", "ai.json");

    internal sealed record Request(string Method, Uri Url, Dictionary<string, string> Headers, string Body, string? Type)
    {
        public string Path => Url.AbsolutePath;
        public JsonElement Json => JsonDocument.Parse(Body).RootElement.Clone();
    }

    /// <summary>OpenAI's side: answers each request with reply(request) → (status, content type, body) and records them in order.</summary>
    sealed class OpenAi(Func<Request, (HttpStatusCode Status, string Type, string Body)> reply) : HttpMessageHandler
    {
        public List<Request> Requests { get; } = [];

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            var r = new Request(request.Method.Method, request.RequestUri!,
                request.Headers.NonValidated.ToDictionary(h => h.Key.ToLowerInvariant(), h => h.Value.ToString()),
                request.Content is null ? "" : await request.Content.ReadAsStringAsync(ct), request.Content?.Headers.ContentType?.MediaType);
            lock (Requests) Requests.Add(r);
            var (status, type, body) = reply(r);
            return new HttpResponseMessage(status) { Content = new StringContent(body, Encoding.UTF8, type) };
        }

        public Request[] To(string path)
        {
            lock (Requests) return Requests.Where(r => r.Path == path).ToArray();
        }
    }

    const string Responses = "/backend-api/codex/responses";
    const string Token = "/oauth/token";

    (HttpClient Client, List<string> Opened) Start(OpenAi api)
    {
        var opened = new List<string>();
        var store = new AiStore(SettingsFile, _ => null, api) { LoginPort = 0, Browser = url => { opened.Add(url); return false; } };
        var server = new Serve(0, requireToken: true, workspace: _dir, ai: store);
        server.Start();
        _servers.Add(server);
        var client = new HttpClient { BaseAddress = new Uri(server.Url) };
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", server.Token);
        return (client, opened);
    }

    /// <summary>A settings file already signed in: ChatGPT with these tokens (lapsing in an hour unless <paramref name="expires"/> says otherwise).</summary>
    void SignedIn(string access = "access-1", string refresh = "refresh-1", long? expires = null) =>
        new AiConfig("chatgpt", ChatGpt.BaseUrl, "gpt-6-astra", "", "file")
        {
            CompleteModel = "gpt-5.6-luna",
            Account = new ChatGptTokens(access, refresh, expires ?? DateTimeOffset.UtcNow.AddHours(1).ToUnixTimeMilliseconds(), "acct-42", "ada@example.com", "plus"),
        }.Write(SettingsFile);

    static StringContent Body(object o) => new(JsonSerializer.Serialize(o), Encoding.UTF8, "application/json");

    /// <summary>A JSON answer that must carry no token.</summary>
    static async Task<JsonElement> Read(HttpResponseMessage response)
    {
        var text = await response.Content.ReadAsStringAsync();
        foreach (var secret in new[] { "access-", "refresh-", "sealed-" }) Assert.DoesNotContain(secret, text);
        return JsonDocument.Parse(text).RootElement.Clone();
    }

    static string Jwt(object claims) => "eyJhbGciOiJub25lIn0." + ChatGpt.Base64Url(JsonSerializer.SerializeToUtf8Bytes(claims)) + ".sig";

    static string TokenAnswer(string access, string refresh, string plan = "plus") => JsonSerializer.Serialize(new Dictionary<string, object>
    {
        ["access_token"] = access,
        ["refresh_token"] = refresh,
        ["expires_in"] = 3600,
        ["id_token"] = Jwt(new Dictionary<string, object>
        {
            ["email"] = "ada@example.com",
            ["https://api.openai.com/auth"] = new Dictionary<string, object> { ["chatgpt_account_id"] = "acct-42", ["chatgpt_plan_type"] = plan },
        }),
    });

    /// <summary>The Responses API's event stream: each event's type line, then its data.</summary>
    static (HttpStatusCode, string, string) Stream(params string[] events) =>
        (HttpStatusCode.OK, "text/event-stream", string.Concat(events.Select(e => $"event: {JsonDocument.Parse(e).RootElement.GetProperty("type").GetString()}\ndata: {e}\n\n")));

    static string TextDelta(string text) => JsonSerializer.Serialize(new { type = "response.output_text.delta", delta = text });
    static string Done(object item) => JsonSerializer.Serialize(new { type = "response.output_item.done", item });
    static string Said(string text) => Done(new { type = "message", id = "msg_1", role = "assistant", status = "completed", content = new[] { new { type = "output_text", text } } });
    const string Completed = """{"type":"response.completed","response":{"id":"resp_1","status":"completed"}}""";

    static (HttpStatusCode, string, string) Json(string body, HttpStatusCode status = HttpStatusCode.OK) => (status, "application/json", body);

    static async Task<List<(string Name, JsonElement Data)>> Events(HttpResponseMessage response)
    {
        var text = await response.Content.ReadAsStringAsync();
        foreach (var secret in new[] { "access-", "refresh-", "sealed-" }) Assert.DoesNotContain(secret, text);
        return text.Split("\n\n", StringSplitOptions.RemoveEmptyEntries).Select(block =>
        {
            var lines = block.Split('\n');
            return (lines[0]["event: ".Length..], JsonDocument.Parse(lines[1]["data: ".Length..]).RootElement.Clone());
        }).ToList();
    }

    [Fact]
    public async Task Signing_in_opens_the_Codex_CLIs_authorize_page_and_the_browsers_return_saves_the_tokens_and_switches_to_ChatGPT()
    {
        var api = new OpenAi(r => r.Path switch
        {
            Token => Json(TokenAnswer("access-1", "refresh-1")),
            "/backend-api/codex/models" => Json("""{"models":[{"slug":"gpt-5.5","display_name":"GPT-5.5"},{"slug":"gpt-6-astra","display_name":"GPT-6 Astra"},{"slug":"gpt-5.6-luna"},{"slug":"codex-auto-review","visibility":"hide"}]}"""),
            _ => Json("{}", HttpStatusCode.NotFound),
        });
        var (client, opened) = Start(api);
        Assert.Equal("none", (await Read(await client.GetAsync("/ai/chatgpt/login"))).GetProperty("state").GetString());

        var start = await Read(await client.PostAsync("/ai/chatgpt/login", Body(new { })));
        var url = new Uri(start.GetProperty("url").GetString()!);
        Assert.False(start.GetProperty("opened").GetBoolean(), "no browser here: the page shows the link");
        Assert.Equal([start.GetProperty("url").GetString()!], opened);
        Assert.Equal("https://auth.openai.com/oauth/authorize", url.GetLeftPart(UriPartial.Path));
        var q = HttpUtility.ParseQueryString(url.Query);
        Assert.Equal(["response_type", "client_id", "redirect_uri", "scope", "code_challenge", "code_challenge_method", "id_token_add_organizations", "codex_cli_simplified_flow", "state", "originator"], q.AllKeys.Select(k => k!));
        Assert.Equal(("code", "app_EMoamEEZ73f0CkXaXp7hrann", "S256", "true", "true", "codex_cli_rs"),
            (q["response_type"], q["client_id"], q["code_challenge_method"], q["id_token_add_organizations"], q["codex_cli_simplified_flow"], q["originator"]));
        Assert.Equal("openid profile email offline_access api.connectors.read api.connectors.invoke", q["scope"]);
        var callback = new Uri(q["redirect_uri"]!);
        Assert.Equal(("http", "localhost", "/auth/callback"), (callback.Scheme, callback.Host, callback.AbsolutePath));
        Assert.Equal("waiting", (await Read(await client.GetAsync("/ai/chatgpt/login"))).GetProperty("state").GetString());

        using var browser = new HttpClient();
        Assert.Contains("没有登录成功", await browser.GetStringAsync(callback + "?code=stolen&state=someone-elses"));
        Assert.Equal("waiting", (await Read(await client.GetAsync("/ai/chatgpt/login"))).GetProperty("state").GetString());
        Assert.Empty(api.To(Token));
        var back = callback + "?code=the-code&state=" + Uri.EscapeDataString(q["state"]!);
        Assert.Contains("已登录 ChatGPT", await browser.GetStringAsync(back));
        Assert.Contains("已登录 ChatGPT", await browser.GetStringAsync(back)); // the page again (a reload): not a second sign-in

        var exchange = api.To(Token).Single();
        Assert.Equal(("POST", "application/x-www-form-urlencoded"), (exchange.Method, exchange.Type));
        var form = HttpUtility.ParseQueryString(exchange.Body);
        Assert.Equal(("authorization_code", "the-code", q["redirect_uri"], "app_EMoamEEZ73f0CkXaXp7hrann"), (form["grant_type"], form["code"], form["redirect_uri"], form["client_id"]));
        Assert.Equal(q["code_challenge"], ChatGpt.Base64Url(SHA256.HashData(Encoding.ASCII.GetBytes(form["code_verifier"]!))));

        var done = await Read(await client.GetAsync("/ai/chatgpt/login"));
        Assert.Equal(("done", "ada@example.com", "plus"), (done.GetProperty("state").GetString(), done.GetProperty("email").GetString(), done.GetProperty("plan").GetString()));
        var ai = await Read(await client.GetAsync("/ai"));
        Assert.Equal(("chatgpt", "https://chatgpt.com/backend-api/codex", "gpt-6-astra", "gpt-5.6-luna", false),
            (ai.GetProperty("provider").GetString(), ai.GetProperty("baseUrl").GetString(), ai.GetProperty("model").GetString(), ai.GetProperty("completeModel").GetString(), ai.GetProperty("hasKey").GetBoolean()));
        Assert.Equal(("ada@example.com", "plus"), (ai.GetProperty("account").GetProperty("email").GetString(), ai.GetProperty("account").GetProperty("plan").GetString()));
        Assert.True((await Read(await client.GetAsync("/files"))).GetProperty("chat").GetBoolean());
        var listed = api.Requests.Single(r => r.Path == "/backend-api/codex/models");
        Assert.Equal(("Bearer access-1", "acct-42", "client_version=0.153.4"), (listed.Headers["authorization"], listed.Headers["chatgpt-account-id"], listed.Url.Query.TrimStart('?')));

        var saved = File.ReadAllText(SettingsFile);
        Assert.Contains("\"refresh\": \"refresh-1\"", saved);
        if (!OperatingSystem.IsWindows()) Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(SettingsFile));
    }

    [Fact]
    public async Task A_sign_in_the_user_refuses_or_cancels_ends_with_the_reason_and_changes_no_settings()
    {
        var api = new OpenAi(r => r.Path == Token ? Json("""{"error":"invalid_grant","error_description":"Code expired"}""", HttpStatusCode.BadRequest) : Json("{}", HttpStatusCode.NotFound));
        var (client, _) = Start(api);
        using var browser = new HttpClient();

        var q = HttpUtility.ParseQueryString(new Uri((await Read(await client.PostAsync("/ai/chatgpt/login", Body(new { })))).GetProperty("url").GetString()!).Query);
        Assert.Contains("没有登录成功", await browser.GetStringAsync(q["redirect_uri"] + "?error=access_denied&error_description=User+cancelled&state=" + Uri.EscapeDataString(q["state"]!)));
        var refused = await Read(await client.GetAsync("/ai/chatgpt/login"));
        Assert.Equal(("error", "ChatGPT 登录被取消或拒绝：User cancelled"), (refused.GetProperty("state").GetString(), refused.GetProperty("error").GetString()));

        q = HttpUtility.ParseQueryString(new Uri((await Read(await client.PostAsync("/ai/chatgpt/login", Body(new { })))).GetProperty("url").GetString()!).Query);
        await browser.GetStringAsync(q["redirect_uri"] + "?code=old&state=" + Uri.EscapeDataString(q["state"]!));
        var expired = await Read(await client.GetAsync("/ai/chatgpt/login"));
        Assert.Equal("ChatGPT 登录没有完成：Code expired", expired.GetProperty("error").GetString());

        await client.PostAsync("/ai/chatgpt/login", Body(new { }));
        Assert.Equal("cancelled", (await Read(await client.PostAsync("/ai/chatgpt/cancel", Body(new { })))).GetProperty("state").GetString());
        Assert.False(File.Exists(SettingsFile), "nothing was saved");
        Assert.Equal(HttpStatusCode.UnsupportedMediaType, (await client.PostAsync("/ai/chatgpt/login", new StringContent(""))).StatusCode);
    }

    [Fact]
    public async Task The_assistant_talks_to_the_Codex_backend_and_gives_back_the_reply_items_with_its_tool_results()
    {
        var file = Path.Combine(_dir, "doc.docx");
        File.WriteAllBytes(file, Docx(P("Old text")));
        SignedIn();
        var turn = 0;
        var api = new OpenAi(r => r.Path != Responses ? Json("{}", HttpStatusCode.NotFound) : turn++ == 0
            ? Stream(
                Done(new { type = "reasoning", id = "rs_1", summary = Array.Empty<object>(), encrypted_content = "sealed-1" }),
                Done(new { type = "function_call", id = "fc_1", call_id = "call_1", name = "writer", arguments = """{"command":"set doc.docx /body/paragraph[1] --prop text=Hi"}""", status = "completed" }),
                Completed)
            : Stream(TextDelta("Do"), TextDelta("ne."), Said("Done."), Completed));
        var (client, _) = Start(api);
        var system = Chat.SystemPrompt(file, _dir, outline: true, web: true); // the document as the turn starts

        var events = await Events(await client.PostAsync("/chat", Body(new { file = "doc.docx", messages = new[] { new { role = "user", content = "Change the first paragraph to Hi" } } })));

        Assert.Equal(["tool", "delta", "delta", "text", "done"], events.Select(e => e.Name));
        Assert.Equal(("set doc.docx /body/paragraph[1] --prop text=Hi", true), (events[0].Data.GetProperty("command").GetString(), events[0].Data.GetProperty("wrote").GetBoolean()));
        Assert.Equal("Done.", events[3].Data.GetProperty("text").GetString());
        using (var doc = OpenDocx(File.ReadAllBytes(file))) Assert.Equal("Hi", doc.Root.Children[0].Children[0].Text);

        var (first, second) = (api.To(Responses)[0], api.To(Responses)[1]);
        Assert.Equal("https://chatgpt.com/backend-api/codex/responses", first.Url.ToString());
        Assert.Equal(("Bearer access-1", "acct-42", "codex_cli_rs", "0.153.4"), (first.Headers["authorization"], first.Headers["chatgpt-account-id"], first.Headers["originator"], first.Headers["version"]));
        Assert.StartsWith("codex_cli_rs/0.153.4 (", first.Headers["user-agent"]);
        Assert.True(Guid.TryParse(first.Headers["session_id"], out _));
        Assert.Equal(first.Headers["session_id"], second.Headers["session_id"]);
        var body = first.Json;
        Assert.Equal(("gpt-6-astra", false, true, "auto"), (body.GetProperty("model").GetString(), body.GetProperty("store").GetBoolean(), body.GetProperty("stream").GetBoolean(), body.GetProperty("tool_choice").GetString()));
        Assert.Equal(first.Headers["session_id"], body.GetProperty("prompt_cache_key").GetString());
        // a document's prompt carries its format's reference (about 49 KB for Word), past what instructions may hold: it goes first, as
        // the developer's message
        Assert.True(Encoding.UTF8.GetByteCount(system) > 30000);
        Assert.Equal("You are the assistant inside Writer, a document editor. Follow the developer message.", body.GetProperty("instructions").GetString());
        Assert.Equal(["reasoning.encrypted_content"], body.GetProperty("include").EnumerateArray().Select(x => x.GetString()));
        var tool = body.GetProperty("tools")[0];
        Assert.Equal(("function", "writer", "command"), (tool.GetProperty("type").GetString(), tool.GetProperty("name").GetString(), tool.GetProperty("parameters").GetProperty("required")[0].GetString()));
        var asked = body.GetProperty("input");
        Assert.Equal(2, asked.GetArrayLength());
        Assert.Equal(("developer", system), (asked[0].GetProperty("role").GetString(), asked[0].GetProperty("content")[0].GetProperty("text").GetString()));
        Assert.Equal("""{"type":"message","role":"user","content":[{"type":"input_text","text":"Change the first paragraph to Hi"}]}""", asked[1].GetRawText());

        var input = second.Json.GetProperty("input");
        Assert.Equal(["message", "message", "reasoning", "function_call", "function_call_output"], input.EnumerateArray().Select(i => i.GetProperty("type").GetString()));
        Assert.Equal("""{"type":"reasoning","summary":[],"encrypted_content":"sealed-1"}""", input[2].GetRawText());
        Assert.Equal(["type", "call_id", "name", "arguments"], input[3].EnumerateObject().Select(p => p.Name)); // no id: nothing is stored to point at
        Assert.Equal(("call_1", "writer", """{"command":"set doc.docx /body/paragraph[1] --prop text=Hi"}"""),
            (input[3].GetProperty("call_id").GetString(), input[3].GetProperty("name").GetString(), input[3].GetProperty("arguments").GetString()));
        Assert.Equal("call_1", input[4].GetProperty("call_id").GetString());
        Assert.Contains("/body/paragraph[1]", input[4].GetProperty("output").GetString());
    }

    [Fact]
    public async Task Tokens_about_to_lapse_are_refreshed_and_saved_before_the_request_and_a_refused_one_gets_one_refresh_and_retry()
    {
        SignedIn(expires: DateTimeOffset.UtcNow.AddMinutes(2).ToUnixTimeMilliseconds());
        var refreshed = 0;
        var refuse = false;
        var api = new OpenAi(r => r.Path switch
        {
            Token => ++refreshed <= 2 ? Json(TokenAnswer("access-" + (refreshed + 1), "refresh-" + (refreshed + 1))) : Json("""{"error":{"code":"refresh_token_reused","message":"Your refresh token has already been used"}}""", HttpStatusCode.Unauthorized),
            Responses when refuse && r.Headers["authorization"] == "Bearer access-2" => Json("""{"error":{"message":"token expired"}}""", HttpStatusCode.Unauthorized),
            Responses => Stream(TextDelta("the plan"), TextDelta("\nand more"), TextDelta("\n"), TextDelta("never read"), Completed),
            _ => Json("{}", HttpStatusCode.NotFound),
        });
        var (client, _) = Start(api);

        var text = await Read(await client.PostAsync("/complete", Body(new { before = "Next week we will present ", hint = "the body of a Word document" })));
        Assert.Equal("the plan", text.GetProperty("text").GetString());
        var refresh = api.To(Token).Single();
        Assert.Equal("application/json", refresh.Type);
        Assert.Equal("""{"client_id":"app_EMoamEEZ73f0CkXaXp7hrann","grant_type":"refresh_token","refresh_token":"refresh-1","scope":"openid profile email"}""", refresh.Body);
        Assert.Equal("Bearer access-2", api.To(Responses).Single().Headers["authorization"]);
        Assert.Contains("\"refresh\": \"refresh-2\"", File.ReadAllText(SettingsFile));
        var asked = api.To(Responses).Single().Json;
        Assert.Equal(("gpt-5.6-luna", "none", Assistant.CompleteSystem), (asked.GetProperty("model").GetString(), asked.GetProperty("reasoning").GetProperty("effort").GetString(), asked.GetProperty("instructions").GetString()));
        Assert.False(asked.TryGetProperty("tools", out _), "a completion runs no tools");
        Assert.StartsWith("<where>the body of a Word document</where>", asked.GetProperty("input")[0].GetProperty("content")[0].GetProperty("text").GetString());

        refuse = true;
        Assert.Equal("the plan", (await Read(await client.PostAsync("/complete", Body(new { before = "x" })))).GetProperty("text").GetString());
        Assert.Equal(["Bearer access-2", "Bearer access-3"], api.To(Responses).Skip(1).Select(r => r.Headers["authorization"]));
        Assert.Equal("refresh-2", JsonDocument.Parse(api.To(Token)[1].Body).RootElement.GetProperty("refresh_token").GetString());

        SignedIn("access-9", "refresh-9", expires: 0);
        var lapsed = await client.PostAsync("/complete", Body(new { before = "x" }));
        Assert.Equal(HttpStatusCode.BadRequest, lapsed.StatusCode);
        var error = (await Read(lapsed)).GetProperty("error");
        Assert.Equal("ChatGPT 登录已失效，请重新登录", error.GetProperty("message").GetString());
    }

    [Fact]
    public async Task A_model_that_must_reason_is_asked_again_with_a_little_and_the_plans_models_come_from_the_backend()
    {
        SignedIn();
        var api = new OpenAi(r => r.Path switch
        {
            Responses when r.Json.GetProperty("reasoning").GetProperty("effort").GetString() == "none" =>
                Json("""{"error":{"message":"Unsupported value: 'none' is not supported with this model.","param":"reasoning.effort"}}""", HttpStatusCode.BadRequest),
            Responses => Stream(TextDelta("quarterly report"), Said("quarterly report"), Completed),
            "/backend-api/codex/models" => Json("""{"models":[{"slug":"gpt-6-astra","display_name":"GPT-6 Astra","visibility":"list"},{"slug":"gpt-5.6-luna","display_name":"GPT-5.6 Luna"},{"slug":"internal-thing","visibility":"hide"}]}"""),
            _ => Json("{}", HttpStatusCode.NotFound),
        });
        var (client, _) = Start(api);

        Assert.Equal("quarterly report", (await Read(await client.PostAsync("/complete", Body(new { before = "Attached is the " })))).GetProperty("text").GetString());
        Assert.Equal(["none", "low"], api.To(Responses).Select(r => r.Json.GetProperty("reasoning").GetProperty("effort").GetString()));

        var models = await Read(await client.PostAsync("/ai/models", Body(new { provider = "chatgpt" })));
        Assert.Equal([("gpt-6-astra", "GPT-6 Astra"), ("gpt-5.6-luna", "GPT-5.6 Luna")],
            models.GetProperty("models").EnumerateArray().Select(m => (m.GetProperty("id").GetString(), m.GetProperty("name").GetString())));
        Assert.Equal("Bearer access-1", api.Requests[^1].Headers["authorization"]);
    }

    [Fact]
    public async Task A_used_up_plan_and_Cloudflares_challenge_page_are_said_plainly()
    {
        SignedIn();
        var answers = new Queue<(HttpStatusCode, string, string)>([
            Json("""{"error":{"type":"usage_limit_reached","message":"The usage limit has been reached","plan_type":"plus","resets_in_seconds":540}}""", (HttpStatusCode)429),
            (HttpStatusCode.Forbidden, "text/html", "<!DOCTYPE html><html><title>Just a moment...</title></html>"),
        ]);
        var api = new OpenAi(_ => answers.Dequeue());
        var (client, _) = Start(api);

        async Task<JsonElement> Failed()
        {
            var events = await Events(await client.PostAsync("/chat", Body(new { messages = new[] { new { role = "user", content = "hi" } } })));
            return Assert.Single(events, e => e.Name == "error").Data;
        }
        var used = await Failed();
        Assert.Equal(Chat.SystemPrompt(null, _dir, outline: true, web: true), api.Requests[0].Json.GetProperty("instructions").GetString()); // no document: a short prompt, as instructions
        Assert.Equal(("ChatGPT 套餐的用量已达上限", "大约 9 分钟后恢复。"), (used.GetProperty("message").GetString(), used.GetProperty("hint").GetString()));
        Assert.Equal("ChatGPT 的服务拦下了这次请求（403）", (await Failed()).GetProperty("message").GetString());
    }

    [Fact]
    public async Task The_sign_in_stays_through_a_switch_to_another_service_its_tokens_go_nowhere_else_and_signing_out_deletes_them()
    {
        SignedIn();
        var api = new OpenAi(r => Json("{}", HttpStatusCode.NotFound));
        var (client, _) = Start(api);

        var deepseek = await Read(await client.PutAsync("/ai", Body(new { provider = "deepseek", baseUrl = "https://api.deepseek.com", model = "deepseek-flash", apiKey = "sk-deepseek-0123456789" })));
        Assert.Equal("ada@example.com", deepseek.GetProperty("account").GetProperty("email").GetString());
        var back = await Read(await client.PutAsync("/ai", Body(new { provider = "chatgpt", baseUrl = "https://elsewhere.example/v1", model = "gpt-5.5", apiKey = "sk-typed-0123456789" })));
        Assert.Equal(("https://chatgpt.com/backend-api/codex", "gpt-5.5", false), (back.GetProperty("baseUrl").GetString(), back.GetProperty("model").GetString(), back.GetProperty("hasKey").GetBoolean()));
        Assert.True((await Read(await client.GetAsync("/files"))).GetProperty("chat").GetBoolean(), "no new sign-in needed");
        Assert.DoesNotContain("sk-typed", File.ReadAllText(SettingsFile));

        File.WriteAllText(SettingsFile, File.ReadAllText(SettingsFile).Replace("https://chatgpt.com/backend-api/codex", "https://elsewhere.example"));
        Assert.Equal("https://chatgpt.com/backend-api/codex", AiConfig.Read(SettingsFile)!.BaseUrl);

        var signedOut = await Read(await client.PostAsync("/ai/chatgpt/logout", Body(new { })));
        Assert.False(signedOut.TryGetProperty("account", out _));
        Assert.DoesNotContain("refresh-1", File.ReadAllText(SettingsFile));
        Assert.Equal(HttpStatusCode.ServiceUnavailable, (await client.PostAsync("/complete", Body(new { before = "x" }))).StatusCode);
        var test = await Read(await client.PostAsync("/ai/test", Body(new { provider = "chatgpt", model = "gpt-5.5" })));
        Assert.Equal("请先用 ChatGPT 登录", test.GetProperty("error").GetString());
        Assert.Empty(api.Requests);
    }

    [Fact]
    public void Claims_come_from_the_id_token_and_the_tokens_never_print()
    {
        var jwt = Jwt(new Dictionary<string, object>
        {
            ["https://api.openai.com/auth"] = new Dictionary<string, object> { ["chatgpt_account_id"] = "acct-7", ["chatgpt_plan_type"] = "pro" },
            ["https://api.openai.com/profile"] = new Dictionary<string, object> { ["email"] = "grace@example.com" },
        });
        Assert.Equal(("acct-7", "grace@example.com", "pro"), ChatGpt.Claims(jwt));
        Assert.Equal(("", "", ""), ChatGpt.Claims("not a token"));
        Assert.Equal(("", "", ""), ChatGpt.Claims("a.%%%.b"));
        Assert.Equal(("", "", ""), ChatGpt.Claims(null));

        var (verifier, challenge) = ChatGpt.Pkce();
        Assert.Equal(43, verifier.Length);
        Assert.Equal(challenge, ChatGpt.Base64Url(SHA256.HashData(Encoding.ASCII.GetBytes(verifier))));
        Assert.DoesNotMatch("[+/=]", verifier + challenge);

        var tokens = new ChatGptTokens("access-secret", "refresh-secret", 0, "acct-7", "grace@example.com", "pro");
        Assert.Equal("ChatGPT grace@example.com (pro)", tokens.ToString());
        Assert.True(tokens.Stale);
        Assert.Equal(tokens, ChatGptTokens.FromJson(tokens.ToJson()));
        Assert.Null(ChatGptTokens.FromJson(new System.Text.Json.Nodes.JsonObject()));
        Assert.DoesNotContain("secret", new AiConfig("chatgpt", ChatGpt.BaseUrl, "gpt-5.5", "", "file") { Account = tokens }.ToString());
    }

    public void Dispose()
    {
        foreach (var server in _servers) server.Dispose();
        Directory.Delete(_dir, true);
    }
}
