using System.Net;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Writer.Core;

namespace Writer.Cli;

/// <summary>用 ChatGPT 登录: the sign-in of OpenAI's Codex CLI, as OpenClaw and other tools use it to spend a ChatGPT
/// subscription (Plus, Pro, Business) instead of an API key. Not an OpenAI offer for other apps: the values below are the Codex
/// CLI's (openai/codex, codex-rs/login, checked 2026-09-27: its client id, the authorize request's ten parameters, the loopback
/// callback on port 1455, the token refresh) and OpenAI may change them. The WRITER_CHATGPT_* variables move the endpoints
/// (the tests' stand-ins), the port and the client version.</summary>
public static class ChatGpt
{
    public const string ClientId = "app_EMoamEEZ73f0CkXaXp7hrann";
    public const string DefaultIssuer = "https://auth.openai.com";
    public const int DefaultPort = 1455;
    public const string CallbackPath = "/auth/callback";
    public const string Scope = "openid profile email offline_access api.connectors.read api.connectors.invoke";
    public const string Originator = "codex_cli_rs";
    public const string DefaultClientVersion = "0.153.4";
    /// <summary>The Codex CLI's User-Agent shape: codex_cli_rs/version (system; architecture) and the program it runs in.</summary>
    public static string UserAgent =>
        $"{Originator}/{ClientVersion} ({(OperatingSystem.IsMacOS() ? "Mac OS" : OperatingSystem.IsWindows() ? "Windows" : "Linux")} {Environment.OSVersion.Version}; "
        + $"{(System.Runtime.InteropServices.RuntimeInformation.OSArchitecture == System.Runtime.InteropServices.Architecture.Arm64 ? "arm64" : "x86_64")}) Writer";
    /// <summary>The models a new sign-in starts with when the plan's list does not say otherwise: the Codex CLI's default (OpenClaw's
    /// docs, checked 2026-09-27) for the assistant, the fast one for the autocomplete.</summary>
    public const string DefaultModel = "gpt-6-astra";
    public const string DefaultCompleteModel = "gpt-5.6-luna";

    static string? Env(string name) => Environment.GetEnvironmentVariable(name) is { Length: > 0 } v ? v.TrimEnd('/') : null;
    public static string Issuer => Env("WRITER_CHATGPT_ISSUER") ?? DefaultIssuer;
    /// <summary>The Codex backend (AiConfig.Providers["chatgpt"]).</summary>
    public static string BaseUrl => Env("WRITER_CHATGPT_BASE_URL") ?? AiConfig.Providers["chatgpt"];
    public static int Port => int.TryParse(Env("WRITER_CHATGPT_PORT"), out var p) ? p : DefaultPort;
    public static string ClientVersion => Env("WRITER_CHATGPT_CLIENT_VERSION") ?? DefaultClientVersion;
    /// <summary>Where the browser comes back to; the sign-in service knows the Codex CLI's only, on port 1455.</summary>
    public static string RedirectUri(int port) => $"http://localhost:{port}{CallbackPath}";

    /// <summary>A PKCE pair: the verifier kept here, its S256 challenge sent with the authorize request.</summary>
    public static (string Verifier, string Challenge) Pkce()
    {
        var verifier = Base64Url(RandomNumberGenerator.GetBytes(32));
        return (verifier, Base64Url(SHA256.HashData(Encoding.ASCII.GetBytes(verifier))));
    }

    public static string Base64Url(byte[] bytes) => Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    /// <summary>The page the browser opens: the Codex CLI's authorize request, parameter for parameter.</summary>
    public static string AuthorizeUrl(string challenge, string state, int port) =>
        Issuer + "/oauth/authorize?" + string.Join("&", new (string, string)[]
        {
            ("response_type", "code"), ("client_id", ClientId), ("redirect_uri", RedirectUri(port)), ("scope", Scope),
            ("code_challenge", challenge), ("code_challenge_method", "S256"), ("id_token_add_organizations", "true"),
            ("codex_cli_simplified_flow", "true"), ("state", state), ("originator", Originator),
        }.Select(p => p.Item1 + "=" + Uri.EscapeDataString(p.Item2)));

    /// <summary>What a token says about the account: its ChatGPT account id (the requests carry it), email and plan.</summary>
    public static (string AccountId, string Email, string Plan) Claims(string? jwt)
    {
        try
        {
            var parts = (jwt ?? "").Split('.');
            if (parts.Length < 2) return ("", "", "");
            var payload = parts[1].Replace('-', '+').Replace('_', '/');
            payload = payload.PadRight(payload.Length + (4 - payload.Length % 4) % 4, '=');
            if (JsonNode.Parse(Convert.FromBase64String(payload)) is not JsonObject o) return ("", "", "");
            var auth = o["https://api.openai.com/auth"] as JsonObject;
            var profile = o["https://api.openai.com/profile"] as JsonObject;
            string S(JsonNode? n) => n is JsonValue v && v.TryGetValue<string>(out var s) ? s : "";
            return (S(auth?["chatgpt_account_id"]), S(o["email"]) is { Length: > 0 } e ? e : S(profile?["email"]), S(auth?["chatgpt_plan_type"]));
        }
        catch (Exception e) when (e is FormatException or JsonException)
        {
            return ("", "", "");
        }
    }

    /// <summary>The code from the callback, exchanged for tokens (the authorization-code grant, form-encoded as the Codex CLI sends it).</summary>
    public static Task<ChatGptTokens> Exchange(HttpClient http, string code, string verifier, int port, CancellationToken ct) =>
        Tokens(http, new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["grant_type"] = "authorization_code", ["code"] = code, ["redirect_uri"] = RedirectUri(port), ["client_id"] = ClientId, ["code_verifier"] = verifier,
        }), null, ct);

    /// <summary>New tokens for a refresh token (JSON, as the Codex CLI sends it). A refresh token is used once: the answer's replaces it.</summary>
    public static Task<ChatGptTokens> Refresh(HttpClient http, ChatGptTokens old, CancellationToken ct) =>
        Tokens(http, new StringContent(new JsonObject
        {
            ["client_id"] = ClientId, ["grant_type"] = "refresh_token", ["refresh_token"] = old.Refresh, ["scope"] = "openid profile email",
        }.ToJsonString(), Encoding.UTF8, "application/json"), old, ct);

    static async Task<ChatGptTokens> Tokens(HttpClient http, HttpContent content, ChatGptTokens? old, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, Issuer + "/oauth/token") { Content = content };
        HttpResponseMessage response;
        try
        {
            response = await http.SendAsync(request, ct);
        }
        catch (Exception ex) when (ex is HttpRequestException or IOException)
        {
            throw new WriterException(ErrorCode.Io, "连不上 ChatGPT 的登录服务", ex.Message);
        }
        using (response)
        {
            var text = await response.Content.ReadAsStringAsync(ct);
            JsonObject? o = null;
            try { o = JsonNode.Parse(text) as JsonObject; } catch (JsonException) { }
            if (!response.IsSuccessStatusCode || o is null)
            {
                var why = o?["error_description"]?.ToString() ?? o?["error"]?.ToString() ?? text;
                throw new WriterException(ErrorCode.Io, old is null ? "ChatGPT 登录没有完成" : "ChatGPT 登录已失效，请重新登录", why.Length > 200 ? why[..200] : why);
            }
            string S(string k) => o[k] is JsonValue v && v.TryGetValue<string>(out var s) ? s : "";
            var access = S("access_token");
            if (access.Length == 0) throw new WriterException(ErrorCode.Io, "ChatGPT 登录没有完成", "The token answer carried no access_token.");
            var idToken = S("id_token");
            var (account, email, plan) = Claims(idToken.Length > 0 ? idToken : access);
            if (account.Length == 0) (account, _, _) = Claims(access);
            var expiresIn = o["expires_in"] is JsonValue ev && ev.TryGetValue<long>(out var sec) ? sec : 3600;
            return new ChatGptTokens(access, S("refresh_token") is { Length: > 0 } r ? r : old?.Refresh ?? "",
                DateTimeOffset.UtcNow.AddSeconds(expiresIn).ToUnixTimeMilliseconds(),
                account.Length > 0 ? account : old?.AccountId ?? "", email.Length > 0 ? email : old?.Email ?? "", plan.Length > 0 ? plan : old?.Plan ?? "");
        }
    }

    /// <summary>Opens the sign-in page in the system's browser; false when there is none to open (the page shows the link).</summary>
    public static bool OpenBrowser(string url)
    {
        if (Environment.GetEnvironmentVariable("WRITER_NO_BROWSER") is { Length: > 0 }) return false;
        try
        {
            var start = OperatingSystem.IsMacOS() ? new System.Diagnostics.ProcessStartInfo("open", url)
                : OperatingSystem.IsWindows() ? new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true }
                : new System.Diagnostics.ProcessStartInfo("xdg-open", url);
            using var _ = System.Diagnostics.Process.Start(start);
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }
}

/// <summary>A ChatGPT sign-in's tokens: expires is when the access token lapses (Unix ms); the account id goes with every request.</summary>
public sealed record ChatGptTokens(string Access, string Refresh, long Expires, string AccountId, string Email, string Plan)
{
    /// <summary>A record prints every member; this one must never print a token.</summary>
    public override string ToString() => $"ChatGPT {Email} ({Plan})";

    /// <summary>Within five minutes of lapsing (the Codex CLI's window): refresh before the next request.</summary>
    public bool Stale => DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() > Expires - 5 * 60 * 1000;

    public JsonObject ToJson() => new() { ["access"] = Access, ["refresh"] = Refresh, ["expires"] = Expires, ["accountId"] = AccountId, ["email"] = Email, ["plan"] = Plan };

    public static ChatGptTokens? FromJson(JsonNode? node)
    {
        if (node is not JsonObject o) return null;
        string S(string k) => o[k] is JsonValue v && v.TryGetValue<string>(out var s) ? s : "";
        var expires = o["expires"] is JsonValue e && e.TryGetValue<long>(out var ms) ? ms : 0;
        return S("refresh").Length == 0 && S("access").Length == 0 ? null : new ChatGptTokens(S("access"), S("refresh"), expires, S("accountId"), S("email"), S("plan"));
    }
}

/// <summary>One sign-in under way: a loopback listener on the Codex CLI's callback port waits for the browser to come back with the
/// code, which is exchanged for tokens and handed to <c>onTokens</c>. Ten minutes, then it gives up.</summary>
public sealed class ChatGptLogin : IDisposable
{
    readonly CancellationTokenSource _stop = new();
    readonly List<TcpListener> _listeners = [];
    readonly string _verifier, _state;
    readonly int _port;
    int _exchanging;
    readonly HttpClient _http;
    readonly Func<ChatGptTokens, Task> _onTokens;

    public string Url { get; }
    /// <summary>waiting, done, error or cancelled.</summary>
    public string State { get; private set; } = "waiting";
    public string Error { get; private set; } = "";
    public ChatGptTokens? Tokens { get; private set; }

    /// <param name="port">The callback port: ChatGpt.Port for the real service; 0 takes a free one (tests).</param>
    public ChatGptLogin(HttpMessageHandler? handler, int port, Func<ChatGptTokens, Task> onTokens)
    {
        (_verifier, var challenge) = ChatGpt.Pkce();
        _state = ChatGpt.Base64Url(RandomNumberGenerator.GetBytes(24));
        _onTokens = onTokens;
        foreach (var address in new[] { IPAddress.Loopback, IPAddress.IPv6Loopback })
        {
            try
            {
                var l = new TcpListener(address, port);
                l.Start();
                _listeners.Add(l);
                port = ((IPEndPoint)l.LocalEndpoint).Port; // ::1 on the same port
            }
            catch (SocketException) when (address.Equals(IPAddress.IPv6Loopback))
            {
                // no IPv6 loopback here: 127.0.0.1 is enough
            }
            catch (SocketException ex)
            {
                throw new WriterException(ErrorCode.Io, $"登录要用的端口 {port} 被占用", "先关掉另一个正在进行的登录（例如 Codex CLI 的），再试一次。" + ex.Message);
            }
        }
        _port = port;
        _http = new HttpClient(handler ?? new SocketsHttpHandler(), disposeHandler: handler is null) { Timeout = TimeSpan.FromSeconds(60) };
        Url = ChatGpt.AuthorizeUrl(challenge, _state, port);
        foreach (var l in _listeners) _ = Accept(l);
        _ = Task.Delay(TimeSpan.FromMinutes(10), _stop.Token).ContinueWith(t => { if (!t.IsCanceled && State == "waiting") Finish("error", "登录超时，请重试"); }, TaskScheduler.Default);
    }

    async Task Accept(TcpListener listener)
    {
        while (!_stop.IsCancellationRequested)
        {
            TcpClient client;
            try
            {
                client = await listener.AcceptTcpClientAsync(_stop.Token);
            }
            catch (Exception)
            {
                return;
            }
            _ = Answer(client);
        }
    }

    /// <summary>One request from the browser: the callback gets the result page; anything else a 404.</summary>
    async Task Answer(TcpClient client)
    {
        using (client)
        {
            try
            {
                var stream = client.GetStream();
                var head = await ReadHead(stream);
                var target = head.Split(' ') is { Length: >= 2 } parts ? parts[1] : "/";
                var uri = new Uri("http://localhost" + target);
                if (uri.AbsolutePath != ChatGpt.CallbackPath)
                {
                    await Respond(stream, "404 Not Found", "");
                    return;
                }
                var q = System.Web.HttpUtility.ParseQueryString(uri.Query);
                string page;
                if (State != "waiting") page = ResultPage(State == "done", State == "done" ? "" : Error);
                // not this sign-in's answer (another tab, a forged link): this one keeps waiting for its own
                else if (q["state"] != _state) page = ResultPage(false, "这个登录链接已经过期，请回到 Writer 重新登录。");
                else if (q["error"] is { } err) { Finish("error", "ChatGPT 登录被取消或拒绝：" + (q["error_description"] ?? err)); page = ResultPage(false, Error); }
                else if (q["code"] is not { Length: > 0 } code) { Finish("error", "登录回来的信息不全，请重试"); page = ResultPage(false, Error); }
                else if (Interlocked.Exchange(ref _exchanging, 1) == 1) page = ResultPage(false, "正在完成登录，请回到 Writer 查看。"); // the browser asked twice
                else
                {
                    try
                    {
                        var tokens = await ChatGpt.Exchange(_http, code, _verifier, _port, _stop.Token);
                        await _onTokens(tokens);
                        Tokens = tokens;
                        Finish("done", "");
                        page = ResultPage(true, "");
                    }
                    catch (WriterException ex)
                    {
                        Finish("error", ex.Message + (ex.Hint.Length > 0 ? "：" + ex.Hint : ""));
                        page = ResultPage(false, Error);
                    }
                }
                await Respond(stream, "200 OK", page);
            }
            catch (Exception)
            {
                // the browser went away
            }
        }
    }

    static async Task<string> ReadHead(NetworkStream stream)
    {
        var buffer = new byte[8192];
        var total = 0;
        while (total < buffer.Length)
        {
            var n = await stream.ReadAsync(buffer.AsMemory(total));
            if (n == 0) break;
            total += n;
            if (Encoding.ASCII.GetString(buffer, 0, total).Contains("\r\n\r\n", StringComparison.Ordinal)) break;
        }
        var text = Encoding.ASCII.GetString(buffer, 0, total);
        var end = text.IndexOf("\r\n", StringComparison.Ordinal);
        return end < 0 ? text : text[..end];
    }

    static async Task Respond(NetworkStream stream, string status, string html)
    {
        var body = Encoding.UTF8.GetBytes(html);
        var head = Encoding.ASCII.GetBytes($"HTTP/1.1 {status}\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {body.Length}\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(head);
        await stream.WriteAsync(body);
        await stream.FlushAsync();
    }

    static string ResultPage(bool ok, string error) =>
        "<!doctype html><meta charset=\"utf-8\"><title>Writer</title><body style=\"font:15px -apple-system,'PingFang SC',system-ui,sans-serif;display:flex;align-items:center;justify-content:center;height:90vh;color:#1D1D1F\"><div style=\"text-align:center\">"
        + (ok ? "<div style=\"font-size:20px;font-weight:600;margin-bottom:8px\">已登录 ChatGPT</div><div style=\"color:#6E6E73\">可以关掉这个页面，回到 Writer 了。<br>Signed in. You can close this page and go back to Writer.</div>"
              : "<div style=\"font-size:20px;font-weight:600;margin-bottom:8px\">没有登录成功</div><div style=\"color:#6E6E73\">" + WebUtility.HtmlEncode(error) + "<br>回到 Writer 再试一次。</div>")
        + "</div></body>";

    void Finish(string state, string error)
    {
        if (State != "waiting") return;
        State = state;
        Error = error;
        // a moment for the result page to reach the browser, then the port is free again
        _ = Task.Delay(1500).ContinueWith(_ => Dispose(), TaskScheduler.Default);
    }

    public void Cancel() => Finish("cancelled", "");

    public void Dispose()
    {
        if (!_stop.IsCancellationRequested) _stop.Cancel();
        foreach (var l in _listeners)
        {
            try { l.Stop(); } catch (Exception) { }
        }
        _http.Dispose();
    }
}
