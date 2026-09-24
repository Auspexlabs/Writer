using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Writer.Core;
using Writer.Formats;

namespace Writer.Cli;

/// <summary>The AI panel's back end: an agent loop whose only tool is the writer command line, over one of two wire formats —
/// the Anthropic Messages API, or an OpenAI-compatible /chat/completions (ChatOpenAi.cs). Each turn is one API call; text and
/// tool calls reach the UI as the same events either way while the file is edited in place.</summary>
public sealed partial class Chat
{
    public const int MaxSteps = 16;
    public const string DefaultModel = "claude-sonnet-5";

    /// <summary>Never follows a redirect: the key goes to the configured address and nowhere else.</summary>
    static readonly HttpMessageHandler Direct = new SocketsHttpHandler { AllowAutoRedirect = false };
    static readonly Regex ToolsRefused = new(@"\btools?\b|function.?call", RegexOptions.IgnoreCase);

    readonly HttpClient _http;
    readonly string _apiKey;
    readonly string _baseUrl;
    readonly bool _openAi;

    public string Model { get; }

    public Chat(string apiKey, string model, string baseUrl, HttpMessageHandler? handler = null, bool openAi = false)
    {
        _apiKey = apiKey;
        Model = model;
        _baseUrl = baseUrl.TrimEnd('/');
        _openAi = openAi;
        _http = new HttpClient(handler ?? Direct, disposeHandler: false) { Timeout = TimeSpan.FromSeconds(180) };
    }

    /// <summary>One model reply: its text and tool calls in order, the assistant message for the transcript, and whether the
    /// model waits for tool results. In a block, Tool is the advertised name (writer, web_search, web_fetch) and Input its
    /// JSON arguments when Id is set (a tool call); both are null for a plain text block.</summary>
    sealed record Reply(List<(string? Text, string? Id, string? Tool, JsonObject? Input)> Blocks, JsonObject Message, bool WantsTools);

    /// <summary>Runs one user turn. <paramref name="history"/> is the UI's transcript: [{role, content}] with the new user message last.
    /// <paramref name="emit"/> receives (event, json): text {text}, tool {command, code, output}, done {steps}, error {message, hint}.
    /// <paramref name="instructions"/> (the user's preferences) are appended to the system prompt. A non-null <paramref name="selection"/>
    /// limits the context to that text: it is quoted before the new message and the document outline is left out of the prompt.
    /// <paramref name="web"/> (联网搜索, on by default) offers web_search and web_fetch alongside writer.</summary>
    public async Task RunAsync(string? file, string workspace, JsonElement history, Func<string, string, Task> emit, CancellationToken ct,
        string? instructions = null, string? selection = null, bool web = true)
    {
        var messages = new JsonArray();
        foreach (var m in history.EnumerateArray())
        {
            var role = m.TryGetProperty("role", out var r) ? r.GetString() : null;
            var content = m.TryGetProperty("content", out var c) ? c.GetString() : null;
            if (role is not ("user" or "assistant") || string.IsNullOrWhiteSpace(content)) continue;
            messages.Add((JsonNode)new JsonObject { ["role"] = role, ["content"] = content });
        }
        if (messages.Count == 0 || messages[^1]!["role"]!.GetValue<string>() != "user")
        {
            await emit("error", Json(new JsonObject { ["message"] = "The last message must be from the user" }));
            return;
        }
        if (!string.IsNullOrWhiteSpace(selection))
            messages[^1]!["content"] = "The user selected this part of the document:\n<selection>\n" + selection.Trim() + "\n</selection>\n\n" + messages[^1]!["content"]!.GetValue<string>();
        var system = SystemPrompt(file, workspace, outline: selection is null, web: web);
        if (!string.IsNullOrWhiteSpace(instructions)) system += "\n\n## The user's preferences\n\n" + instructions.Trim() + "\n";
        // Seeded once, up front, from what the user and the document already put in front of the model; web_search/web_fetch
        // results grow it as the turn runs. web_fetch may only target a URL already in this set — see RunTool.
        var seen = new HashSet<string>(StringComparer.Ordinal);
        if (web)
        {
            foreach (var url in Web.ExtractUrls(system)) seen.Add(url);
            foreach (var m in messages)
                if (Str(m?["role"]) == "user" && Str(m?["content"]) is { } content)
                    foreach (var url in Web.ExtractUrls(content)) seen.Add(url);
        }
        var steps = 0;
        for (; steps < MaxSteps; steps++)
        {
            ct.ThrowIfCancellationRequested();
            Reply reply;
            try
            {
                reply = _openAi ? await OpenAiTurn(system, messages, web, ct) : await AnthropicTurn(system, messages, web, ct);
            }
            catch (WriterException ex)
            {
                await emit("error", Json(new JsonObject { ["message"] = ex.Message, ["hint"] = ex.Hint }));
                return;
            }
            var results = new List<(string Id, string Output, bool Failed)>();
            foreach (var (text, id, tool, input) in reply.Blocks)
            {
                ct.ThrowIfCancellationRequested(); // the client may have gone away mid-reply: stop before the next text or tool block
                if (id is null)
                {
                    await emit("text", Json(new JsonObject { ["text"] = text }));
                    continue;
                }
                var (display, code, output) = await RunTool(tool, input, workspace, seen, ct);
                var shown = output.Length > 6000 ? output[..6000] + "\n…(truncated)" : output;
                await emit("tool", Json(new JsonObject { ["command"] = display, ["code"] = code, ["output"] = shown }));
                results.Add((id, shown.Length == 0 ? "(no output)" : shown, code != 0));
            }
            messages.Add((JsonNode)reply.Message);
            if (!reply.WantsTools || results.Count == 0) break;
            if (_openAi)
                foreach (var (id, output, _) in results)
                    messages.Add((JsonNode)new JsonObject { ["role"] = "tool", ["tool_call_id"] = id, ["content"] = output });
            else
                messages.Add((JsonNode)new JsonObject
                {
                    ["role"] = "user",
                    ["content"] = new JsonArray(results.Select(x => (JsonNode)new JsonObject { ["type"] = "tool_result", ["tool_use_id"] = x.Id, ["content"] = x.Output, ["is_error"] = x.Failed }).ToArray()),
                });
        }
        await emit("done", Json(new JsonObject { ["steps"] = steps + 1 }));
    }

    /// <summary>One small request with the tool attached: null when the model answered, otherwise why not (short, Chinese).</summary>
    public async Task<string?> TestAsync(CancellationToken ct)
    {
        using var limit = CancellationTokenSource.CreateLinkedTokenSource(ct);
        limit.CancelAfter(TimeSpan.FromSeconds(60));
        const string system = "This is a connection test.";
        var messages = new JsonArray(new JsonObject { ["role"] = "user", ["content"] = "Reply with one word: ok" });
        try
        {
            _ = _openAi ? await OpenAiTurn(system, messages, false, limit.Token) : await AnthropicTurn(system, messages, false, limit.Token);
            return null;
        }
        catch (WriterException ex)
        {
            return ex.Message;
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            return TimedOut().Message;
        }
    }

    async Task<Reply> AnthropicTurn(string system, JsonArray messages, bool web, CancellationToken ct)
    {
        var body = new JsonObject
        {
            ["model"] = Model,
            ["max_tokens"] = 8192,
            ["system"] = system,
            ["tools"] = AnthropicTools(web),
            ["messages"] = messages.DeepClone(),
        };
        using var request = new HttpRequestMessage(HttpMethod.Post, _baseUrl + "/v1/messages") { Content = new StringContent(Json(body), Encoding.UTF8, "application/json") };
        request.Headers.Add("x-api-key", _apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        using var response = await Send(request, HttpCompletionOption.ResponseContentRead, ct);
        var text = await response.Content.ReadAsStringAsync(ct);
        JsonObject reply;
        try
        {
            reply = JsonNode.Parse(text) as JsonObject ?? throw new JsonException();
        }
        catch (JsonException)
        {
            throw NotAnApi(text);
        }
        var content = reply["content"] as JsonArray ?? [];
        var blocks = new List<(string?, string?, string?, JsonObject?)>();
        foreach (var block in content)
        {
            var type = block?["type"]?.GetValue<string>();
            if (type == "text") blocks.Add((block!["text"]?.GetValue<string>() ?? "", null, null, null));
            else if (type == "tool_use") blocks.Add((null, block!["id"]?.GetValue<string>() ?? "", block["name"]?.GetValue<string>(), block["input"] as JsonObject));
        }
        return new Reply(blocks, new JsonObject { ["role"] = "assistant", ["content"] = content.DeepClone() }, reply["stop_reason"]?.GetValue<string>() == "tool_use");
    }

    /// <summary>The tools offered on /v1/messages: writer always, web_search and web_fetch when 联网搜索 is on.</summary>
    static JsonArray AnthropicTools(bool web)
    {
        // JsonNode-typed adds: JsonArray.Add<T> for a JsonObject is the reflection overload, which the NativeAOT release engine refuses
        var tools = new JsonArray(new JsonObject { ["name"] = Mcp.ToolName, ["description"] = Mcp.ToolDescription, ["input_schema"] = Parameters() });
        if (web)
        {
            tools.Add((JsonNode)new JsonObject { ["name"] = Web.SearchToolName, ["description"] = Web.SearchDescription, ["input_schema"] = Web.SearchParameters() });
            tools.Add((JsonNode)new JsonObject { ["name"] = Web.FetchToolName, ["description"] = Web.FetchDescription, ["input_schema"] = Web.FetchParameters() });
        }
        return tools;
    }

    /// <summary>Runs one tool call and returns what to show in the tool event: a display command line (the writer command as
    /// given, or 搜索/打开 for the web tools), the exit code, and the output text. Newly seen URLs (a search's results, or the
    /// links printed in a fetched page) are folded into <paramref name="seen"/> for a later web_fetch call this turn.</summary>
    async Task<(string Display, int Code, string Output)> RunTool(string? tool, JsonObject? input, string workspace, HashSet<string> seen, CancellationToken ct)
    {
        if (tool == Web.SearchToolName)
        {
            var query = Str(input?[Web.QueryParam]);
            if (string.IsNullOrWhiteSpace(query)) return (Web.SearchToolName, 1, "The tool input must be JSON like {\"query\": \"...\"}.");
            var (code, output, urls) = await Web.SearchAsync(query, ct);
            foreach (var url in urls) seen.Add(url);
            return ($"搜索 \"{query}\"", code, output);
        }
        if (tool == Web.FetchToolName)
        {
            var url = Str(input?[Web.UrlParam]);
            if (string.IsNullOrWhiteSpace(url)) return (Web.FetchToolName, 1, "The tool input must be JSON like {\"url\": \"https://...\"}.");
            var (code, output, urls) = await Web.FetchAsync(url, seen, ct);
            foreach (var u in urls) seen.Add(u);
            return ($"打开 {url}", code, output);
        }
        var command = Str(input?["command"]);
        var (writerCode, writerOutput) = command is null ? (1, "The tool input must be JSON like {\"command\": \"view report.docx outline\"}.")
            : Serve.RunArgv(Mcp.Tokenize(command), workspace);
        return (command ?? "", writerCode, writerOutput);
    }

    /// <summary>Sends a request. Network trouble, timeouts and error statuses become WriterExceptions; the caller disposes the response.</summary>
    async Task<HttpResponseMessage> Send(HttpRequestMessage request, HttpCompletionOption option, CancellationToken ct)
    {
        try
        {
            var response = await _http.SendAsync(request, option, ct);
            if (response.IsSuccessStatusCode) return response;
            using (response) throw Failure((int)response.StatusCode, ErrorText(await response.Content.ReadAsStringAsync(ct)));
        }
        catch (Exception ex) when (ex is HttpRequestException or IOException)
        {
            throw Unreachable(ex);
        }
        catch (TaskCanceledException) when (!ct.IsCancellationRequested)
        {
            throw TimedOut();
        }
    }

    /// <summary>A failed API call as a short Chinese message; the provider's own words, with the key cut out, go in the hint.</summary>
    WriterException Failure(int status, string detail)
    {
        detail = Trim(Redact(detail), 300).Trim();
        if (status == 401) return new(ErrorCode.Io, "Key 无效", detail);
        if (status == 404) return new(ErrorCode.Io, "地址或模型不对", detail);
        // ponytail: recognised by the provider's wording ("does not support tools", "Function Calling"); a new wording falls through to the generic message
        if (status is >= 400 and < 500 && ToolsRefused.IsMatch(detail)) return new(ErrorCode.Io, "这个模型不支持工具调用，请换一个支持 function calling 的模型", detail);
        return new(ErrorCode.Io, (status > 0 ? $"模型服务出错（{status}）：" : "模型服务出错：") + detail, "");
    }

    WriterException Unreachable(Exception ex) => new(ErrorCode.Io, "连不上", Redact(ex.Message));

    static WriterException TimedOut() => new(ErrorCode.Io, "模型服务没有响应（超时）", "");

    WriterException NotAnApi(string body) => new(ErrorCode.Io, "接口地址不对：那里返回的不是模型接口的数据", Trim(Redact(body), 200));

    string Redact(string s) => _apiKey.Length >= 8 ? s.Replace(_apiKey, "***", StringComparison.Ordinal) : s;

    /// <summary>The message of an API error body ({"error": {"message"}}, {"error": "..."}, {"message": "..."}), else the body.</summary>
    static string ErrorText(string body)
    {
        try
        {
            if (JsonNode.Parse(body) is JsonObject o)
            {
                var e = o["error"];
                var text = e is JsonValue ? e.GetValue<string>() : (e as JsonObject)?["message"]?.GetValue<string>() ?? o["message"]?.GetValue<string>();
                if (!string.IsNullOrWhiteSpace(text)) return text;
            }
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException)
        {
            // not the usual shape: show the body
        }
        return body;
    }

    /// <summary>The tool's input: one writer command line.</summary>
    static JsonObject Parameters() => new()
    {
        ["type"] = "object",
        ["properties"] = new JsonObject
        {
            ["command"] = new JsonObject { ["type"] = "string", ["description"] = "The writer command line, without the program name." },
        },
        ["required"] = new JsonArray("command"),
    };

    /// <summary>Rules, the command reference for the open file's format, and its current outline.</summary>
    public static string SystemPrompt(string? file, string workspace, bool outline = true, bool web = false)
    {
        var sb = new StringBuilder();
        sb.Append("You are the assistant built into a document editor. The user sees the document on the left and talks to you on the right. ");
        sb.Append("You change documents only through the writer tool, which edits the file in place; the editor reloads it after every command. ");
        sb.Append("Look before you change: start with `view <file> outline` (or `get`/`query`) so you address the right paths, then make the smallest edit that does the job. ");
        sb.Append("Use `help <format> <element>` when unsure which properties exist. Prefer `--prop html=` or `--prop md=` for formatted text. ");
        sb.Append("Never tell the user to run commands themselves. When the user asks a question, answer it from the document. ");
        if (web) sb.Append("Use web_search and web_fetch when the task needs current or outside facts the document doesn't have, and say which URLs you used. ");
        sb.Append("Reply in the user's language, briefly, saying what you changed and where; no command output, no paths unless the user asks.\n\n");
        sb.Append(Mcp.Instructions).Append("\n\n");
        var format = file is null ? null : Adapters.CanonicalFormat(Path.GetExtension(file));
        if (file is not null)
        {
            var relative = Path.IsPathRooted(file) && file.StartsWith(workspace + Path.DirectorySeparatorChar, StringComparison.Ordinal)
                ? Path.GetRelativePath(workspace, file) : file;
            sb.Append("Open file: ").Append(relative).Append(" (use exactly this path in commands).\n");
            if (format is "pdf") sb.Append("PDF files are read-only: read them with `view <file> text`; to edit, `export --to <name>.docx` first and tell the user.\n");
        }
        sb.Append("\n## Command reference\n\n").Append(Run(["help"]));
        if (format is not null and not "pdf")
            foreach (var kind in Registry.ForFormat(format))
                sb.Append('\n').Append(Run(["help", format, kind.Name]));
        if (file is not null && outline)
            sb.Append("\n## Current outline\n\n").Append(Trim(Run(["view", file, "outline"]), 16000));
        return sb.ToString();
    }

    static string Run(string[] argv)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var code = Runner.Run(argv, stdout, stderr);
        return code == 0 ? stdout.ToString() : stderr.ToString();
    }

    static string Trim(string s, int max) => s.Length <= max ? s : s[..max] + "\n…(truncated)";

    static string Json(JsonNode node) => node.ToJsonString(new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
}
