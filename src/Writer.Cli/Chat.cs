using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Writer.Core;
using Writer.Formats;

namespace Writer.Cli;

/// <summary>The AI panel's back end: an agent loop over the writer command line (one command, or an atomic batch of them) and a
/// scratch plan, over one of two wire formats — the Anthropic Messages API, or an OpenAI-compatible /chat/completions (ChatOpenAi.cs).
/// Each turn is one API call; text streams to the UI as it arrives, and tool calls reach it as the same events either way while the
/// file is edited in place.</summary>
public sealed partial class Chat
{
    public const int MaxSteps = 24;
    public const string DefaultModel = "claude-sonnet-5";
    public const string BatchToolName = Assistant.BatchToolName;
    public const string PlanToolName = Assistant.PlanToolName;

    public const int OutlineBudget = Assistant.OutlineBudget;

    /// <summary>Never follows a redirect: the key goes to the configured address and nowhere else.</summary>
    static readonly HttpMessageHandler Direct = new SocketsHttpHandler { AllowAutoRedirect = false };
    static readonly Regex ToolsRefused = new(@"\btools?\b|function.?call", RegexOptions.IgnoreCase);

    readonly HttpClient _http;
    readonly string _apiKey;
    readonly string _baseUrl;
    readonly bool _openAi;
    readonly string _provider;

    public string Model { get; }

    /// <summary>The autocomplete's model (ChatComplete.cs); "" means <see cref="Model"/>.</summary>
    public string CompleteModel { get; init; } = "";

    /// <param name="provider">An id of <see cref="AiConfig.Providers"/>: how to turn a model's thinking off for the autocomplete.</param>
    public Chat(string apiKey, string model, string baseUrl, HttpMessageHandler? handler = null, bool openAi = false, string provider = "")
    {
        _apiKey = apiKey;
        Model = model;
        _baseUrl = baseUrl.TrimEnd('/');
        _openAi = openAi;
        _provider = provider;
        _http = new HttpClient(handler ?? Direct, disposeHandler: false) { Timeout = TimeSpan.FromSeconds(180) };
    }

    /// <summary>One model reply: its text and tool calls in order, the assistant message for the transcript, and whether the
    /// model waits for tool results. In a block, Tool is the advertised name (writer, batch, plan, web_search, web_fetch) and Input its
    /// JSON arguments when Id is set (a tool call); both are null for a plain text block.</summary>
    sealed record Reply(List<(string? Text, string? Id, string? Tool, JsonObject? Input)> Blocks, JsonObject Message, bool WantsTools);

    /// <summary>Runs one user turn. <paramref name="history"/> is the UI's transcript: [{role, content}] with the new user message last.
    /// <paramref name="emit"/> receives (event, json): delta {text} as the reply streams, text {text} once a text block is whole,
    /// tool {command, code, output, wrote}, done {steps}, error {message, hint}. <paramref name="instructions"/> (the user's preferences)
    /// are appended to the system prompt. A non-null <paramref name="selection"/> limits the context to that text: it is quoted before
    /// the new message and the document outline is left out of the prompt. <paramref name="web"/> (联网搜索, on by default) offers
    /// web_search and web_fetch alongside the editing tools.</summary>
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
        Task Delta(string text) => emit("delta", Json(new JsonObject { ["text"] = text }));
        var steps = 0;
        for (; steps < MaxSteps; steps++)
        {
            ct.ThrowIfCancellationRequested();
            Reply reply;
            try
            {
                reply = Codex ? await CodexTurn(system, messages, web, Delta, ct)
                    : _openAi ? await OpenAiTurn(system, messages, web, Delta, ct) : await AnthropicTurn(system, messages, web, Delta, ct);
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
                var (display, code, output, wrote) = await RunTool(tool, input, workspace, seen, ct);
                var shown = output.Length > 6000 ? output[..6000] + "\n…(truncated)" : output;
                await emit("tool", Json(new JsonObject { ["command"] = display, ["code"] = code, ["output"] = shown, ["wrote"] = wrote }));
                results.Add((id, shown.Length == 0 ? "(no output)" : shown, code != 0));
            }
            messages.Add((JsonNode)reply.Message);
            if (!reply.WantsTools || results.Count == 0) break;
            if (_openAi || Codex)
                foreach (var (id, output, _) in results)
                    messages.Add((JsonNode)new JsonObject { ["role"] = "tool", ["tool_call_id"] = id, ["content"] = output });
            else
                messages.Add((JsonNode)new JsonObject
                {
                    ["role"] = "user",
                    ["content"] = new JsonArray(results.Select(x => (JsonNode)new JsonObject { ["type"] = "tool_result", ["tool_use_id"] = x.Id, ["content"] = x.Output, ["is_error"] = x.Failed }).ToArray()),
                });
            if (steps == MaxSteps - 1) await emit("text", Json(new JsonObject { ["text"] = "已达到本轮的步数上限，先停在这里；回复「继续」可以接着做。" }));
        }
        await emit("done", Json(new JsonObject { ["steps"] = Math.Min(steps + 1, MaxSteps) }));
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
            _ = Codex ? await CodexTurn(system, messages, false, null, limit.Token)
                : _openAi ? await OpenAiTurn(system, messages, false, null, limit.Token) : await AnthropicTurn(system, messages, false, null, limit.Token);
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

    /// <summary>POST /v1/messages with stream: true. Text deltas go to <paramref name="onDelta"/> as they arrive. A server that answers with
    /// one JSON body instead (a proxy that ignores stream) is read the plain way.</summary>
    async Task<Reply> AnthropicTurn(string system, JsonArray messages, bool web, Func<string, Task>? onDelta, CancellationToken ct)
    {
        var body = new JsonObject
        {
            ["model"] = Model,
            ["max_tokens"] = 8192,
            ["system"] = system,
            ["tools"] = AnthropicTools(web),
            ["messages"] = messages.DeepClone(),
            ["stream"] = true,
        };
        using var request = new HttpRequestMessage(HttpMethod.Post, _baseUrl + "/v1/messages") { Content = new StringContent(Json(body), Encoding.UTF8, "application/json") };
        request.Headers.Add("x-api-key", _apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        using var response = await Send(request, HttpCompletionOption.ResponseHeadersRead, ct);
        if (response.Content.Headers.ContentType?.MediaType == "text/event-stream")
        {
            var (streamed, stop) = await AnthropicStream(response, onDelta, ct);
            return FromContent(streamed, stop);
        }
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
        return FromContent(reply["content"] as JsonArray ?? [], reply["stop_reason"]?.GetValue<string>());
    }

    /// <summary>The Messages API's event stream folded back into the content array a plain reply carries: text blocks from their
    /// text_delta pieces, tool_use blocks from their input_json_delta pieces, the stop reason from message_delta.</summary>
    async Task<(JsonArray Content, string? Stop)> AnthropicStream(HttpResponseMessage response, Func<string, Task>? onDelta, CancellationToken ct)
    {
        var blocks = new SortedDictionary<int, (string Type, string? Id, string? Name, StringBuilder Text)>();
        string? stop = null;
        var streamed = false;
        using var idle = CancellationTokenSource.CreateLinkedTokenSource(ct);
        try
        {
            using var reader = new StreamReader(await response.Content.ReadAsStreamAsync(idle.Token));
            while (true)
            {
                idle.CancelAfter(Idle);
                var line = await reader.ReadLineAsync(idle.Token);
                if (line is null) break;
                if (!line.StartsWith("data:", StringComparison.Ordinal)) continue;
                var data = line[5..].Trim();
                JsonNode? node;
                try
                {
                    node = JsonNode.Parse(data);
                }
                catch (JsonException)
                {
                    continue;
                }
                if (node is not JsonObject e) continue;
                streamed = true;
                var index = e["index"] is JsonValue v && v.TryGetValue<int>(out var n) ? n : 0;
                switch (Str(e["type"]))
                {
                    case "content_block_start":
                        var start = e["content_block"] as JsonObject;
                        var text = Str(start?["text"]) ?? "";
                        blocks[index] = (Str(start?["type"]) ?? "text", Str(start?["id"]), Str(start?["name"]), new StringBuilder(text));
                        if (text.Length > 0 && onDelta is not null) await onDelta(text);
                        break;
                    case "content_block_delta":
                        if (!blocks.TryGetValue(index, out var block)) break;
                        var delta = e["delta"] as JsonObject;
                        var piece = Str(delta?["text"]) ?? Str(delta?["partial_json"]);
                        if (piece is null) break;
                        block.Text.Append(piece);
                        if (block.Type == "text" && onDelta is not null) await onDelta(piece);
                        break;
                    case "message_delta":
                        stop = Str((e["delta"] as JsonObject)?["stop_reason"]) ?? stop;
                        break;
                    case "error":
                        throw Failure(0, ErrorText(data));
                }
            }
        }
        catch (Exception ex) when (ex is IOException or HttpRequestException)
        {
            throw Unreachable(ex);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            throw TimedOut();
        }
        if (!streamed) throw NotAnApi("");
        var content = new JsonArray();
        foreach (var (_, block) in blocks)
        {
            if (block.Type == "text") content.Add((JsonNode)new JsonObject { ["type"] = "text", ["text"] = block.Text.ToString() });
            else if (block.Type == "tool_use")
            {
                JsonObject? input = null;
                try
                {
                    input = JsonNode.Parse(block.Text.Length == 0 ? "{}" : block.Text.ToString()) as JsonObject;
                }
                catch (JsonException)
                {
                    // the tool result says what was wrong with the arguments
                }
                content.Add((JsonNode)new JsonObject { ["type"] = "tool_use", ["id"] = block.Id ?? "", ["name"] = block.Name ?? Mcp.ToolName, ["input"] = input ?? new JsonObject() });
            }
        }
        return (content, stop);
    }

    static Reply FromContent(JsonArray content, string? stopReason)
    {
        var blocks = new List<(string?, string?, string?, JsonObject?)>();
        foreach (var block in content)
        {
            var type = block?["type"]?.GetValue<string>();
            if (type == "text") blocks.Add((block!["text"]?.GetValue<string>() ?? "", null, null, null));
            else if (type == "tool_use") blocks.Add((null, block!["id"]?.GetValue<string>() ?? "", block["name"]?.GetValue<string>(), block["input"] as JsonObject));
        }
        return new Reply(blocks, new JsonObject { ["role"] = "assistant", ["content"] = content.DeepClone() }, stopReason == "tool_use");
    }

    /// <summary>The tools offered on /v1/messages: writer, batch and plan always, web_search and web_fetch when 联网搜索 is on.</summary>
    static JsonArray AnthropicTools(bool web)
    {
        // JsonNode-typed adds: JsonArray.Add<T> for a JsonObject is the reflection overload, which the NativeAOT release engine refuses
        var tools = new JsonArray();
        foreach (var (name, description, schema) in EditingTools())
            tools.Add((JsonNode)new JsonObject { ["name"] = name, ["description"] = description, ["input_schema"] = schema });
        if (web)
        {
            tools.Add((JsonNode)new JsonObject { ["name"] = Web.SearchToolName, ["description"] = Web.SearchDescription, ["input_schema"] = Web.SearchParameters() });
            tools.Add((JsonNode)new JsonObject { ["name"] = Web.FetchToolName, ["description"] = Web.FetchDescription, ["input_schema"] = Web.FetchParameters() });
        }
        return tools;
    }

    static IEnumerable<(string Name, string Description, JsonObject Schema)> EditingTools() => Assistant.EditingTools();

    /// <summary>Runs one tool call and returns what to show in the tool event: a display line (the writer command as given, a batch's
    /// label, 计划, or 搜索/打开 for the web tools), the exit code, the output text, and whether the file was written. Newly seen URLs
    /// (a search's results, or the links printed in a fetched page) are folded into <paramref name="seen"/> for a later web_fetch call this turn.</summary>
    async Task<(string Display, int Code, string Output, bool Wrote)> RunTool(string? tool, JsonObject? input, string workspace, HashSet<string> seen, CancellationToken ct)
    {
        if (tool == Web.SearchToolName)
        {
            var query = Str(input?[Web.QueryParam]);
            if (string.IsNullOrWhiteSpace(query)) return (Web.SearchToolName, 1, "The tool input must be JSON like {\"query\": \"...\"}.", false);
            var (code, output, urls) = await Web.SearchAsync(query, ct);
            foreach (var url in urls) seen.Add(url);
            return ($"搜索 \"{query}\"", code, output, false);
        }
        if (tool == Web.FetchToolName)
        {
            var url = Str(input?[Web.UrlParam]);
            if (string.IsNullOrWhiteSpace(url)) return (Web.FetchToolName, 1, "The tool input must be JSON like {\"url\": \"https://...\"}.", false);
            var (code, output, urls) = await Web.FetchAsync(url, seen, ct);
            foreach (var u in urls) seen.Add(u);
            return ($"打开 {url}", code, output, false);
        }
        return Assistant.RunEditingTool(tool, input, argv => Serve.RunArgv(argv, workspace));
    }

    public static bool Writes(string[] argv) => Assistant.Writes(argv);

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

    string Redact(string s)
    {
        if (_apiKey.Length >= 8) s = s.Replace(_apiKey, "***", StringComparison.Ordinal);
        return _access.Length >= 8 ? s.Replace(_access, "***", StringComparison.Ordinal) : s;
    }

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

    public static string SystemPrompt(string? file, string workspace, bool outline = true, bool web = false) => Assistant.SystemPrompt(file, workspace, outline, web);

    static string Trim(string s, int max) => s.Length <= max ? s : s[..max] + "\n…(truncated)";

    static string Json(JsonNode node) => node.ToJsonString(new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
}
