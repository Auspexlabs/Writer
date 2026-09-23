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

    /// <summary>One model reply: its text and tool calls in order (a call has an Id; a null Command means unusable input),
    /// the assistant message for the transcript, and whether the model waits for tool results.</summary>
    sealed record Reply(List<(string? Text, string? Id, string? Command)> Blocks, JsonObject Message, bool WantsTools);

    /// <summary>Runs one user turn. <paramref name="history"/> is the UI's transcript: [{role, content}] with the new user message last.
    /// <paramref name="emit"/> receives (event, json): text {text}, tool {command, code, output}, done {steps}, error {message, hint}.
    /// <paramref name="instructions"/> (the user's preferences) are appended to the system prompt. A non-null <paramref name="selection"/>
    /// limits the context to that text: it is quoted before the new message and the document outline is left out of the prompt.</summary>
    public async Task RunAsync(string? file, string workspace, JsonElement history, Func<string, string, Task> emit, CancellationToken ct,
        string? instructions = null, string? selection = null)
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
        var system = SystemPrompt(file, workspace, outline: selection is null);
        if (!string.IsNullOrWhiteSpace(instructions)) system += "\n\n## The user's preferences\n\n" + instructions.Trim() + "\n";
        var steps = 0;
        for (; steps < MaxSteps; steps++)
        {
            ct.ThrowIfCancellationRequested();
            Reply reply;
            try
            {
                reply = _openAi ? await OpenAiTurn(system, messages, ct) : await AnthropicTurn(system, messages, ct);
            }
            catch (WriterException ex)
            {
                await emit("error", Json(new JsonObject { ["message"] = ex.Message, ["hint"] = ex.Hint }));
                return;
            }
            var results = new List<(string Id, string Output, bool Failed)>();
            foreach (var (text, id, command) in reply.Blocks)
            {
                if (id is null)
                {
                    await emit("text", Json(new JsonObject { ["text"] = text }));
                    continue;
                }
                var (code, output) = command is null ? (1, "The tool input must be JSON like {\"command\": \"view report.docx outline\"}.")
                    : Serve.RunArgv(Mcp.Tokenize(command), workspace);
                var shown = output.Length > 6000 ? output[..6000] + "\n…(truncated)" : output;
                await emit("tool", Json(new JsonObject { ["command"] = command ?? "", ["code"] = code, ["output"] = shown }));
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
            _ = _openAi ? await OpenAiTurn(system, messages, limit.Token) : await AnthropicTurn(system, messages, limit.Token);
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

    async Task<Reply> AnthropicTurn(string system, JsonArray messages, CancellationToken ct)
    {
        var body = new JsonObject
        {
            ["model"] = Model,
            ["max_tokens"] = 8192,
            ["system"] = system,
            ["tools"] = new JsonArray(new JsonObject { ["name"] = Mcp.ToolName, ["description"] = Mcp.ToolDescription, ["input_schema"] = Parameters() }),
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
        var blocks = new List<(string?, string?, string?)>();
        foreach (var block in content)
        {
            var type = block?["type"]?.GetValue<string>();
            if (type == "text") blocks.Add((block!["text"]?.GetValue<string>() ?? "", null, null));
            else if (type == "tool_use") blocks.Add((null, block!["id"]?.GetValue<string>() ?? "", block["input"]?["command"]?.GetValue<string>() ?? ""));
        }
        return new Reply(blocks, new JsonObject { ["role"] = "assistant", ["content"] = content.DeepClone() }, reply["stop_reason"]?.GetValue<string>() == "tool_use");
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
    public static string SystemPrompt(string? file, string workspace, bool outline = true)
    {
        var sb = new StringBuilder();
        sb.Append("You are the assistant built into a document editor. The user sees the document on the left and talks to you on the right. ");
        sb.Append("You change documents only through the writer tool, which edits the file in place; the editor reloads it after every command. ");
        sb.Append("Look before you change: start with `view <file> outline` (or `get`/`query`) so you address the right paths, then make the smallest edit that does the job. ");
        sb.Append("Use `help <format> <element>` when unsure which properties exist. Prefer `--prop html=` or `--prop md=` for formatted text. ");
        sb.Append("Never tell the user to run commands themselves. When the user asks a question, answer it from the document. ");
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
