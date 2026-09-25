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
    public const string BatchToolName = "batch";
    public const string PlanToolName = "plan";

    /// <summary>An outline up to this long goes into the prompt whole; a longer document is shown as its structure instead, and the
    /// model reads the parts it needs with section, search and get.</summary>
    public const int OutlineBudget = 16000;

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
                reply = _openAi ? await OpenAiTurn(system, messages, web, Delta, ct) : await AnthropicTurn(system, messages, web, Delta, ct);
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
            if (_openAi)
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
            _ = _openAi ? await OpenAiTurn(system, messages, false, null, limit.Token) : await AnthropicTurn(system, messages, false, null, limit.Token);
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

    /// <summary>The editing tools in order: one writer command, an atomic batch of them, and the scratch plan.</summary>
    static IEnumerable<(string Name, string Description, JsonObject Schema)> EditingTools()
    {
        yield return (Mcp.ToolName, Mcp.ToolDescription, Parameters());
        yield return (BatchToolName,
            "Run several writer commands on the open file as one step: they apply in order and the file is saved once; when any command fails nothing is written and the error names the command. "
            + "Use it for every change that takes more than one command (the edits of one section, a table's rows, a summary row with its formatting). "
            + "label is the step's name in the user's language, shown while it runs, e.g. 正在改第 2 节…",
            new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["commands"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "Writer command lines, without the program name, all on the open file." },
                    ["label"] = new JsonObject { ["type"] = "string", ["description"] = "What this step does, a few words in the user's language." },
                },
                ["required"] = new JsonArray("commands"),
            });
        yield return (PlanToolName,
            "Your scratch plan for a task with several parts: numbered steps, one line each, marked done as you go. Call it before the first edit of such a task and again whenever the plan changes. It writes nothing.",
            new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject { ["text"] = new JsonObject { ["type"] = "string", ["description"] = "The plan, e.g. 1. 读第 2 节 ✓ 2. 改写 3. 补表格" } },
                ["required"] = new JsonArray("text"),
            });
    }

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
        if (tool == PlanToolName)
        {
            var text = Str(input?["text"])?.Trim();
            if (string.IsNullOrWhiteSpace(text)) return (PlanToolName, 1, "The tool input must be JSON like {\"text\": \"1. ... 2. ...\"}.", false);
            return ("计划 " + text.ReplaceLineEndings(" / "), 0, "Plan noted. Work through it and call plan again when it changes.", false);
        }
        if (tool == BatchToolName)
        {
            var commands = (input?["commands"] as JsonArray)?.Select(Str).Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => c!).ToList() ?? [];
            if (commands.Count == 0) return (BatchToolName, 1, "The tool input must be JSON like {\"commands\": [\"set a.docx /body/paragraph[1] --prop text=Hi\"], \"label\": \"...\"}.", false);
            var label = Str(input?["label"])?.Trim() is { Length: > 0 } l ? l : $"批量修改 {commands.Count} 条";
            string[] first;
            try
            {
                first = Mcp.Tokenize(commands[0]);
            }
            catch (WriterException ex)
            {
                return (label, 1, ex.Message + " " + ex.Hint, false);
            }
            if (first.Length > 0 && first[0] == "writer") first = first[1..];
            if (first.Length < 2) return (label, 1, "Each batch command names the file, e.g. set a.docx /body/paragraph[1] --prop text=Hi.", false);
            var argv = new List<string> { "batch", first[1] };
            foreach (var command in commands)
            {
                argv.Add("--run");
                argv.Add(command);
            }
            var (code, output) = Serve.RunArgv(argv.ToArray(), workspace);
            return (label, code, output, code == 0);
        }
        var line = Str(input?["command"]);
        if (line is null) return ("", 1, "The tool input must be JSON like {\"command\": \"view report.docx outline\"}.", false);
        string[] tokens;
        try
        {
            tokens = Mcp.Tokenize(line);
        }
        catch (WriterException ex)
        {
            return (line, 1, ex.Message + " " + ex.Hint, false);
        }
        var (writerCode, writerOutput) = Serve.RunArgv(tokens, workspace);
        return (line, writerCode, writerOutput, writerCode == 0 && Writes(tokens));
    }

    /// <summary>Whether a command line changes the file: the editing verbs, and section, replace and formula unless they only read.</summary>
    public static bool Writes(string[] argv)
    {
        if (argv.Length > 0 && argv[0] == "writer") argv = argv[1..];
        if (argv.Length == 0) return false;
        return argv[0] switch
        {
            "add" or "set" or "remove" or "move" or "copy" or "create" or "export" or "batch" => true,
            "section" => argv.Contains("--md") || argv.Any(t => t.StartsWith("--md=", StringComparison.Ordinal)) || argv.Contains("--remove"),
            "replace" => !argv.Contains("--preview"),
            "formula" => !argv.Contains("--check"),
            _ => false,
        };
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

    /// <summary>The writer tool's input: one command line.</summary>
    static JsonObject Parameters() => new()
    {
        ["type"] = "object",
        ["properties"] = new JsonObject
        {
            ["command"] = new JsonObject { ["type"] = "string", ["description"] = "The writer command line, without the program name." },
        },
        ["required"] = new JsonArray("command"),
    };

    /// <summary>Rules, the rules for the open file's format with examples, the command reference for that format, and its current
    /// outline — or, for a long document and for workbooks, its structure.</summary>
    public static string SystemPrompt(string? file, string workspace, bool outline = true, bool web = false)
    {
        var sb = new StringBuilder();
        sb.Append("You are the assistant built into a document editor. The user sees the document on the left and talks to you on the right. ");
        sb.Append("You change documents only through the tools, which edit the file in place; the editor shows each change as it lands, and at the end of your turn the user gets one card to keep or undo everything you changed.\n\n");
        sb.Append("## How to work\n\n");
        sb.Append("- Look before you change: the outline or structure below gives every path; read the exact part with `section <file> <heading path>`, `search <file> <text>`, `get <file> <path>` or `view <file> text` instead of guessing what it says.\n");
        sb.Append("- Smallest edit that does the job: change the paragraph, cell or section asked for and nothing else. Formatting, styles, numbering and wording elsewhere stay as they are; never rewrite a paragraph to change one word.\n");
        sb.Append("- Text goes into an existing paragraph with `set … --prop text=` (plain), `md=` or `html=` (with inline formatting): the paragraph keeps its style and, in Word, the character formatting of the words that stay. New blocks go in with `add … --after <path>`; a whole section with `section … --md`.\n");
        sb.Append("- Two or more related commands go in one `batch` call with a short label in the user's language (正在改第 2 节…): they land together or not at all. One command: the writer tool. Paths count per kind from the top, so after adding or removing blocks read the structure again before addressing what follows.\n");
        sb.Append("- A task with several parts: call `plan` first with numbered steps, then work through them and update it as they finish. A failed command comes back with its reason — fix that one command and carry on; do not start over, and do not repeat a command that just succeeded.\n");
        sb.Append("- Before removing a table, a sheet, a whole section or more than a few paragraphs, say so in one line, then do it. Never delete or rewrite content the user did not ask about; `replace … --preview` first when a find and replace may touch many places, and say the count.\n");
        sb.Append("- Never tell the user to run commands. When the user asks a question, answer it from the document without editing.\n");
        sb.Append("- Reply in the user's language. While working, at most one short sentence between steps; finish with one line saying what changed and where (段落、表格、单元格 by their content or position, not paths), or the answer. No command output, no paths unless the user asks.\n");
        if (web) sb.Append("- Use web_search and web_fetch when the task needs current or outside facts the document doesn't have, and say which URLs you used.\n");
        sb.Append('\n').Append(Mcp.Instructions).Append("\n\n");
        var format = file is null ? null : Adapters.CanonicalFormat(Path.GetExtension(file));
        if (file is not null)
        {
            var relative = Path.IsPathRooted(file) && file.StartsWith(workspace + Path.DirectorySeparatorChar, StringComparison.Ordinal)
                ? Path.GetRelativePath(workspace, file) : file;
            sb.Append("Open file: ").Append(relative).Append(" (use exactly this path in commands).\n");
            if (format is "pdf") sb.Append("PDF files are read-only: read them with `view <file> text`; to edit, `export --to <name>.docx` first and tell the user.\n");
            if (format is "docx") sb.Append('\n').Append(WordRules);
            if (format is "xlsx") sb.Append('\n').Append(ExcelRules);
            if (format is "md") sb.Append('\n').Append(MarkdownRules);
            if (format is "pptx") sb.Append('\n').Append(SlideDesign);
        }
        sb.Append("\n## Command reference\n\n").Append(Run(["help"]));
        if (format is not null and not "pdf")
            foreach (var kind in Registry.ForFormat(format))
                sb.Append('\n').Append(Run(["help", format, kind.Name]));
        if (file is not null && outline)
        {
            var text = format == "xlsx" ? "" : Run(["view", file, "outline"]);
            if (format != "xlsx" && text.Length <= OutlineBudget) sb.Append("\n## Current outline\n\n").Append(text);
            else sb.Append("\n## Structure\n\nThe document is shown in short; read the parts you need with section, search, get or view text.\n\n").Append(Trim(Run(["view", file, "structure"]), OutlineBudget));
        }
        return sb.ToString();
    }

    /// <summary>What a Word document asks for: edits that keep its formatting, structure through properties, Chinese typography.</summary>
    public const string WordRules =
        "## Word (docx)\n\n"
        + "- Paragraph text: `set f.docx /body/paragraph[3] --prop text=\"…\"` rewrites the words and keeps the bold, colour and size of the words that stay; `md=` adds inline formatting (**bold**, *italic*, [link](url)). Never remove and re-add a paragraph to change its text.\n"
        + "- Structure through properties: headings are `heading` nodes with level; list items are paragraphs with list=bullet|number and level; style=Quote, align=center. Keep the document's own heading levels, numbering and styles; do not hand-format runs one by one to imitate a style.\n"
        + "- New content: `add f.docx /body --type paragraph --prop md=\"…\" --after /body/paragraph[3]`; a whole section from markdown: `section f.docx /body/heading[2] --md \"## 标题\\n\\n正文…\\n\\n- 要点\"` (the heading line included; headings, lists and tables come out as Word blocks).\n"
        + "- Tables: cells are `/body/table[1]/row[2]/cell[3]` with text/md; `add f.docx /body/table[1] --type row --prop data='[\"…\",\"…\"]'`; `--prop data=` on the table rewrites every cell. Sums and averages in a Word table are numbers you compute from the cells and write.\n"
        + "- Chinese text: 全角标点（，。：；！？「」“”）, no space between CJK and punctuation, one space between CJK and Latin letters or digits, 半角 digits and units; match the document's quote style and number style. English text: English typography.\n"
        + "- Tracked changes: when the document has track=true, text you write is recorded as revisions; do not accept or reject revisions unless asked.\n\n"
        + "Example — 把第二段写得委婉些，再在最后加一段总结:\n"
        + "  batch label=\"改写第二段并加总结\" commands=[\"set f.docx /body/paragraph[2] --prop text=\\\"…\\\"\", \"add f.docx /body --type heading --prop text=总结 --prop level=2\", \"add f.docx /body --type paragraph --prop md=\\\"…\\\"\"]\n"
        + "Example — 全文把「甲方」改成「委托方」: `replace f.docx --find 甲方 --with 委托方 --preview`, then without --preview; reply with the count.\n";

    /// <summary>What a workbook asks for: read the data first, formulas through the formula verb and its check, ranges for formatting.</summary>
    public const string ExcelRules =
        "## Excel (xlsx)\n\n"
        + "- Address cells as /sheet[1]/cell[B3], blocks as /sheet[1]/range[A2:D20], rows as /sheet[1]/row[3]; the structure below gives each sheet's used range, header row and first data row. Read the data with `get f.xlsx /sheet[1]/range[A1:D20]` before computing or describing anything.\n"
        + "- Formulas: `formula f.xlsx /sheet[1]/cell[E2] \"C2*D2\"` writes one; on a range it fills like Excel's fill handle, shifting relative references per row and column (`formula f.xlsx /sheet[1]/range[E2:E20] \"C2*D2\"`), $ pins a reference. The result lists what the formula refers to and warns about text in a numeric range, functions this editor cannot compute, unbalanced parentheses and circular references: when it warns, fix the formula and run it once more. Results compute when the sheet opens; never write a computed number where a formula belongs, and never overwrite a formula cell with a value.\n"
        + "- Summary row: fill the formulas across (`formula f.xlsx /sheet[1]/range[B21:E21] \"SUM(B2:B20)\"`), label it (`set f.xlsx /sheet[1]/cell[A21] --prop value=合计`), then format the row (`set f.xlsx /sheet[1]/range[A21:E21] --prop bold=true --prop border=thin`) — one batch.\n"
        + "- Formatting goes on ranges: bold, fill, color, size, format (a number format: 0.00, #,##0, 0%, yyyy-mm-dd), align, wrap, border=thin, borderColor; column widths, frozen panes and the filter on the sheet (`set f.xlsx /sheet[1] --prop freeze=A2 --prop filter=A1:E20 --prop widths='{\"A\":18}'`).\n"
        + "- Sort the data rows without the header: `set f.xlsx /sheet[1]/range[A2:E20] --prop sort=C:desc` (cells move with their formatting and formulas).\n"
        + "- Charts: `add f.xlsx /sheet[1] --type chart --prop type=column --prop title=\"…\" --prop categories=A2:A13 --prop series='[{\"name\":\"B1\",\"values\":\"B2:B13\"}]'`; line for a trend, pie for shares, bar for a ranking.\n"
        + "- Values: numbers as plain digits (1234.5 — units and thousands separators belong to the number format or the header), dates as yyyy-mm-dd, text as text; `--prop values=` on a range writes a block of cells at once. Never clear a range the user did not name.\n\n"
        + "Example — 在最后加一行合计并加粗:\n"
        + "  batch label=\"添加合计行\" commands=[\"set f.xlsx /sheet[1]/cell[A21] --prop value=合计\", \"formula f.xlsx /sheet[1]/range[B21:E21] \\\"SUM(B2:B20)\\\"\", \"set f.xlsx /sheet[1]/range[A21:E21] --prop bold=true\"]\n";

    /// <summary>What a markdown file asks for: sections rewritten as markdown, inline formatting kept, the file's conventions respected.</summary>
    public const string MarkdownRules =
        "## Markdown (md)\n\n"
        + "- Blocks are /body/heading[n], /body/paragraph[n] (list items are paragraphs with list= and level=), /body/table[n], /body/code[n]; paths count per kind from the top.\n"
        + "- A section is a heading and everything up to the next heading of the same or a higher level: `section f.md /body/heading[3]` reads it, `section f.md /body/heading[3] --md \"### 标题\\n\\n…\"` rewrites it exactly as written (heading line included), `--remove` deletes it. This is the way to rewrite, expand or restructure a part.\n"
        + "- Inline formatting: `set f.md /body/paragraph[2] --prop md=\"**要点**：…\"`; text= drops the formatting. Code blocks: text= and lang=.\n"
        + "- Tables: `/body/table[1]/row[2]/cell[1]` with text/md; `add f.md /body/table[1] --type row --prop data='[\"…\",\"…\"]'`; `--prop data='[[…],[…]]'` rewrites the whole table (the first row is the header).\n"
        + "- Lists: `add f.md /body --type paragraph --prop md=\"…\" --prop list=bullet --prop level=0 --after /body/paragraph[4]` continues a list; numbering renders itself.\n"
        + "- Keep the file's conventions: heading levels, bullet marker, blank lines, front matter, raw HTML and quote blocks stay as they are.\n\n"
        + "Example — 把「安装」一节改成三步: `section f.md /body/heading[2] --md \"## 安装\\n\\n1. …\\n2. …\\n3. …\"`\n";

    /// <summary>The design rules the assistant works by on a slide deck, and the properties that carry them out (the slide editor's
    /// ✦ 美化 asks for exactly this; a user's own "make it nicer" gets the same treatment).</summary>
    public const string SlideDesign =
        "## Slide design\n\n"
        + "When asked to design, polish or tidy slides, work by these rules and say in one line what you applied:\n"
        + "- One idea per slide: a title that states the point (a sentence, not a topic), at most five bullets of at most ten words; a crowded slide is split (a new slide of the same layout takes the rest) rather than shrunk below 18pt.\n"
        + "- Type scale: title 36–44pt, section title 40–48pt, body 20–24pt, captions and sources 12–14pt; at most two text sizes besides the title on a slide; the same size for the same role on every slide.\n"
        + "- Contrast: dark text on a light background or light on dark, never mid-grey on grey; one accent colour, on at most two elements per slide.\n"
        + "- Alignment and margins: nothing closer to an edge than a sixteenth of the slide's width; text left-aligned; neighbouring shapes share an edge or a centre line; equal gaps between siblings.\n"
        + "- Spacing: body line spacing 1.1–1.3, clear room between title and body, no shape touching another.\n"
        + "- Layout choice: title for the opening and the close, section for a chapter break, content for a list, two or comparison for two things side by side, picture or caption for one image with a short text, quote for a quotation, titleOnly for a chart or one big number, blank for a full-bleed picture.\n"
        + "- Palette: one palette for the whole deck (paper or mist for reports, ink or night for a keynote, sea for technology, clay or sand for warm subjects, rose for consumer topics) and one fonts=heading,body; never colour shapes one by one to fake a theme.\n\n"
        + "The tools: the document's palette= and fonts=; a slide's layout= (moves its placeholders as PowerPoint does), align=edge:paths, distribute=axis:paths and background=; a shape's fit=shrink when it reads overflow=true (or shorten its text, or enlarge the box), size=, and x= y= w= h=. "
        + "Read `view <file> outline` first — it shows every shape's box, size and overflow — then make the fewest edits that satisfy the rules, one command per change, and never rewrite the user's words unless asked.\n";

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
