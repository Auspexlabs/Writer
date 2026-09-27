using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Writer.Core;

namespace Writer.Cli;

/// <summary>用 ChatGPT 登录's wire format: the Codex backend's Responses API, POST {base}/responses, always streamed, with the
/// sign-in's access token, its account id and the Codex CLI's own headers (openai/codex, codex-rs/core, checked 2026-09-27).
/// Nothing is kept on OpenAI's side (store: false), so every request carries the whole turn: the reply's items, the model's
/// reasoning among them (encrypted), go back without their ids.</summary>
public sealed partial class Chat
{
    /// <summary>Set for a ChatGPT sign-in: the tokens for the next request (AiStore.ChatGptAccess); given the access token the backend
    /// refused, fresh ones.</summary>
    public Func<string?, CancellationToken, Task<ChatGptTokens>>? Account { get; init; }

    bool Codex => Account is not null;

    /// <summary>The Codex CLI sends its conversation's id with each request; the backend caches the prompt under it.</summary>
    readonly string _session = Guid.NewGuid().ToString();

    /// <summary>The access token last sent: cut out of any error text shown.</summary>
    string _access = "";

    /// <summary>The backend caps instructions (about 32 KiB): a longer system prompt goes in as the first, developer, message.</summary>
    const int InstructionsBytes = 30000;

    async Task<Reply> CodexTurn(string system, JsonArray messages, bool web, Func<string, Task>? onDelta, CancellationToken ct)
    {
        var body = CodexBody(Model, system, CodexInput(messages));
        body["tools"] = CodexTools(web);
        body["tool_choice"] = "auto";
        body["parallel_tool_calls"] = false;
        body["include"] = new JsonArray((JsonNode)"reasoning.encrypted_content");
        using var response = await CodexSend(() => Post("/responses", body), HttpCompletionOption.ResponseHeadersRead, ct);
        var (items, streamed) = await CodexEvents(response, onDelta, null, ct);

        var blocks = new List<(string?, string?, string?, JsonObject?)>();
        var echo = new JsonArray();
        var said = false;
        foreach (var item in items)
            switch (Str(item["type"]))
            {
                case "message":
                    var text = MessageText(item);
                    said = true;
                    if (text.Length == 0) break;
                    blocks.Add((text, null, null, null));
                    echo.Add(Message("assistant", "output_text", text));
                    break;
                case "function_call":
                    var id = Str(item["call_id"]) is { Length: > 0 } c ? c : "call_" + blocks.Count;
                    var name = Str(item["name"]) is { Length: > 0 } n ? n : Mcp.ToolName;
                    var args = Str(item["arguments"]) ?? "";
                    JsonObject? input = null;
                    try
                    {
                        input = JsonNode.Parse(args.Length == 0 ? "{}" : args) as JsonObject;
                    }
                    catch (JsonException)
                    {
                        // the tool result says what was wrong with the arguments
                    }
                    blocks.Add((null, id, name, input));
                    echo.Add((JsonNode)new JsonObject { ["type"] = "function_call", ["call_id"] = id, ["name"] = name, ["arguments"] = input is null ? "{}" : args });
                    break;
                case "reasoning" when Str(item["encrypted_content"]) is { Length: > 0 } thought:
                    // the model's reasoning, sealed: given back within the turn it carries on from it
                    echo.Add((JsonNode)new JsonObject { ["type"] = "reasoning", ["summary"] = item["summary"]?.DeepClone() ?? new JsonArray(), ["encrypted_content"] = thought });
                    break;
            }
        if (!said && streamed.Length > 0) // the stream ended before its message item was whole: what arrived
        {
            blocks.Insert(0, (streamed.ToString(), null, null, null));
            echo.Insert(0, Message("assistant", "output_text", streamed.ToString()));
        }
        var tools = blocks.Any(b => b.Item2 is not null);
        return new Reply(blocks, new JsonObject { ["role"] = "assistant", ["codex"] = echo }, tools);
    }

    /// <summary>A completion: one short reply without tools or reasoning, read until its first line is whole (the backend takes no
    /// output limit). A model that refuses to skip reasoning is asked again with a little.</summary>
    async Task<string> CodexComplete(string model, string user, CancellationToken ct)
    {
        for (var effort = "none"; ; effort = "low")
        {
            var body = CodexBody(model, Assistant.CompleteSystem, new JsonArray(Message("user", "input_text", user)));
            body["reasoning"] = new JsonObject { ["effort"] = effort };
            try
            {
                using var response = await CodexSend(() => Post("/responses", body), HttpCompletionOption.ResponseHeadersRead, ct);
                var (items, streamed) = await CodexEvents(response, null, t => t.Length > 300 || t.ToString().Count(ch => ch == '\n') >= 2, ct);
                return streamed.Length > 0 ? streamed.ToString() : string.Concat(items.Where(i => Str(i["type"]) == "message").Select(MessageText));
            }
            catch (WriterException ex) when (effort == "none" && (Says(ex, "effort") || Says(ex, "reasoning")))
            {
                // an older model: its lowest effort instead
            }
        }
    }

    /// <summary>GET {base}/models: what this ChatGPT plan may use, as the Codex CLI lists it ({models: [{slug, display_name, visibility}]}).</summary>
    async Task<JsonObject> CodexModels(CancellationToken ct)
    {
        using var response = await CodexSend(() => new HttpRequestMessage(HttpMethod.Get, _baseUrl + "/models?client_version=" + Uri.EscapeDataString(ChatGpt.ClientVersion)),
            HttpCompletionOption.ResponseContentRead, ct);
        var text = await response.Content.ReadAsStringAsync(ct);
        try
        {
            var reply = JsonNode.Parse(text) as JsonObject ?? throw new JsonException();
            // models the Codex CLI keeps out of its picker
            if (reply["models"] is JsonArray all)
                foreach (var hidden in all.OfType<JsonObject>().Where(m => Str(m["visibility"]) is "hide" or "none").ToList()) all.Remove(hidden);
            return reply;
        }
        catch (JsonException)
        {
            throw NotAnApi(text);
        }
    }

    JsonObject CodexBody(string model, string system, JsonArray input)
    {
        if (Encoding.UTF8.GetByteCount(system) > InstructionsBytes)
        {
            input.Insert(0, Message("developer", "input_text", system));
            system = "You are the assistant inside Writer, a document editor. Follow the developer message.";
        }
        return new JsonObject
        {
            ["model"] = model,
            ["instructions"] = system,
            ["input"] = input,
            ["store"] = false, // the backend refuses a request that leaves it out
            ["stream"] = true,
            ["prompt_cache_key"] = _session,
        };
    }

    /// <summary>The loop's transcript as Responses API input items: text turns as messages, a reply's items as they came (the
    /// "codex" list its Reply message carries), tool results as function_call_output.</summary>
    static JsonArray CodexInput(JsonArray messages)
    {
        var input = new JsonArray();
        foreach (var m in messages.OfType<JsonObject>())
            switch (Str(m["role"]))
            {
                case "user":
                    input.Add(Message("user", "input_text", Str(m["content"]) ?? ""));
                    break;
                case "assistant" when m["codex"] is JsonArray items:
                    foreach (var item in items) input.Add(item!.DeepClone());
                    break;
                case "assistant":
                    if (Str(m["content"]) is { Length: > 0 } text) input.Add(Message("assistant", "output_text", text));
                    break;
                case "tool":
                    input.Add((JsonNode)new JsonObject { ["type"] = "function_call_output", ["call_id"] = Str(m["tool_call_id"]) ?? "", ["output"] = Str(m["content"]) ?? "" });
                    break;
            }
        return input;
    }

    static JsonNode Message(string role, string type, string text) => new JsonObject
    {
        ["type"] = "message",
        ["role"] = role,
        ["content"] = new JsonArray((JsonNode)new JsonObject { ["type"] = type, ["text"] = text }),
    };

    static string MessageText(JsonObject item) =>
        string.Concat((item["content"] as JsonArray ?? []).OfType<JsonObject>().Where(c => Str(c["type"]) == "output_text").Select(c => Str(c["text"]) ?? ""));

    /// <summary>The tools as Responses API functions: writer, batch and plan always, web_search and web_fetch when 联网搜索 is on.</summary>
    static JsonArray CodexTools(bool web)
    {
        static JsonNode Fn(string name, string description, JsonObject schema) => new JsonObject
        {
            ["type"] = "function", ["name"] = name, ["description"] = description, ["strict"] = false, ["parameters"] = schema,
        };
        var tools = new JsonArray();
        foreach (var (name, description, schema) in EditingTools()) tools.Add(Fn(name, description, schema));
        if (web)
        {
            tools.Add(Fn(Web.SearchToolName, Web.SearchDescription, Web.SearchParameters()));
            tools.Add(Fn(Web.FetchToolName, Web.FetchDescription, Web.FetchParameters()));
        }
        return tools;
    }

    HttpRequestMessage Post(string path, JsonObject body) => new(HttpMethod.Post, _baseUrl + path) { Content = new StringContent(Json(body), Encoding.UTF8, "application/json") };

    /// <summary>Sends a request with the sign-in's tokens and the Codex CLI's headers. A 401 gets fresh tokens and one more try; other
    /// refusals become WriterExceptions. The caller disposes the response.</summary>
    async Task<HttpResponseMessage> CodexSend(Func<HttpRequestMessage> build, HttpCompletionOption option, CancellationToken ct)
    {
        string? refused = null;
        while (true)
        {
            var tokens = await Account!(refused, ct);
            _access = tokens.Access;
            using var request = build();
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", tokens.Access);
            if (tokens.AccountId.Length > 0) request.Headers.TryAddWithoutValidation("chatgpt-account-id", tokens.AccountId);
            request.Headers.TryAddWithoutValidation("originator", ChatGpt.Originator);
            request.Headers.TryAddWithoutValidation("version", ChatGpt.ClientVersion);
            request.Headers.TryAddWithoutValidation("session_id", _session);
            request.Headers.TryAddWithoutValidation("User-Agent", ChatGpt.UserAgent);
            request.Headers.TryAddWithoutValidation("OpenAI-Beta", "responses=experimental");
            if (request.Method == HttpMethod.Post) request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
            HttpResponseMessage response;
            try
            {
                response = await _http.SendAsync(request, option, ct);
            }
            catch (Exception ex) when (ex is HttpRequestException or IOException)
            {
                throw Unreachable(ex);
            }
            catch (TaskCanceledException) when (!ct.IsCancellationRequested)
            {
                throw TimedOut();
            }
            if (response.IsSuccessStatusCode) return response;
            using (response)
            {
                if (response.StatusCode == HttpStatusCode.Unauthorized && refused is null)
                {
                    refused = tokens.Access; // lapsed early, or revoked: refresh once and try again
                    continue;
                }
                throw CodexFailure((int)response.StatusCode, await response.Content.ReadAsStringAsync(ct), response.Content.Headers.ContentType?.MediaType);
            }
        }
    }

    /// <summary>A refusal from the Codex backend in the user's words: a lapsed sign-in, a used-up plan, Cloudflare's challenge page.</summary>
    WriterException CodexFailure(int status, string body, string? mediaType = null)
    {
        if (status == 401) return new(ErrorCode.Io, "ChatGPT 登录已失效，请重新登录", Trim(Redact(ErrorText(body)), 300));
        if (mediaType == "text/html" || body.TrimStart().StartsWith('<'))
            return new(ErrorCode.Io, $"ChatGPT 的服务拦下了这次请求（{status}）", "多半是网络环境被 Cloudflare 拦截：换个网络，或稍后再试。");
        JsonObject? error = null;
        try
        {
            var o = JsonNode.Parse(body) as JsonObject;
            error = o?["error"] as JsonObject ?? (o?["response"] as JsonObject)?["error"] as JsonObject ?? (Str(o?["type"]) is not null ? o : null);
        }
        catch (JsonException)
        {
            // not JSON: the generic message below
        }
        var type = Str(error?["type"]) ?? Str(error?["code"]) ?? "";
        if (status == 429 || type is "usage_limit_reached" or "usage_not_included" or "rate_limit_exceeded")
        {
            var wait = error?["resets_in_seconds"] is JsonValue v && v.TryGetValue<long>(out var s) ? s
                : error?["resets_at"] is JsonValue a && a.TryGetValue<long>(out var at) ? at - DateTimeOffset.UtcNow.ToUnixTimeSeconds() : -1;
            var hint = wait > 0 ? $"大约 {Math.Max(1, (wait + 59) / 60)} 分钟后恢复。" : Trim(Redact(Str(error?["message"]) ?? ""), 200);
            return type == "usage_not_included"
                ? new(ErrorCode.Io, "这个 ChatGPT 套餐不含 Codex，用不了这个模型", hint)
                : new(ErrorCode.Io, "ChatGPT 套餐的用量已达上限", hint);
        }
        // the parameter a refusal names goes along: a completion retries on "reasoning.effort"
        var said = error is null ? ErrorText(body) : (Str(error["message"]) ?? ErrorText(body)) + (Str(error["param"]) is { Length: > 0 } param ? $" ({param})" : "");
        return Failure(status, said);
    }

    /// <summary>The backend's event stream: text deltas to <paramref name="onDelta"/> as they arrive; the reply's items in order once each is
    /// done, and the text streamed. <paramref name="enough"/> may end the reading early (a completion needs one line).</summary>
    async Task<(List<JsonObject> Items, StringBuilder Text)> CodexEvents(HttpResponseMessage response, Func<string, Task>? onDelta, Func<StringBuilder, bool>? enough, CancellationToken ct)
    {
        var items = new List<JsonObject>();
        var text = new StringBuilder();
        string? other = null;
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
                if (!line.StartsWith("data:", StringComparison.Ordinal))
                {
                    if (line.Length > 0 && !line.StartsWith("event:", StringComparison.Ordinal)) other ??= line;
                    continue;
                }
                var data = line[5..].Trim();
                if (data == "[DONE]") break;
                JsonObject? e;
                try
                {
                    e = JsonNode.Parse(data) as JsonObject;
                }
                catch (JsonException)
                {
                    continue;
                }
                if (e is null) continue;
                streamed = true;
                switch (Str(e["type"]))
                {
                    case "response.output_text.delta":
                        if (Str(e["delta"]) is not { Length: > 0 } piece) break;
                        text.Append(piece);
                        if (onDelta is not null) await onDelta(piece);
                        if (enough?.Invoke(text) == true) return (items, text);
                        break;
                    case "response.output_item.done":
                        if (e["item"] is JsonObject item) items.Add(item);
                        break;
                    case "response.completed":
                        return (items, text);
                    case "response.incomplete":
                        var why = Str(((e["response"] as JsonObject)?["incomplete_details"] as JsonObject)?["reason"]);
                        if (why == "content_filter") throw new WriterException(ErrorCode.Io, "模型服务拦截了这次回复（内容审核）", "");
                        return (items, text); // cut short (the output limit): what came
                    case "response.failed":
                    case "error":
                        throw CodexFailure(0, data);
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
        if (!streamed) throw NotAnApi(other ?? "");
        if (items.Count == 0 && text.Length == 0) throw new WriterException(ErrorCode.Io, "模型服务的回复中断了，请重试", "");
        return (items, text);
    }
}
