using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Writer.Core;

namespace Writer.Cli;

/// <summary>The OpenAI-compatible wire format: POST {base}/chat/completions with the writer tool as a function and stream: true.
/// Text and reasoning arrive as deltas; tool calls arrive as fragments (id and name first, arguments in pieces) keyed by index.</summary>
public sealed partial class Chat
{
    /// <summary>How long a stream may stay silent (reasoning models can think for a while before the first token).</summary>
    static readonly TimeSpan Idle = TimeSpan.FromSeconds(180);

    // ponytail: set once a provider refused earlier turns that carry no reasoning_content (DeepSeek in thinking mode wants every
    // earlier turn's reasoning back when tools are on). The UI keeps only text, so from then on earlier turns travel as quoted text
    // inside the new user message. Keeping each message's reasoning in the UI transcript would lift this.
    bool _fold;

    // ponytail: set once a provider refused function tools together with reasoning (OpenAI's gpt-5.4 and later take tools on
    // /chat/completions only with reasoning_effort "none", and older models refuse the parameter itself), so it is sent only
    // after such a refusal. Their /responses API would keep the reasoning; that is a third wire format.
    bool _noReasoning;

    sealed class Call
    {
        public string Id = "", Name = "";
        public readonly StringBuilder Args = new();
    }

    /// <summary>One request, retried at most once per refusal that names what to change: reasoning_effort, reasoning_content.</summary>
    async Task<Reply> OpenAiTurn(string system, JsonArray messages, bool web, Func<string, Task>? onDelta, CancellationToken ct)
    {
        if (_fold) Fold(messages);
        while (true)
        {
            try
            {
                return await OpenAiRequest(system, messages, web, onDelta, ct);
            }
            catch (WriterException ex) when (!_noReasoning && Says(ex, "reasoning_effort"))
            {
                _noReasoning = true;
            }
            catch (WriterException ex) when (!_fold && Says(ex, "reasoning_content"))
            {
                _fold = true;
                if (!Fold(messages)) throw;
            }
        }
    }

    /// <summary>The provider's words are in the message or, when they were given a short Chinese name, in the hint.</summary>
    static bool Says(WriterException ex, string word) => ex.Message.Contains(word, StringComparison.Ordinal) || ex.Hint.Contains(word, StringComparison.Ordinal);

    /// <summary>One streamed request; each content delta goes to <paramref name="onDelta"/> as it arrives.</summary>
    async Task<Reply> OpenAiRequest(string system, JsonArray messages, bool web, Func<string, Task>? onDelta, CancellationToken ct)
    {
        var all = new JsonArray(new JsonObject { ["role"] = "system", ["content"] = system });
        foreach (var m in messages) all.Add(m!.DeepClone());
        var body = new JsonObject
        {
            ["model"] = Model,
            ["messages"] = all,
            ["tools"] = OpenAiTools(web),
            ["stream"] = true,
        };
        if (_noReasoning) body["reasoning_effort"] = "none";
        using var request = new HttpRequestMessage(HttpMethod.Post, _baseUrl + "/chat/completions") { Content = new StringContent(Json(body), Encoding.UTF8, "application/json") };
        if (_apiKey.Length > 0) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
        using var response = await Send(request, HttpCompletionOption.ResponseHeadersRead, ct);

        var text = new StringBuilder();
        StringBuilder? reasoning = null;
        var calls = new List<Call>();
        var byIndex = new Dictionary<int, Call>();
        string? finish = null, other = null;
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
                    other ??= line.Length > 0 ? line : null;
                    continue;
                }
                var data = line[5..].Trim();
                if (data == "[DONE]") break;
                JsonNode? node;
                try
                {
                    node = JsonNode.Parse(data);
                }
                catch (JsonException)
                {
                    continue;
                }
                if (node is not JsonObject chunk) continue;
                streamed = true;
                if (chunk["error"] is not null) throw Failure(0, ErrorText(data));
                if (chunk["choices"] is not JsonArray { Count: > 0 } choices || choices[0] is not JsonObject choice) continue;
                var delta = choice["delta"] as JsonObject;
                if (Str(delta?["content"]) is { Length: > 0 } piece)
                {
                    text.Append(piece);
                    if (onDelta is not null) await onDelta(piece);
                }
                if (Str(delta?["reasoning_content"]) is { } thought) (reasoning ??= new()).Append(thought);
                if (delta?["tool_calls"] is JsonArray parts)
                    foreach (var part in parts.OfType<JsonObject>()) Accumulate(part, calls, byIndex);
                if (Str(choice["finish_reason"]) is { } reason) finish = reason;
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
        if (finish is "content_filter" or "sensitive") throw new WriterException(ErrorCode.Io, "模型服务拦截了这次回复（内容审核）", "");

        var blocks = new List<(string?, string?, string?, JsonObject?)>();
        if (text.Length > 0) blocks.Add((text.ToString(), null, null, null));
        var toolCalls = new JsonArray();
        for (var i = 0; i < calls.Count; i++)
        {
            var call = calls[i];
            if (call.Id.Length == 0) call.Id = "call_" + i;
            var args = call.Args.ToString();
            JsonObject? input = null;
            try
            {
                input = JsonNode.Parse(args.Length == 0 ? "{}" : args) as JsonObject;
            }
            catch (JsonException)
            {
                args = "{}"; // echoing broken JSON back could get the next request refused; the tool result says what went wrong
            }
            var name = call.Name.Length > 0 ? call.Name : Mcp.ToolName;
            blocks.Add((null, call.Id, name, input));
            toolCalls.Add((JsonNode)new JsonObject
            {
                ["id"] = call.Id,
                ["type"] = "function",
                ["function"] = new JsonObject { ["name"] = name, ["arguments"] = args },
            });
        }
        var message = new JsonObject { ["role"] = "assistant", ["content"] = text.Length > 0 ? text.ToString() : null };
        if (reasoning is not null) message["reasoning_content"] = reasoning.ToString(); // thinking models want it back within the turn
        if (toolCalls.Count > 0) message["tool_calls"] = toolCalls;
        return new Reply(blocks, message, calls.Count > 0);
    }

    /// <summary>The tools offered on /chat/completions: writer, batch and plan always, web_search and web_fetch when 联网搜索 is on.</summary>
    static JsonArray OpenAiTools(bool web)
    {
        // JsonNode, not JsonObject: JsonArray.Add<T> for a JsonObject is the reflection overload, which NativeAOT refuses
        static JsonNode Fn(string name, string description, JsonObject schema) => new JsonObject
        {
            ["type"] = "function",
            ["function"] = new JsonObject { ["name"] = name, ["description"] = description, ["parameters"] = schema },
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

    /// <summary>Adds one tool_calls fragment. The first fragment of a call brings its id and name, later ones (same index) add
    /// argument text. A fragment with a new id starts a new call even under a used index, since some local servers number every call 0.</summary>
    static void Accumulate(JsonObject part, List<Call> calls, Dictionary<int, Call> byIndex)
    {
        var id = Str(part["id"]) is { Length: > 0 } s ? s : null;
        int? index = part["index"] is JsonValue v && v.TryGetValue<int>(out var n) ? n : null;
        Call? call = null;
        if (index is { } i) byIndex.TryGetValue(i, out call);
        else if (id is null) call = calls.LastOrDefault();
        if (call is null || (id is not null && call.Id.Length > 0 && id != call.Id))
        {
            calls.Add(call = new Call());
            if (index is { } j) byIndex[j] = call;
        }
        if (id is not null) call.Id = id;
        var function = part["function"] as JsonObject;
        if (Str(function?["name"]) is { Length: > 0 } name) call.Name = name;
        if (Str(function?["arguments"]) is { } args) call.Args.Append(args);
    }

    /// <summary>Earlier turns become quoted text at the start of the newest user message, so no earlier assistant message is sent.
    /// False when there was nothing earlier to fold.</summary>
    static bool Fold(JsonArray messages)
    {
        var turn = -1;
        for (var i = messages.Count - 1; i >= 0 && turn < 0; i--)
            if (Str(messages[i]?["role"]) == "user") turn = i;
        if (turn <= 0) return false;
        var sb = new StringBuilder("Earlier in this conversation:\n<conversation>\n");
        for (var i = 0; i < turn; i++)
            sb.Append(Str(messages[i]?["role"])).Append(": ").Append(Str(messages[i]?["content"])).Append("\n\n");
        sb.Append("</conversation>\n\n").Append(Str(messages[turn]?["content"]));
        var rest = messages.Skip(turn + 1).Select(m => m!.DeepClone()).ToList();
        messages.Clear();
        messages.Add((JsonNode)new JsonObject { ["role"] = "user", ["content"] = sb.ToString() });
        foreach (var m in rest) messages.Add(m);
        return true;
    }

    static string? Str(JsonNode? node) => node is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;
}
