using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Writer.Core;

namespace Writer.Cli;

/// <summary>The autocomplete and the model list: two small requests beside the assistant's turns. A completion is one short reply
/// without tools and without streaming, from the fast model the settings name (CompleteModel), with the model's thinking turned off
/// where the provider has a switch for it, since a completion that comes after the user typed on is no use.</summary>
public sealed partial class Chat
{
    static readonly TimeSpan CompleteTimeout = TimeSpan.FromSeconds(15);
    const int CompleteTokens = 400; // room for a model that thinks anyway (GLM-5.3, Kimi K3); the prompt asks for a few words

    /// <summary>POST /complete: what the user is likely to type next at the caret, one line, or "" (Assistant.CleanCompletion).</summary>
    public async Task<string> CompleteAsync(string before, string after, CancellationToken ct) => await CompleteAsync(before, after, null, ct);

    /// <param name="hint">Where the text is typed, as the editor describes it (Assistant.CompleteMessage).</param>
    public async Task<string> CompleteAsync(string before, string after, string? hint, CancellationToken ct)
    {
        if (before.Trim().Length == 0) return "";
        using var limit = CancellationTokenSource.CreateLinkedTokenSource(ct);
        limit.CancelAfter(CompleteTimeout);
        var model = CompleteModel.Length > 0 ? CompleteModel : Model;
        var user = Assistant.CompleteMessage(before, after, hint);
        try
        {
            var text = Codex ? await CodexComplete(model, user, limit.Token)
                : _openAi ? await OpenAiComplete(model, user, limit.Token) : await AnthropicComplete(model, user, limit.Token);
            return Assistant.CleanCompletion(text, before);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            throw TimedOut();
        }
    }

    async Task<string> AnthropicComplete(string model, string user, CancellationToken ct)
    {
        var body = new JsonObject
        {
            ["model"] = model,
            ["max_tokens"] = CompleteTokens,
            ["system"] = Assistant.CompleteSystem,
            ["messages"] = new JsonArray((JsonNode)new JsonObject { ["role"] = "user", ["content"] = user }),
        };
        using var request = new HttpRequestMessage(HttpMethod.Post, _baseUrl + "/v1/messages") { Content = new StringContent(Json(body), Encoding.UTF8, "application/json") };
        request.Headers.Add("x-api-key", _apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        var reply = await ReadJson(request, ct);
        // text blocks only: a model that thinks puts its thinking in blocks of its own
        return string.Concat((reply["content"] as JsonArray ?? []).OfType<JsonObject>().Where(b => Str(b["type"]) == "text").Select(b => Str(b["text"]) ?? ""));
    }

    /// <summary>With the provider's switch for thinking first, then, when the provider refuses that parameter, without it.</summary>
    async Task<string> OpenAiComplete(string model, string user, CancellationToken ct)
    {
        for (var quiet = true; ; quiet = false)
        {
            var body = new JsonObject
            {
                ["model"] = model,
                ["messages"] = new JsonArray(
                    (JsonNode)new JsonObject { ["role"] = "system", ["content"] = Assistant.CompleteSystem },
                    (JsonNode)new JsonObject { ["role"] = "user", ["content"] = user }),
                ["stream"] = false,
            };
            var extra = quiet && NoThinking(body);
            if (!extra) body["max_tokens"] = CompleteTokens;
            using var request = new HttpRequestMessage(HttpMethod.Post, _baseUrl + "/chat/completions") { Content = new StringContent(Json(body), Encoding.UTF8, "application/json") };
            if (_apiKey.Length > 0) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            try
            {
                var reply = await ReadJson(request, ct);
                var choice = (reply["choices"] as JsonArray)?.FirstOrDefault() as JsonObject;
                return Str((choice?["message"] as JsonObject)?["content"]) ?? "";
            }
            catch (WriterException ex) when (extra && Refused(ex))
            {
                // the provider does not take the switch (an older model, a compatible server): the same request without it
            }
        }
    }

    /// <summary>The provider's way to answer without thinking first (checked 2026-09-26 in each provider's API docs), with the
    /// output limit under the name that provider takes; false when it has none (the caller sets max_tokens).</summary>
    bool NoThinking(JsonObject body)
    {
        switch (_provider)
        {
            case "openai":
                body["max_completion_tokens"] = CompleteTokens;
                body["reasoning_effort"] = "none";
                return true;
            case "qwen":
                body["max_tokens"] = CompleteTokens;
                body["enable_thinking"] = false;
                return true;
            case "deepseek" or "glm" or "doubao" or "kimi":
                body["max_tokens"] = CompleteTokens;
                body["thinking"] = new JsonObject { ["type"] = "disabled" };
                return true;
            default:
                return false;
        }
    }

    /// <summary>A 400-range refusal that names one of the parameters NoThinking adds.</summary>
    static bool Refused(WriterException ex) =>
        !ex.Message.StartsWith("Key", StringComparison.Ordinal) && !ex.Message.StartsWith("地址", StringComparison.Ordinal)
        && (Says(ex, "thinking") || Says(ex, "reasoning_effort") || Says(ex, "max_completion_tokens") || Says(ex, "parameter") || Says(ex, "参数"));

    /// <summary>POST /ai/models: the models the provider's API lists for this key (GET {base}/v1/models on Anthropic, {base}/models
    /// elsewhere; the plan's models for a ChatGPT sign-in), as (id, display name). Embedding, speech, image and moderation models are
    /// left out: the assistant cannot use them.</summary>
    public async Task<List<(string Id, string Name)>> ModelsAsync(CancellationToken ct)
    {
        using var limit = CancellationTokenSource.CreateLinkedTokenSource(ct);
        limit.CancelAfter(TimeSpan.FromSeconds(20));
        using var request = new HttpRequestMessage(HttpMethod.Get, _baseUrl + (_openAi ? "/models" : "/v1/models?limit=100"));
        if (!_openAi)
        {
            request.Headers.Add("x-api-key", _apiKey);
            request.Headers.Add("anthropic-version", "2023-06-01");
        }
        else if (_apiKey.Length > 0) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
        JsonObject reply;
        try
        {
            reply = Codex ? await CodexModels(limit.Token) : await ReadJson(request, limit.Token);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            throw TimedOut();
        }
        var list = new List<(string, string)>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var m in ((reply["data"] ?? reply["models"]) as JsonArray ?? []).OfType<JsonObject>())
        {
            var id = Str(m["id"]) ?? Str(m["slug"]) ?? Str(m["name"]) ?? Str(m["model"]);
            if (string.IsNullOrWhiteSpace(id) || NotForChat.IsMatch(id) || !seen.Add(id)) continue;
            list.Add((id, Str(m["display_name"]) ?? ""));
        }
        return list;
    }

    static readonly System.Text.RegularExpressions.Regex NotForChat = new(
        @"embed|tts|whisper|dall-e|gpt-image|image-|-image|audio|realtime|transcribe|moderation|rerank|speech|sora|seedream|seedance|cogview|cogvideo|wanx",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    /// <summary>A request whose answer is one JSON object; an answer that is not JSON is not a model API.</summary>
    async Task<JsonObject> ReadJson(HttpRequestMessage request, CancellationToken ct)
    {
        using var response = await Send(request, HttpCompletionOption.ResponseContentRead, ct);
        var text = await response.Content.ReadAsStringAsync(ct);
        try
        {
            return JsonNode.Parse(text) as JsonObject ?? throw new JsonException();
        }
        catch (JsonException)
        {
            throw NotAnApi(text);
        }
    }
}
