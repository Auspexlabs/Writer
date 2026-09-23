using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace Writer.Cli;

/// <summary>The assistant's two no-key internet tools, wired into Chat.cs's tool loop alongside writer: web_search (Bing's
/// results page, no API) and web_fetch (any http/https page already named in the conversation). Neither ever throws for an
/// ordinary failure — trouble comes back as (1, message), the same shape Serve.RunArgv uses for the writer tool.</summary>
public static partial class Web
{
    public const string SearchToolName = "web_search";
    public const string FetchToolName = "web_fetch";
    public const string QueryParam = "query";
    public const string UrlParam = "url";

    public const string SearchDescription =
        "Search the web (Bing, no API key) for current or outside facts the document and your training don't have. "
        + "Returns up to 8 results: title, URL and a short snippet. Follow up with web_fetch on a promising URL to read it in full.";

    public const string FetchDescription =
        "Fetch a web page by URL and return its title and text (http/https only). The URL must already appear earlier in "
        + "this conversation — in the user's message, the document, or a web_search/web_fetch result — not one you construct yourself.";

    public static JsonObject SearchParameters() => new()
    {
        ["type"] = "object",
        ["properties"] = new JsonObject { [QueryParam] = new JsonObject { ["type"] = "string", ["description"] = "The search query." } },
        ["required"] = new JsonArray(QueryParam),
    };

    public static JsonObject FetchParameters() => new()
    {
        ["type"] = "object",
        ["properties"] = new JsonObject { [UrlParam] = new JsonObject { ["type"] = "string", ["description"] = "The http(s) URL to fetch; must have appeared earlier in the conversation." } },
        ["required"] = new JsonArray(UrlParam),
    };

    // Redirects are followed by hand (Get below) so every hop gets the address check, and a direct connection can only reach
    // an address that passes it: ConnectCallback resolves the host itself, so a DNS answer that changes between the check and
    // the connect (DNS rebinding) cannot slip a local address through. A connection to the user's own proxy is left alone:
    // the proxy resolves the host, and the per-hop check has already run.
    static readonly HttpClient Http = new(new SocketsHttpHandler
    {
        AllowAutoRedirect = false,
        PooledConnectionLifetime = TimeSpan.FromMinutes(2),
        ConnectCallback = async (context, ct) =>
        {
            var host = context.DnsEndPoint.Host.Trim('[', ']');
            var direct = string.Equals(host, context.InitialRequestMessage.RequestUri?.IdnHost.Trim('[', ']'), StringComparison.OrdinalIgnoreCase);
            var addresses = await Dns.GetHostAddressesAsync(host, ct);
            if (direct) addresses = addresses.Where(a => !IsBlockedAddress(a)).ToArray();
            if (addresses.Length == 0) throw new HttpRequestException(LocalRefusal);
            var socket = new Socket(SocketType.Stream, ProtocolType.Tcp) { NoDelay = true };
            try
            {
                await socket.ConnectAsync(addresses, context.DnsEndPoint.Port, ct);
                return new NetworkStream(socket, ownsSocket: true);
            }
            catch
            {
                socket.Dispose();
                throw;
            }
        },
    }) { Timeout = TimeSpan.FromSeconds(30) };
    const string LocalRefusal = "Refused: that address is local or private and cannot be fetched.";
    const int MaxRedirects = 5;

    /// <summary>A refusal the model should read as-is (a local address, a bad redirect), not as a network failure.</summary>
    sealed class Refused(string message) : Exception(message);

    /// <summary>GET with redirects followed by hand, at most <see cref="MaxRedirects"/>: every hop must be http(s) and its host
    /// must resolve to public addresses only. The caller disposes the response. <paramref name="client"/> is for tests.</summary>
    public static async Task<HttpResponseMessage> Get(Uri uri, CancellationToken ct, HttpClient? client = null)
    {
        for (var hop = 0; ; hop++)
        {
            IPAddress[] addresses;
            try
            {
                addresses = await Dns.GetHostAddressesAsync(uri.IdnHost.Trim('[', ']'), ct);
            }
            catch (SocketException ex)
            {
                throw new Refused($"Could not resolve {uri.Host}: {ex.Message}");
            }
            if (addresses.Length == 0 || addresses.Any(IsBlockedAddress)) throw new Refused(LocalRefusal);
            var request = new HttpRequestMessage(HttpMethod.Get, uri);
            request.Headers.UserAgent.ParseAdd(UserAgent);
            request.Headers.AcceptLanguage.ParseAdd(AcceptLanguage);
            var response = await (client ?? Http).SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct);
            if ((int)response.StatusCode is not (301 or 302 or 303 or 307 or 308) || response.Headers.Location is not { } location) return response;
            response.Dispose();
            if (hop == MaxRedirects) throw new Refused("Too many redirects.");
            uri = location.IsAbsoluteUri ? location : new Uri(uri, location);
            if (uri.Scheme is not ("http" or "https")) throw new Refused("Refused: the page redirects to a non-web address.");
        }
    }
    const int RequestTimeoutSeconds = 15;
    const int MaxFetchBytes = 2 * 1024 * 1024;
    const int MaxChars = 20_000;
    // A normal desktop Chrome UA and an Accept-Language that favours Chinese, since this app is Chinese-first; Bing serves
    // English queries fine either way, and mainland China's redirect to cn.bing.com follows automatically (Http allows it).
    const string UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36";
    const string AcceptLanguage = "zh-CN,zh;q=0.9,en;q=0.8";

    public sealed record BingResult(string Title, string Url, string Snippet);

    /// <summary>Runs a Bing search and formats the top ~8 organic results as numbered title/URL/snippet lines. Urls is every
    /// result link (already decoded), which Chat adds to the turn's seen-URL set so a follow-up web_fetch can use one.</summary>
    public static async Task<(int Code, string Output, IReadOnlyList<string> Urls)> SearchAsync(string query, CancellationToken ct)
    {
        query = query.Trim();
        if (query.Length == 0) return (1, "The search query is empty.", []);
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(RequestTimeoutSeconds));
        string html;
        try
        {
            using var response = await Get(new Uri("https://www.bing.com/search?q=" + Uri.EscapeDataString(query)), timeout.Token);
            if (!response.IsSuccessStatusCode) return (1, $"Bing returned {(int)response.StatusCode} {response.ReasonPhrase}.", []);
            html = Decode(await ReadCapped(response.Content, MaxFetchBytes, timeout.Token), response.Content.Headers.ContentType?.CharSet);
        }
        catch (Exception ex) when (ex is HttpRequestException or IOException or Refused)
        {
            return (1, $"Web search failed: {ex.Message}", []);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            return (1, "Web search timed out.", []);
        }
        var results = ParseBingResults(html).Take(8).ToList();
        if (results.Count == 0) return (0, $"No results for \"{query}\".", []);
        var sb = new StringBuilder();
        for (var i = 0; i < results.Count; i++)
        {
            if (i > 0) sb.Append('\n');
            sb.Append(i + 1).Append(". ").Append(results[i].Title).Append('\n').Append(results[i].Url).Append('\n').Append(results[i].Snippet).Append('\n');
        }
        return (0, sb.ToString().TrimEnd(), results.Select(r => r.Url).ToList());
    }

    /// <summary>Fetches a page already named in the conversation (<paramref name="seenUrls"/>) and returns its title and text
    /// (HTML) or its body as-is (text/plain, JSON); anything else is refused. Refuses loopback, private, link-local and
    /// unique-local destinations after resolving DNS. Urls is every link found in the returned text, which Chat folds into
    /// the seen-URL set so the model can follow a link the page itself printed.</summary>
    public static async Task<(int Code, string Output, IReadOnlyList<string> Urls)> FetchAsync(string url, IReadOnlySet<string> seenUrls, CancellationToken ct)
    {
        if (NormalizeUrl(url) is not { } normalized) return (1, "Only http and https URLs are supported.", []);
        if (!seenUrls.Contains(normalized))
            return (1, "That URL hasn't appeared in this conversation (not in the user's message, the document, or an earlier web_search/web_fetch result). Ask the user for the link, or find it with web_search first.", []);
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(RequestTimeoutSeconds));
        HttpResponseMessage response;
        try
        {
            response = await Get(new Uri(normalized), timeout.Token);
        }
        catch (Refused ex)
        {
            return (1, ex.Message, []);
        }
        catch (Exception ex) when (ex is HttpRequestException or IOException)
        {
            return (1, $"Could not fetch the page: {ex.Message}", []);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            return (1, "Fetching the page timed out.", []);
        }
        using (response)
        {
            if (!response.IsSuccessStatusCode) return (1, $"The server returned {(int)response.StatusCode} {response.ReasonPhrase}.", []);
            var mediaType = response.Content.Headers.ContentType?.MediaType?.ToLowerInvariant() ?? "";
            byte[] bytes;
            try
            {
                bytes = await ReadCapped(response.Content, MaxFetchBytes, timeout.Token);
            }
            catch (OperationCanceledException) when (!ct.IsCancellationRequested)
            {
                return (1, "Fetching the page timed out.", []);
            }
            var text = Decode(bytes, response.Content.Headers.ContentType?.CharSet);
            string output;
            if (mediaType is "text/html" or "application/xhtml+xml")
            {
                var (title, body) = HtmlToText(text);
                output = title.Length > 0 ? title + "\n\n" + body : body;
            }
            else if (mediaType is "text/plain" || mediaType == "application/json" || mediaType.EndsWith("+json", StringComparison.Ordinal))
            {
                output = Truncate(text.Trim(), MaxChars);
            }
            else
            {
                return (1, $"Cannot show this content type ({(mediaType.Length > 0 ? mediaType : "unknown")}); only HTML, text and JSON pages are supported.", []);
            }
            return (0, output, ExtractUrls(output).ToList());
        }
    }

    // ---------- Bing result parsing (pure; no network) ----------

    /// <summary>Every organic result (li.b_algo): the title and link from its h2, the snippet from its caption paragraph.
    /// ponytail: regex-scrapes a specific, known-as-of-2026 Bing markup shape. Bing can change this without notice; a result
    /// that doesn't match just drops out rather than erroring. Upgrade path: a real HTML parser if scraping breaks often.</summary>
    public static List<BingResult> ParseBingResults(string html)
    {
        var results = new List<BingResult>();
        foreach (var block in ExtractElements(html, "li", "b_algo"))
        {
            var h2 = H2Pattern().Match(block);
            if (!h2.Success) continue;
            var href = HrefPattern().Match(h2.Value);
            if (!href.Success) continue;
            var title = StripTags(h2.Groups[1].Value);
            if (title.Length == 0) continue;
            var captionAt = block.IndexOf("b_caption", StringComparison.OrdinalIgnoreCase);
            var caption = captionAt >= 0 ? block[captionAt..] : block;
            var p = ParagraphPattern().Match(caption);
            results.Add(new BingResult(title, DecodeCkLink(href.Groups[1].Value), p.Success ? StripTags(p.Groups[1].Value) : ""));
        }
        return results;
    }

    /// <summary>Bing wraps organic result links in a /ck/a redirect whose u= parameter is "a1" followed by the base64url of
    /// the real URL; anything else (a direct link, an unrecognised redirect) is returned unchanged.</summary>
    public static string DecodeCkLink(string href)
    {
        var decoded = WebUtility.HtmlDecode(href).Trim();
        if (!Uri.TryCreate(decoded, UriKind.Absolute, out var uri) || !uri.AbsolutePath.Contains("/ck/a", StringComparison.OrdinalIgnoreCase))
            return decoded;
        foreach (var pair in uri.Query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var eq = pair.IndexOf('=');
            if (eq < 0 || pair[..eq] != "u") continue;
            var value = Uri.UnescapeDataString(pair[(eq + 1)..]);
            if (!value.StartsWith("a1", StringComparison.Ordinal)) continue;
            if (TryBase64Url(value[2..]) is { } real) return real;
        }
        return decoded;
    }

    static string? TryBase64Url(string s)
    {
        var padded = s.Replace('-', '+').Replace('_', '/');
        padded += (padded.Length % 4) switch { 2 => "==", 3 => "=", _ => "" };
        try
        {
            return padded.Length % 4 == 1 ? null : Encoding.UTF8.GetString(Convert.FromBase64String(padded));
        }
        catch (FormatException)
        {
            return null;
        }
    }

    /// <summary>Every top-level occurrence of &lt;tag class="...name..."&gt;, tracking nested same-name tags so an inner list
    /// doesn't end the block early, with its inner HTML (not the wrapping tag itself).</summary>
    static IEnumerable<string> ExtractElements(string html, string tag, string className)
    {
        var open = new Regex($@"<{tag}\b[^>]*\bclass\s*=\s*""[^""]*\b{Regex.Escape(className)}\b[^""]*""[^>]*>", RegexOptions.IgnoreCase);
        var any = new Regex($@"<{tag}\b[^>]*>|</{tag}\s*>", RegexOptions.IgnoreCase);
        foreach (Match start in open.Matches(html))
        {
            var depth = 1;
            var pos = start.Index + start.Length;
            Match t;
            while (depth > 0 && (t = any.Match(html, pos)).Success)
            {
                depth += t.Value.StartsWith("</", StringComparison.Ordinal) ? -1 : 1;
                pos = t.Index + t.Length;
                if (depth == 0) { yield return html[(start.Index + start.Length)..t.Index]; break; }
            }
        }
    }

    // ---------- HTML to text (pure; no network) ----------

    /// <summary>Drops head/script/style/nav/header/footer, strips every remaining tag, decodes entities and collapses
    /// whitespace; the title comes from &lt;title&gt; separately (dropping the whole head keeps it out of the body text
    /// too). Cuts the text to ~20k characters.</summary>
    public static (string Title, string Text) HtmlToText(string html)
    {
        var titleMatch = TitleTagPattern().Match(html);
        var title = titleMatch.Success ? StripTags(titleMatch.Groups[1].Value) : "";
        // ponytail: a non-greedy match, so a nested element of the same tag name (a <nav> inside a <nav>) leaves its outer
        // close tag as stray text instead of being dropped too. Rare in real pages; a real HTML parser would close this.
        var body = DroppedElementPattern().Replace(html, " ");
        return (title, Truncate(StripTags(body), MaxChars));
    }

    /// <summary>Removes tags (as a space, so adjoining cells/words don't run together), decodes entities, collapses runs of
    /// whitespace into one space.</summary>
    public static string StripTags(string html) => CollapseWhitespace(WebUtility.HtmlDecode(TagPattern().Replace(html, " ")));

    static string CollapseWhitespace(string s) => WhitespacePattern().Replace(s, " ").Trim();

    static string Truncate(string s, int max) => s.Length <= max ? s : s[..max] + "…(truncated)";

    // ---------- address safety (pure; no network) ----------

    /// <summary>Loopback, RFC1918/link-local/this-network IPv4, and link-local/unique-local/site-local IPv6 — the ranges
    /// web_fetch must never reach. Unwraps an IPv4-mapped IPv6 address first, a well-known way to smuggle one past a naive
    /// v4-only check.</summary>
    public static bool IsBlockedAddress(IPAddress address)
    {
        var ip = address.IsIPv4MappedToIPv6 ? address.MapToIPv4() : address;
        if (IPAddress.IsLoopback(ip)) return true;
        if (ip.AddressFamily == AddressFamily.InterNetwork)
        {
            var b = ip.GetAddressBytes();
            return b[0] == 0 || b[0] == 10 || (b[0] == 172 && b[1] is >= 16 and <= 31) || (b[0] == 192 && b[1] == 168) || (b[0] == 169 && b[1] == 254);
        }
        if (ip.AddressFamily == AddressFamily.InterNetworkV6)
            return ip.Equals(IPAddress.IPv6Any) || ip.IsIPv6LinkLocal || ip.IsIPv6SiteLocal || ip.IsIPv6UniqueLocal;
        return true; // an address family we don't recognise: refuse rather than guess
    }

    // ---------- URLs seen so far this turn (pure; no network) ----------

    /// <summary>Every http/https URL in <paramref name="text"/>, normalised with <see cref="NormalizeUrl"/> (trailing
    /// sentence punctuation trimmed first).</summary>
    public static IEnumerable<string> ExtractUrls(string? text)
    {
        if (string.IsNullOrEmpty(text)) yield break;
        foreach (Match m in UrlPattern().Matches(text))
            if (NormalizeUrl(TrimTrailingProse(m.Value)) is { } url)
                yield return url;
    }

    /// <summary>Sentence punctuation right after a URL is never part of it; a trailing ')' is only part of the surrounding
    /// prose (or markdown link syntax) when it has no matching '(' inside the URL — otherwise it's a real Wikipedia-style
    /// "(disambiguation)" URL and must stay. Getting this wrong would make a seen URL fail to match its clean form later.</summary>
    static string TrimTrailingProse(string url)
    {
        url = url.TrimEnd('.', ',', ';', ':', '!', '?', '"', '\'');
        return url.EndsWith(')') && url.Count(c => c == '(') < url.Count(c => c == ')') ? url[..^1] : url;
    }

    /// <summary>An absolute http/https URL's canonical form (for the seen-URL set and for comparing against it), or null.</summary>
    public static string? NormalizeUrl(string url) =>
        Uri.TryCreate(url.Trim(), UriKind.Absolute, out var uri) && uri.Scheme is "http" or "https" ? uri.AbsoluteUri : null;

    // ---------- shared plumbing ----------

    static async Task<byte[]> ReadCapped(HttpContent content, int maxBytes, CancellationToken ct)
    {
        await using var stream = await content.ReadAsStreamAsync(ct);
        using var buffer = new MemoryStream();
        var chunk = new byte[81920];
        int read;
        while (buffer.Length < maxBytes && (read = await stream.ReadAsync(chunk, ct)) > 0)
            buffer.Write(chunk, 0, Math.Min(read, maxBytes - (int)buffer.Length));
        return buffer.ToArray();
    }

    static string Decode(byte[] bytes, string? charset)
    {
        if (!string.IsNullOrWhiteSpace(charset))
        {
            try
            {
                return Encoding.GetEncoding(charset.Trim('"')).GetString(bytes);
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
            {
                // an unrecognised or unavailable (InvariantGlobalization) charset: fall back to UTF-8 below
            }
        }
        return Encoding.UTF8.GetString(bytes);
    }

    [GeneratedRegex(@"<h2\b[^>]*>(.*?)</h2>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex H2Pattern();

    [GeneratedRegex(@"href\s*=\s*""([^""]*)""", RegexOptions.IgnoreCase)]
    private static partial Regex HrefPattern();

    [GeneratedRegex(@"<p\b[^>]*>(.*?)</p>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex ParagraphPattern();

    [GeneratedRegex(@"<title\b[^>]*>(.*?)</title>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex TitleTagPattern();

    [GeneratedRegex(@"<(head|script|style|nav|header|footer)\b[^>]*>.*?</\1\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex DroppedElementPattern();

    [GeneratedRegex(@"<[^>]*>")]
    private static partial Regex TagPattern();

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespacePattern();

    [GeneratedRegex(@"https?://[^\s<>""']+", RegexOptions.IgnoreCase)]
    private static partial Regex UrlPattern();
}
