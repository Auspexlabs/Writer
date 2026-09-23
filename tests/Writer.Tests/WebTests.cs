using System.Net;
using Writer.Cli;

namespace Writer.Tests;

/// <summary>The assistant's no-key internet tools (Web.cs): the Bing result parser, the /ck/a redirect decoder, HTML-to-text,
/// the private-address refusal and the seen-URL rule. Pure string/logic pieces only — no live network calls.</summary>
public class WebTests
{
    // A trimmed, representative sample of a Bing SERP: one b_ad (must be skipped), one organic result behind a /ck/a
    // redirect with formatting tags inside the title and snippet, and one organic result with a direct link and a nested
    // sitelinks <ul><li> inside its caption (checks the block extractor doesn't stop at the sitelinks' own </li>).
    const string BingSample = """
        <!doctype html>
        <html><head><title>cats - Bing</title></head>
        <body>
        <ol id="b_results">
        <li class="b_ad"><div class="sb_add"><h2><a href="https://ads.example.com/click?id=1">Buy Cat Food Online</a></h2></div></li>
        <li class="b_algo">
        <h2><a target="_blank" href="https://www.bing.com/ck/a?!&amp;&amp;p=abc123&u=a1aHR0cHM6Ly9lbi53aWtpcGVkaWEub3JnL3dpa2kvQ2F0&amp;ntb=1" h="ID=SERP,1.1">Cat - <strong>Wikipedia</strong></a></h2>
        <div class="b_caption">
        <div class="b_attribution"><cite>https://en.wikipedia.org/wiki/Cat</cite></div>
        <p>The cat (<b>Felis catus</b>), commonly referred to as the domestic cat, is a small carnivorous mammal.</p>
        </div>
        </li>
        <li class="b_algo">
        <h2><a href="https://www.example.com/cats/breeds?sort=popular">Popular Cat Breeds &amp; Care Tips</a></h2>
        <div class="b_caption">
        <ul class="b_vList">
        <li><a href="https://www.example.com/cats/breeds/persian">Persian</a></li>
        <li><a href="https://www.example.com/cats/breeds/siamese">Siamese</a></li>
        </ul>
        <p>Explore popular cat breeds, their temperament and care needs in one guide.</p>
        </div>
        </li>
        </ol>
        </body></html>
        """;

    [Fact]
    public void ParseBingResults_reads_title_link_and_snippet_and_skips_ads_and_nested_sitelinks()
    {
        var results = Web.ParseBingResults(BingSample);

        Assert.Equal(2, results.Count); // the b_ad block is not a b_algo and must not appear
        Assert.Equal("Cat - Wikipedia", results[0].Title); // <strong> stripped
        Assert.Equal("https://en.wikipedia.org/wiki/Cat", results[0].Url); // decoded from the /ck/a redirect
        Assert.Equal("The cat ( Felis catus ), commonly referred to as the domestic cat, is a small carnivorous mammal.", results[0].Snippet);

        Assert.Equal("Popular Cat Breeds & Care Tips", results[1].Title); // &amp; decoded
        Assert.Equal("https://www.example.com/cats/breeds?sort=popular", results[1].Url); // a direct link, not a redirect
        Assert.Equal("Explore popular cat breeds, their temperament and care needs in one guide.", results[1].Snippet); // the nested <li> sitelinks didn't end the block early
    }

    [Fact]
    public async Task SearchAsync_formats_the_top_results_as_numbered_title_url_snippet_and_collects_their_urls_as_seen()
    {
        // ParseBingResults is exercised above; this checks SearchAsync's own formatting and empty-query guard without any network call.
        Assert.Equal((1, "The search query is empty.", (IReadOnlyList<string>)[]), await Web.SearchAsync("   ", CancellationToken.None));
    }

    [Theory]
    [InlineData("https://www.bing.com/ck/a?p=x&u=a1aHR0cHM6Ly9lbi53aWtpcGVkaWEub3JnL3dpa2kvQ2F0&ntb=1", "https://en.wikipedia.org/wiki/Cat")]
    [InlineData("https://www.bing.com/ck/a?p=x&amp;u=a1aHR0cHM6Ly93d3cuZXhhbXBsZS5jb20vY2F0cy9icmVlZHM_c29ydD1wb3B1bGFy&amp;ntb=1", "https://www.example.com/cats/breeds?sort=popular")]
    [InlineData("https://www.example.com/direct-link", "https://www.example.com/direct-link")] // not a /ck/a redirect: unchanged
    [InlineData("https://www.bing.com/ck/a?p=x&u=not-base64url&ntb=1", "https://www.bing.com/ck/a?p=x&u=not-base64url&ntb=1")] // doesn't start with a1: unchanged
    [InlineData("https://www.bing.com/ck/a?p=x&u=a1***not-valid-base64***&ntb=1", "https://www.bing.com/ck/a?p=x&u=a1***not-valid-base64***&ntb=1")] // a1 prefix but garbage: falls back rather than throwing
    public void DecodeCkLink_recovers_the_real_url_from_a_ck_a_redirect_and_passes_through_anything_else(string href, string expected) =>
        Assert.Equal(expected, Web.DecodeCkLink(href));

    [Fact]
    public void HtmlToText_drops_script_style_nav_header_footer_strips_tags_decodes_entities_and_collapses_whitespace()
    {
        const string html = """
            <html><head><title>Cats &amp; Kittens</title><style>body{color:red}</style></head>
            <body>
            <header>Site header</header>
            <nav><a href="/">Home</a></nav>
            <script>track('pageview');</script>
            <main><h1>About cats</h1><p>Cats   are   great &amp; independent
            pets.</p></main>
            <footer>Copyright 2026</footer>
            </body></html>
            """;

        var (title, text) = Web.HtmlToText(html);

        Assert.Equal("Cats & Kittens", title);
        Assert.Equal("About cats Cats are great & independent pets.", text); // the head (title included) must not also show up in the body text
    }

    [Fact]
    public void HtmlToText_truncates_long_pages_to_about_20000_characters()
    {
        var html = "<p>" + new string('x', 25_000) + "</p>";
        var (_, text) = Web.HtmlToText(html);
        Assert.True(text.Length < 20_100, $"expected roughly 20k chars, got {text.Length}");
        Assert.EndsWith("…(truncated)", text);
    }

    [Theory]
    [InlineData("127.0.0.1", true)]      // loopback v4
    [InlineData("::1", true)]            // loopback v6
    [InlineData("::ffff:127.0.0.1", true)] // IPv4-mapped loopback: must be unwrapped, not waved through
    [InlineData("10.1.2.3", true)]       // RFC1918
    [InlineData("172.16.0.5", true)]     // RFC1918
    [InlineData("172.31.255.255", true)] // RFC1918 (top of the /12)
    [InlineData("172.32.0.1", false)]    // just outside the 172.16.0.0/12 private block
    [InlineData("192.168.1.1", true)]    // RFC1918
    [InlineData("169.254.1.1", true)]    // link-local v4
    [InlineData("0.0.0.0", true)]        // this-network / often routes to localhost
    [InlineData("fe80::1", true)]        // link-local v6
    [InlineData("fc00::1", true)]        // unique-local v6
    [InlineData("fd12:3456::1", true)]   // unique-local v6
    [InlineData("8.8.8.8", false)]       // public
    [InlineData("2001:4860:4860::8888", false)] // public v6 (Google DNS)
    public void IsBlockedAddress_refuses_loopback_private_link_local_and_unique_local(string address, bool blocked) =>
        Assert.Equal(blocked, Web.IsBlockedAddress(IPAddress.Parse(address)));

    [Fact]
    public async Task FetchAsync_refuses_a_url_that_has_not_appeared_in_the_conversation()
    {
        var (code, output, urls) = await Web.FetchAsync("https://example.com/not-seen", new HashSet<string>(), CancellationToken.None);
        Assert.Equal(1, code);
        Assert.Contains("hasn't appeared", output);
        Assert.Empty(urls);
    }

    [Fact]
    public async Task FetchAsync_allows_a_url_seen_earlier_in_the_conversation_but_refuses_a_private_address()
    {
        const string url = "http://127.0.0.1:9/probe"; // seen, but resolves to loopback: must be refused before any request
        var (code, output, _) = await Web.FetchAsync(url, new HashSet<string> { url }, CancellationToken.None);
        Assert.Equal(1, code);
        Assert.Contains("local or private", output);
    }

    /// <summary>Answers every request with a redirect to <paramref name="location"/> and records what was asked.</summary>
    sealed class Redirects(string location) : HttpMessageHandler
    {
        public List<Uri> Asked { get; } = [];

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            Asked.Add(request.RequestUri!);
            var response = new HttpResponseMessage(HttpStatusCode.Found);
            response.Headers.Location = new Uri(location);
            return Task.FromResult(response);
        }
    }

    [Fact]
    public async Task Get_checks_every_redirect_hop_so_a_public_page_cannot_bounce_to_a_local_address()
    {
        // an IP literal needs no DNS lookup, so this runs offline; the page answers with a redirect to loopback
        var server = new Redirects("http://127.0.0.1:9/admin");
        using var client = new HttpClient(server);

        var refused = await Assert.ThrowsAnyAsync<Exception>(() => Web.Get(new Uri("http://93.184.216.34/start"), CancellationToken.None, client));

        Assert.Contains("local or private", refused.Message);
        Assert.Equal([new Uri("http://93.184.216.34/start")], server.Asked); // the loopback hop was never requested
    }

    [Fact]
    public async Task FetchAsync_refuses_non_http_schemes_before_checking_whether_the_url_was_seen()
    {
        var (code, output, _) = await Web.FetchAsync("ftp://example.com/file", new HashSet<string> { "ftp://example.com/file" }, CancellationToken.None);
        Assert.Equal(1, code);
        Assert.Contains("http and https", output);
    }

    [Fact]
    public void ExtractUrls_finds_http_and_https_urls_and_trims_trailing_punctuation()
    {
        var text = "See https://example.com/a, and also (https://example.com/b) or https://example.com/c. Not ftp://example.com/d.";
        Assert.Equal(["https://example.com/a", "https://example.com/b", "https://example.com/c"], Web.ExtractUrls(text));
    }

    [Fact]
    public void ExtractUrls_keeps_a_closing_paren_that_is_part_of_the_url_itself()
    {
        // unlike the markdown/prose "(https://example.com/b)" case above, this URL's own trailing ')' balances a '(' inside it
        var text = "See https://en.wikipedia.org/wiki/Cat_(disambiguation) for other uses.";
        Assert.Equal(["https://en.wikipedia.org/wiki/Cat_(disambiguation)"], Web.ExtractUrls(text));
    }

    [Fact]
    public void NormalizeUrl_accepts_only_absolute_http_and_https_urls()
    {
        Assert.Equal("https://example.com/a", Web.NormalizeUrl("https://example.com/a"));
        Assert.Null(Web.NormalizeUrl("ftp://example.com/a"));
        Assert.Null(Web.NormalizeUrl("not a url"));
        Assert.Null(Web.NormalizeUrl("/relative/path"));
    }
}
