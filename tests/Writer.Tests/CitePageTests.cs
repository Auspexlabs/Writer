using System.Net;
using System.Net.Http.Headers;
using System.Text;
using Writer.Cli;

namespace Writer.Tests;

/// <summary>GET /cite/page: a web page's html, for looking a source up from its address (ui/cite.js reads the tags a page carries for
/// citation managers) — the page is on another site, which a page of the app may not read itself. The web is a stand-in.</summary>
public class CitePageTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("writer-citepage").FullName;
    Serve? _server;

    public void Dispose()
    {
        _server?.Dispose();
        Directory.Delete(_dir, true);
    }

    sealed class Web : HttpMessageHandler
    {
        public List<Uri> Asked { get; } = [];

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            Asked.Add(request.RequestUri!);
            return Task.FromResult(request.RequestUri!.AbsolutePath switch
            {
                "/article" => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("<html><head><meta name=\"citation_title\" content=\"Behavior of technetium\"></head></html>", Encoding.UTF8, "text/html") },
                "/paper.pdf" => new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent([0x25, 0x50, 0x44, 0x46]) { Headers = { ContentType = new MediaTypeHeaderValue("application/pdf") } } },
                _ => new HttpResponseMessage(HttpStatusCode.NotFound) { Content = new StringContent("gone", Encoding.UTF8, "text/html") },
            });
        }
    }

    HttpClient Start(Web web)
    {
        _server = new Serve(0, requireToken: true, workspace: _dir, ai: new AiStore(Path.Combine(_dir, "ai.json"), _ => null, web));
        _server.Start();
        var client = new HttpClient { BaseAddress = new Uri(_server.Url) };
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _server.Token);
        return client;
    }

    [Fact]
    public async Task A_page_is_read_for_the_app_and_only_a_web_page_over_http()
    {
        var web = new Web();
        var client = Start(web);
        var page = await client.GetAsync("/cite/page?url=" + Uri.EscapeDataString("https://link.springer.com/article"));
        Assert.Equal(HttpStatusCode.OK, page.StatusCode);
        Assert.Contains("citation_title", await page.Content.ReadAsStringAsync());
        Assert.Equal("https://link.springer.com/article", web.Asked.Single().ToString());

        Assert.Equal(HttpStatusCode.BadGateway, (await client.GetAsync("/cite/page?url=" + Uri.EscapeDataString("https://example.org/paper.pdf"))).StatusCode); // not a page
        Assert.Equal(HttpStatusCode.BadGateway, (await client.GetAsync("/cite/page?url=" + Uri.EscapeDataString("https://example.org/missing"))).StatusCode);
        Assert.Equal(HttpStatusCode.BadRequest, (await client.GetAsync("/cite/page?url=" + Uri.EscapeDataString("file:///etc/passwd"))).StatusCode);
        Assert.Equal(3, web.Asked.Count); // the file address never left the engine

        using var stranger = new HttpClient { BaseAddress = new Uri(_server!.Url) };
        Assert.Equal(HttpStatusCode.Unauthorized, (await stranger.GetAsync("/cite/page?url=" + Uri.EscapeDataString("https://link.springer.com/article"))).StatusCode);
    }
}
