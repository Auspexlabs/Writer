using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Writer.Cli;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

public class ServeTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("writer-serve").FullName;
    readonly Serve _server;
    readonly HttpClient _client = new(new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false });

    public ServeTests()
    {
        _server = new Serve(0, requireToken: true, ["http://localhost:5173"], _dir, Serve.FindUiDir());
        _server.Start();
        _client.BaseAddress = new Uri(_server.Url);
    }

    [Fact]
    public async Task Workspace_files_can_be_listed_uploaded_fetched_and_addressed_relatively()
    {
        Doc("report.docx", P("Hello"));
        Directory.CreateDirectory(Path.Combine(_dir, "sub"));
        File.WriteAllText(Path.Combine(_dir, "sub", "notes.md"), "# Notes\n");
        File.WriteAllText(Path.Combine(_dir, "ignore.xyz"), "x");

        var listed = JsonDocument.Parse(await (await _client.SendAsync(Request(HttpMethod.Get, "/files"))).Content.ReadAsStringAsync()).RootElement;
        Assert.Equal(_dir, listed.GetProperty("workspace").GetString());
        Assert.Equal(Runner.Version, listed.GetProperty("version").GetString());
        var files = listed.GetProperty("files").EnumerateArray().Select(f => f.GetProperty("path").GetString()).ToList();
        Assert.Contains("report.docx", files);
        Assert.Contains("sub/notes.md", files);
        Assert.DoesNotContain("ignore.xyz", files);

        var relative = await _client.SendAsync(Request(HttpMethod.Get, "/outline?file=report.docx"));
        Assert.Equal(HttpStatusCode.OK, relative.StatusCode);
        var run = await _client.SendAsync(Request(HttpMethod.Post, "/run", JsonSerializer.Serialize(new { command = "view report.docx text" })));
        Assert.Equal("Hello\n", JsonDocument.Parse(await run.Content.ReadAsStringAsync()).RootElement.GetProperty("output").GetString());

        var upload = Request(HttpMethod.Put, "/file?file=uploads/new.md");
        upload.Content = new ByteArrayContent(Encoding.UTF8.GetBytes("# Uploaded\n"));
        var uploaded = await _client.SendAsync(upload);
        Assert.Equal(HttpStatusCode.OK, uploaded.StatusCode);
        Assert.Equal("# Uploaded\n", File.ReadAllText(Path.Combine(_dir, "uploads", "new.md")));
        var escape = Request(HttpMethod.Put, "/file?file=../outside.md");
        escape.Content = new ByteArrayContent([1]);
        Assert.Equal(HttpStatusCode.BadRequest, (await _client.SendAsync(escape)).StatusCode);

        var moved = Request(HttpMethod.Put, "/file?file=uploads/renamed.md&from=uploads/new.md");
        moved.Content = new ByteArrayContent([]);
        Assert.Equal(HttpStatusCode.OK, (await _client.SendAsync(moved)).StatusCode);
        Assert.False(File.Exists(Path.Combine(_dir, "uploads", "new.md")));
        Assert.Equal("# Uploaded\n", File.ReadAllText(Path.Combine(_dir, "uploads", "renamed.md")));

        var raw = await _client.SendAsync(Request(HttpMethod.Get, "/file?file=sub/notes.md"));
        Assert.Equal("text/markdown", raw.Content.Headers.ContentType?.MediaType);
        Assert.Equal("attachment", raw.Content.Headers.ContentDisposition?.DispositionType);
        Assert.Equal("nosniff", raw.Headers.GetValues("X-Content-Type-Options").Single());
        Assert.Equal("# Notes\n", await raw.Content.ReadAsStringAsync());
        File.WriteAllText(Path.Combine(_dir, "说明 文件.md"), "# 中文\n");
        var unicode = await _client.SendAsync(Request(HttpMethod.Get, "/file?file=" + Uri.EscapeDataString("说明 文件.md")));
        Assert.Equal(HttpStatusCode.OK, unicode.StatusCode);
        Assert.Contains("filename*=UTF-8''%E8%AF%B4%E6%98%8E%20%E6%96%87%E4%BB%B6.md", unicode.Content.Headers.GetValues("Content-Disposition").Single());
        File.WriteAllText(Path.Combine(_dir, "page.html"), "<script>1</script>");
        var html = await _client.SendAsync(Request(HttpMethod.Get, "/file?file=page.html"));
        Assert.Equal("application/octet-stream", html.Content.Headers.ContentType?.MediaType);
        Assert.Equal(HttpStatusCode.BadRequest, (await _client.SendAsync(Request(HttpMethod.Get, "/file?file=" + Uri.EscapeDataString(Path.Combine(Path.GetTempPath(), "x.md"))))).StatusCode);
        if (!OperatingSystem.IsWindows())
        {
            File.CreateSymbolicLink(Path.Combine(_dir, "link.md"), Path.Combine(Path.GetTempPath(), "nowhere.md"));
            var linked = Request(HttpMethod.Put, "/file?file=link.md");
            linked.Content = new ByteArrayContent([1]);
            Assert.Equal(HttpStatusCode.BadRequest, (await _client.SendAsync(linked)).StatusCode);
        }

        var png = Path.Combine(_dir, "pic.png");
        File.WriteAllBytes(png, FakePng(2, 2));
        Assert.Equal(0, JsonDocument.Parse(await (await _client.SendAsync(Request(HttpMethod.Post, "/run",
            JsonSerializer.Serialize(new { argv = new[] { "add", "report.docx", "/body", "--type", "image", "--prop", "src=" + png } })))).Content.ReadAsStringAsync()).RootElement.GetProperty("code").GetInt32());
        var binary = await _client.SendAsync(Request(HttpMethod.Get, "/binary?file=report.docx&path=/body/image[1]"));
        Assert.Equal("image/png", binary.Content.Headers.ContentType?.MediaType);
        Assert.Equal(FakePng(2, 2), await binary.Content.ReadAsByteArrayAsync());
    }

    [Fact]
    public async Task Stat_reports_mtime_and_put_returns_the_mtime_it_just_wrote()
    {
        var file = Path.Combine(_dir, "a.md");
        File.WriteAllText(file, "one");
        var diskMs = new DateTimeOffset(File.GetLastWriteTimeUtc(file)).ToUnixTimeMilliseconds();

        var stat = JsonDocument.Parse(await (await _client.SendAsync(Request(HttpMethod.Get, "/stat?file=a.md"))).Content.ReadAsStringAsync()).RootElement;
        Assert.Equal("a.md", stat.GetProperty("path").GetString());
        Assert.Equal(3, stat.GetProperty("size").GetInt64());
        Assert.Equal(diskMs, stat.GetProperty("mtime").GetInt64());
        Assert.Equal(HttpStatusCode.NotFound, (await _client.SendAsync(Request(HttpMethod.Get, "/stat?file=missing.md"))).StatusCode);

        await Task.Delay(50);
        var put = Request(HttpMethod.Put, "/file?file=a.md");
        put.Content = new ByteArrayContent(Encoding.UTF8.GetBytes("two"));
        var written = JsonDocument.Parse(await (await _client.SendAsync(put)).Content.ReadAsStringAsync()).RootElement.GetProperty("mtime").GetInt64();
        var after = JsonDocument.Parse(await (await _client.SendAsync(Request(HttpMethod.Get, "/stat?file=a.md"))).Content.ReadAsStringAsync()).RootElement.GetProperty("mtime").GetInt64();
        Assert.Equal(after, written);
        Assert.NotEqual(diskMs, written);
    }

    [Fact]
    public async Task Put_with_a_stale_ifMtime_is_refused_with_409_and_writes_nothing()
    {
        var file = Path.Combine(_dir, "b.md");
        File.WriteAllText(file, "one");
        var mtime = JsonDocument.Parse(await (await _client.SendAsync(Request(HttpMethod.Get, "/stat?file=b.md"))).Content.ReadAsStringAsync()).RootElement.GetProperty("mtime").GetInt64();

        await Task.Delay(50);
        File.WriteAllText(file, "someone else's edit"); // another program (the MCP server, an agent's CLI call) writes it first

        var stale = Request(HttpMethod.Put, "/file?file=b.md&ifMtime=" + mtime);
        stale.Content = new ByteArrayContent(Encoding.UTF8.GetBytes("clobber"));
        var refused = await _client.SendAsync(stale);
        Assert.Equal(HttpStatusCode.Conflict, refused.StatusCode);
        Assert.Equal("CONFLICT", JsonDocument.Parse(await refused.Content.ReadAsStringAsync()).RootElement.GetProperty("error").GetProperty("code").GetString());
        Assert.Equal("someone else's edit", File.ReadAllText(file));

        var fresh = JsonDocument.Parse(await (await _client.SendAsync(Request(HttpMethod.Get, "/stat?file=b.md"))).Content.ReadAsStringAsync()).RootElement.GetProperty("mtime").GetInt64();
        var retry = Request(HttpMethod.Put, "/file?file=b.md&ifMtime=" + fresh);
        retry.Content = new ByteArrayContent(Encoding.UTF8.GetBytes("merged"));
        Assert.Equal(HttpStatusCode.OK, (await _client.SendAsync(retry)).StatusCode);
        Assert.Equal("merged", File.ReadAllText(file));
    }

    [Fact]
    public async Task Drafts_and_granted_files_are_the_only_places_outside_the_workspace()
    {
        var drafts = Directory.CreateTempSubdirectory("writer-drafts").FullName;
        var elsewhere = Directory.CreateTempSubdirectory("writer-elsewhere").FullName;
        using var server = new Serve(0, requireToken: true, null, _dir, Serve.FindUiDir(), drafts: drafts);
        server.Start();
        using var client = new HttpClient { BaseAddress = new Uri(server.Url) };
        HttpRequestMessage R(HttpMethod m, string path, byte[]? body = null)
        {
            var r = new HttpRequestMessage(m, path);
            r.Headers.Authorization = new AuthenticationHeaderValue("Bearer", server.Token);
            if (body is not null) r.Content = new ByteArrayContent(body);
            return r;
        }
        string U(string p) => Uri.EscapeDataString(p.Replace('\\', '/'));

        var listed = JsonDocument.Parse(await (await client.SendAsync(R(HttpMethod.Get, "/files"))).Content.ReadAsStringAsync()).RootElement;
        Assert.Equal(drafts.Replace('\\', '/'), listed.GetProperty("drafts").GetString());

        // a draft can be written, listed, read and deleted
        var draft = Path.Combine(drafts, "未命名.md");
        Assert.Equal(HttpStatusCode.OK, (await client.SendAsync(R(HttpMethod.Put, "/file?file=" + U(draft), "# d\n"u8.ToArray()))).StatusCode);
        var inDrafts = JsonDocument.Parse(await (await client.SendAsync(R(HttpMethod.Get, "/files?drafts=1"))).Content.ReadAsStringAsync()).RootElement;
        Assert.Equal(draft.Replace('\\', '/'), inDrafts.GetProperty("files")[0].GetProperty("path").GetString());
        Assert.Equal(HttpStatusCode.OK, (await client.SendAsync(R(HttpMethod.Get, "/file?file=" + U(draft)))).StatusCode);

        // elsewhere is refused until granted; a granted target receives the draft (move), then autosaves there
        var chosen = Path.Combine(elsewhere, "报告.md");
        Assert.Equal(HttpStatusCode.BadRequest, (await client.SendAsync(R(HttpMethod.Put, $"/file?file={U(chosen)}&from={U(draft)}", []))).StatusCode);
        Assert.True(server.ApplyGrantLine("allow " + chosen));
        Assert.False(server.ApplyGrantLine("grant " + chosen));
        Assert.False(server.ApplyGrantLine("allow relative/path.md"));
        Assert.Equal(HttpStatusCode.OK, (await client.SendAsync(R(HttpMethod.Put, $"/file?file={U(chosen)}&from={U(draft)}", []))).StatusCode);
        Assert.False(File.Exists(draft));
        Assert.Equal("# d\n", File.ReadAllText(chosen));
        Assert.Equal(HttpStatusCode.OK, (await client.SendAsync(R(HttpMethod.Put, "/file?file=" + U(chosen), "# e\n"u8.ToArray()))).StatusCode);
        Assert.Equal("# e\n", File.ReadAllText(chosen));

        // keep=1 copies (另存为); an existing target is replaced only when granted
        var copy = Path.Combine(elsewhere, "报告 副本.md");
        server.Grant(copy);
        Assert.Equal(HttpStatusCode.OK, (await client.SendAsync(R(HttpMethod.Put, $"/file?file={U(copy)}&from={U(chosen)}&keep=1", []))).StatusCode);
        Assert.True(File.Exists(chosen) && File.Exists(copy));
        File.WriteAllText(Path.Combine(elsewhere, "别人的.md"), "x");
        Assert.Equal(HttpStatusCode.BadRequest, (await client.SendAsync(R(HttpMethod.Put, $"/file?file={U(Path.Combine(elsewhere, "别人的.md"))}&from={U(copy)}", []))).StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await client.SendAsync(R(HttpMethod.Put, $"/file?file={U(chosen)}&from={U(copy)}&keep=1", []))).StatusCode);

        // a granted file may be renamed beside itself with the same extension, not elsewhere or to another type
        var renamed = Path.Combine(elsewhere, "总结.md");
        Assert.Equal(HttpStatusCode.OK, (await client.SendAsync(R(HttpMethod.Put, $"/file?file={U(renamed)}&from={U(chosen)}", []))).StatusCode);
        Assert.Equal(HttpStatusCode.OK, (await client.SendAsync(R(HttpMethod.Put, "/file?file=" + U(renamed), "# f\n"u8.ToArray()))).StatusCode);
        Assert.Equal(HttpStatusCode.BadRequest, (await client.SendAsync(R(HttpMethod.Put, $"/file?file={U(Path.Combine(elsewhere, "x.command"))}&from={U(renamed)}", []))).StatusCode);
        Assert.Equal(HttpStatusCode.BadRequest, (await client.SendAsync(R(HttpMethod.Put, $"/file?file={U(Path.Combine(drafts, "..", "y.md"))}&from={U(renamed)}", []))).StatusCode);

        // only drafts can be deleted
        Assert.Equal(HttpStatusCode.BadRequest, (await client.SendAsync(R(HttpMethod.Delete, "/file?file=" + U(renamed)))).StatusCode);
        File.WriteAllText(Path.Combine(drafts, "b.md"), "x");
        Assert.Equal(HttpStatusCode.OK, (await client.SendAsync(R(HttpMethod.Delete, "/file?file=" + U(Path.Combine(drafts, "b.md"))))).StatusCode);
        Assert.False(File.Exists(Path.Combine(drafts, "b.md")));
    }

    [Fact]
    public async Task A_granted_file_in_a_read_only_folder_is_replaced_in_place()
    {
        if (OperatingSystem.IsWindows()) return;
        var ro = Directory.CreateTempSubdirectory("writer-ro").FullName;
        var file = Path.Combine(ro, "a.md");
        File.WriteAllText(file, "old");
        File.SetUnixFileMode(ro, UnixFileMode.UserRead | UnixFileMode.UserExecute);
        try
        {
            _server.Grant(file);
            var put = Request(HttpMethod.Put, "/file?file=" + Uri.EscapeDataString(file));
            put.Content = new ByteArrayContent("new"u8.ToArray());
            Assert.Equal(HttpStatusCode.OK, (await _client.SendAsync(put)).StatusCode);
            Assert.Equal("new", File.ReadAllText(file));
        }
        finally { File.SetUnixFileMode(ro, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute); }
    }

    [Fact]
    public async Task The_grants_channel_applies_allow_lines_and_stops_the_server_at_the_end()
    {
        var file = Path.Combine(Path.GetTempPath(), "writer-granted.md");
        var stopped = false;
        Serve.ReadGrants(new StringReader("allow " + file + "\nnoise\n"), line => _server.ApplyGrantLine(line), () => stopped = true);
        Assert.True(stopped);
        File.WriteAllText(file, "x");
        Assert.Equal(HttpStatusCode.OK, (await _client.SendAsync(Request(HttpMethod.Get, "/file?file=" + Uri.EscapeDataString(file)))).StatusCode);
    }

    [Fact]
    public async Task Json_can_leave_out_run_nodes()
    {
        Doc("runs.docx", P("Hello"), P("World"));
        var full = await (await _client.SendAsync(Request(HttpMethod.Get, "/json?file=runs.docx"))).Content.ReadAsStringAsync();
        var lean = await (await _client.SendAsync(Request(HttpMethod.Get, "/json?file=runs.docx&skip=run"))).Content.ReadAsStringAsync();
        Assert.Contains("\"kind\": \"run\"", full);
        Assert.DoesNotContain("\"kind\": \"run\"", lean);
        var body = JsonDocument.Parse(lean).RootElement.GetProperty("children")[0];
        var first = body.GetProperty("children")[0];
        Assert.Equal("Hello", first.GetProperty("props").GetProperty("text").GetString());
        Assert.False(first.TryGetProperty("children", out _));
    }

    [Fact]
    public async Task Bundled_ui_is_served_under_app_without_a_token()
    {
        Assert.NotNull(_server.UiDir);
        var page = await _client.GetAsync("/app/index.dc.html");
        Assert.Equal(HttpStatusCode.OK, page.StatusCode);
        Assert.Equal("text/html", page.Content.Headers.ContentType?.MediaType);
        Assert.Contains("<x-dc>", await page.Content.ReadAsStringAsync());
        var script = await _client.GetAsync("/app/support.js");
        Assert.Equal("text/javascript", script.Content.Headers.ContentType?.MediaType);
        Assert.Equal(HttpStatusCode.OK, (await _client.GetAsync("/app/")).StatusCode);
        Assert.Equal(HttpStatusCode.Forbidden, (await _client.GetAsync("/app/..%2FWriter.slnx")).StatusCode);
        Assert.NotEqual(HttpStatusCode.OK, (await _client.GetAsync("/app/../Writer.slnx")).StatusCode);
        Assert.Equal(HttpStatusCode.NotFound, (await _client.GetAsync("/app/missing.js")).StatusCode);

        var bootstrap = _server.Bootstrap();
        var first = await _client.GetAsync("/app/index.dc.html?auth=" + bootstrap);
        Assert.Equal(HttpStatusCode.Redirect, first.StatusCode);
        Assert.Equal("/app/index.dc.html", first.Headers.Location?.ToString());
        Assert.Contains("HttpOnly", first.Headers.GetValues("Set-Cookie").Single());
        Assert.StartsWith($"writer_token_{_server.Port}=", first.Headers.GetValues("Set-Cookie").Single());
    }

    [Fact]
    public async Task Listing_follows_list_depth_and_keeps_the_500_newest()
    {
        var dir = Directory.CreateTempSubdirectory("writer-depth").FullName;
        try
        {
            Directory.CreateDirectory(Path.Combine(dir, "sub"));
            File.WriteAllText(Path.Combine(dir, "sub", "deep.md"), "# Deep\n");
            var start = DateTime.UtcNow.AddHours(-1);
            for (var i = 0; i < 505; i++)
            {
                var f = Path.Combine(dir, $"n{i:D3}.md");
                File.WriteAllText(f, "x");
                File.SetLastWriteTimeUtc(f, start.AddSeconds(i));
            }
            using var server = new Serve(0, requireToken: false, workspace: dir, listDepth: 0);
            server.Start();
            using var client = new HttpClient { BaseAddress = new Uri(server.Url) };
            var files = JsonDocument.Parse(await client.GetStringAsync("/files")).RootElement.GetProperty("files")
                .EnumerateArray().Select(f => f.GetProperty("path").GetString()).ToList();
            Assert.Equal(500, files.Count);
            Assert.Equal("n504.md", files[0]);
            Assert.DoesNotContain("n004.md", files);
            Assert.DoesNotContain("sub/deep.md", files);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact]
    public async Task Files_listing_tolerates_a_missing_workspace_and_includes_granted_files()
    {
        var missing = Path.Combine(Path.GetTempPath(), "writer-missing-" + Guid.NewGuid().ToString("N"));
        using var server = new Serve(0, requireToken: false, workspace: missing);
        server.Start();
        using var client = new HttpClient { BaseAddress = new Uri(server.Url) };
        var elsewhere = Directory.CreateTempSubdirectory("writer-granted-list").FullName;
        var granted = Path.Combine(elsewhere, "granted.md");
        File.WriteAllText(granted, "# g\n");
        server.Grant(granted);

        var response = await client.GetAsync("/files");
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var files = JsonDocument.Parse(await response.Content.ReadAsStringAsync()).RootElement.GetProperty("files")
            .EnumerateArray().Select(f => f.GetProperty("path").GetString()).ToList();
        Assert.Contains(granted.Replace('\\', '/'), files);
    }

    [Fact]
    public async Task Files_opened_in_the_window_are_listed_but_not_replaceable_like_a_Save_dialog_target()
    {
        var opened = Doc("opened.docx", P("Hello"));
        var other = Doc("other.docx", P("World"));
        Assert.True(_server.ApplyGrantLine("list " + opened));
        Assert.False(_server.ApplyGrantLine("list relative.docx"));
        Assert.True(_server.ApplyGrantLine("list " + Path.Combine(Path.GetTempPath(), "outside-the-workspace.md")));

        var listing = await (await _client.SendAsync(Request(HttpMethod.Get, "/files"))).Content.ReadAsStringAsync();
        var paths = JsonDocument.Parse(listing).RootElement.GetProperty("files").EnumerateArray().Select(f => f.GetProperty("path").GetString()).ToList();
        Assert.Single(paths, p => p == "opened.docx"); // relative, once, although the folder listing has it too
        Assert.DoesNotContain(paths, p => p!.Contains("outside-the-workspace"));

        // renaming another document onto it is refused, as for any existing file
        var move = Request(HttpMethod.Put, $"/file?file={Uri.EscapeDataString(opened)}&from={Uri.EscapeDataString(other)}");
        move.Content = new ByteArrayContent([]);
        Assert.Equal(HttpStatusCode.BadRequest, (await _client.SendAsync(move)).StatusCode);
        Assert.True(File.Exists(other));
    }

    public void Dispose()
    {
        _client.Dispose();
        _server.Dispose();
        Directory.Delete(_dir, true);
    }

    HttpRequestMessage Request(HttpMethod method, string path, string? body = null, bool auth = true)
    {
        var request = new HttpRequestMessage(method, path);
        if (auth) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _server.Token);
        if (body is not null) request.Content = new StringContent(body, Encoding.UTF8, "application/json");
        return request;
    }

    string Doc(string name, params DocumentFormat.OpenXml.OpenXmlElement[] blocks)
    {
        var file = Path.Combine(_dir, name);
        File.WriteAllBytes(file, Docx(blocks));
        return file;
    }

    [Fact]
    public async Task Requires_a_bearer_token_or_the_preview_cookie()
    {
        var file = Doc("a.docx", P("Hello"));
        var q = "?file=" + Uri.EscapeDataString(file);

        var anonymous = await _client.SendAsync(Request(HttpMethod.Get, "/outline" + q, auth: false));
        Assert.Equal(HttpStatusCode.Unauthorized, anonymous.StatusCode);
        var queryToken = await _client.SendAsync(Request(HttpMethod.Get, "/outline" + q + "&token=" + _server.Token, auth: false));
        Assert.Equal(HttpStatusCode.Unauthorized, queryToken.StatusCode);
        var wrong = new HttpRequestMessage(HttpMethod.Get, "/outline" + q);
        wrong.Headers.Authorization = new AuthenticationHeaderValue("Bearer", "nope");
        Assert.Equal(HttpStatusCode.Unauthorized, (await _client.SendAsync(wrong)).StatusCode);

        var header = await _client.SendAsync(Request(HttpMethod.Get, "/outline" + q));
        Assert.Equal(HttpStatusCode.OK, header.StatusCode);
        Assert.Contains("/body/paragraph[1]", await header.Content.ReadAsStringAsync());

        var bootstrap = _server.Bootstrap();
        var first = await _client.SendAsync(Request(HttpMethod.Get, "/" + q + "&auth=" + bootstrap, auth: false));
        Assert.Equal(HttpStatusCode.Redirect, first.StatusCode);
        Assert.Equal("/" + q, first.Headers.Location?.ToString());
        var cookie = first.Headers.GetValues("Set-Cookie").Single();
        Assert.Contains("HttpOnly", cookie);
        Assert.Contains("SameSite=Strict", cookie);
        Assert.DoesNotContain("auth=", first.Headers.Location?.ToString());

        var reused = await _client.SendAsync(Request(HttpMethod.Get, "/" + q + "&auth=" + bootstrap, auth: false));
        Assert.Equal(HttpStatusCode.Unauthorized, reused.StatusCode);

        var withCookie = new HttpRequestMessage(HttpMethod.Get, "/html" + q);
        withCookie.Headers.Add("Cookie", cookie.Split(';')[0]);
        var page = await _client.SendAsync(withCookie);
        Assert.Equal(HttpStatusCode.OK, page.StatusCode);
        Assert.Contains("Hello", await page.Content.ReadAsStringAsync());

        var previewNoCookie = await _client.SendAsync(Request(HttpMethod.Get, "/" + q, auth: false));
        Assert.Equal(HttpStatusCode.Unauthorized, previewNoCookie.StatusCode);
    }

    [Fact]
    public async Task Refuses_foreign_origins_and_non_json_posts()
    {
        var file = Doc("b.docx", P("Hello"));
        var q = "?file=" + Uri.EscapeDataString(file);

        var evil = Request(HttpMethod.Get, "/outline" + q);
        evil.Headers.Add("Origin", "http://evil.example");
        Assert.Equal(HttpStatusCode.Forbidden, (await _client.SendAsync(evil)).StatusCode);

        var preflight = new HttpRequestMessage(HttpMethod.Options, "/run");
        preflight.Headers.Add("Origin", "http://evil.example");
        Assert.Equal(HttpStatusCode.Forbidden, (await _client.SendAsync(preflight)).StatusCode);

        var allowed = Request(HttpMethod.Get, "/outline" + q);
        allowed.Headers.Add("Origin", "http://localhost:5173");
        var ok = await _client.SendAsync(allowed);
        Assert.Equal(HttpStatusCode.OK, ok.StatusCode);
        Assert.Equal("http://localhost:5173", ok.Headers.GetValues("Access-Control-Allow-Origin").Single());
        Assert.Equal("Origin", ok.Headers.GetValues("Vary").Single());

        var own = new HttpRequestMessage(HttpMethod.Options, "/run");
        own.Headers.Add("Origin", _server.Url);
        var ownPreflight = await _client.SendAsync(own);
        Assert.Equal(HttpStatusCode.NoContent, ownPreflight.StatusCode);
        Assert.Contains("POST", ownPreflight.Headers.GetValues("Access-Control-Allow-Methods").Single());

        var plain = Request(HttpMethod.Post, "/run");
        plain.Content = new StringContent("{\"command\":\"help\"}", Encoding.UTF8, "text/plain");
        Assert.Equal(HttpStatusCode.UnsupportedMediaType, (await _client.SendAsync(plain)).StatusCode);
    }

    [Fact]
    public async Task Runs_commands_and_renders()
    {
        var file = Doc("c.docx", P("Intro", "Heading1"), P("Body text"));
        var q = "?file=" + Uri.EscapeDataString(file);

        var run = await _client.SendAsync(Request(HttpMethod.Post, "/run", JsonSerializer.Serialize(new { command = $"view \"{file}\" text" })));
        Assert.Equal(HttpStatusCode.OK, run.StatusCode);
        var result = JsonDocument.Parse(await run.Content.ReadAsStringAsync()).RootElement;
        Assert.Equal(0, result.GetProperty("code").GetInt32());
        Assert.Equal("Intro\n\nBody text\n", result.GetProperty("output").GetString());

        var edit = await _client.SendAsync(Request(HttpMethod.Post, "/run", JsonSerializer.Serialize(new { argv = new[] { "set", file, "/body/paragraph[1]", "--prop", "text=Changed" } })));
        Assert.Equal(0, JsonDocument.Parse(await edit.Content.ReadAsStringAsync()).RootElement.GetProperty("code").GetInt32());

        var bad = await _client.SendAsync(Request(HttpMethod.Post, "/run", JsonSerializer.Serialize(new { command = $"get \"{file}\" /body/paragraph[9]" })));
        var badResult = JsonDocument.Parse(await bad.Content.ReadAsStringAsync()).RootElement;
        Assert.Equal(2, badResult.GetProperty("code").GetInt32());
        Assert.Equal("PATH_NOT_FOUND", badResult.GetProperty("error").GetProperty("code").GetString());

        var html = await (await _client.SendAsync(Request(HttpMethod.Get, "/html" + q))).Content.ReadAsStringAsync();
        Assert.Matches("<h1[^>]*>Intro</h1>", html);
        Assert.Contains("Changed", html);

        var json = await _client.SendAsync(Request(HttpMethod.Get, "/json" + q));
        Assert.Equal("document", JsonDocument.Parse(await json.Content.ReadAsStringAsync()).RootElement.GetProperty("kind").GetString());

        var missing = await _client.SendAsync(Request(HttpMethod.Get, "/html?file=" + Uri.EscapeDataString(Path.Combine(_dir, "nope.docx"))));
        Assert.Equal(HttpStatusCode.NotFound, missing.StatusCode);

        var page = await _client.SendAsync(Request(HttpMethod.Get, "/" + q));
        var pageText = await page.Content.ReadAsStringAsync();
        Assert.Contains("EventSource", pageText);
        Assert.DoesNotContain(_server.Token!, pageText);

        var nested = await _client.SendAsync(Request(HttpMethod.Post, "/run", JsonSerializer.Serialize(new { command = "serve" })));
        Assert.Equal(HttpStatusCode.OK, nested.StatusCode);
        var refused = JsonDocument.Parse(await nested.Content.ReadAsStringAsync()).RootElement;
        Assert.Equal(1, refused.GetProperty("code").GetInt32());
        Assert.Equal("USAGE", refused.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public async Task Events_report_file_changes()
    {
        var file = Path.Combine(_dir, "d.md");
        File.WriteAllText(file, "# One\n");
        var q = "?file=" + Uri.EscapeDataString(file) + "&token=" + _server.Token;
        using var response = await _client.SendAsync(Request(HttpMethod.Get, "/events" + q, auth: false), HttpCompletionOption.ResponseHeadersRead);
        Assert.Equal("text/event-stream", response.Content.Headers.ContentType?.MediaType);
        using var stream = await response.Content.ReadAsStreamAsync();
        using var reader = new StreamReader(stream);
        Assert.Equal("event: ready", await reader.ReadLineAsync());
        await Task.Delay(300);
        File.WriteAllText(file, "# Two\n");
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        string? line;
        while ((line = await reader.ReadLineAsync(timeout.Token)) is not null)
            if (line == "event: change") return;
        Assert.Fail("no change event received");
    }
}
