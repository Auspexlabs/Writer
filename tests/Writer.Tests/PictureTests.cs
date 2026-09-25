using System.Buffers.Binary;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats;

namespace Writer.Tests;

/// <summary>The picture tools on Word, PowerPoint and Excel pictures: each prop is Office markup that reads back and validates.</summary>
public class PictureTests : IDisposable
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    readonly string _png = Path.Combine(Path.GetTempPath(), $"writer-picture-{Guid.NewGuid():N}.png");

    public PictureTests() => File.WriteAllBytes(_png, Png.Encode(40, 30, (x, y) => (200, (byte)(x * 6), (byte)(y * 8), 255)));

    public void Dispose() => File.Delete(_png);

    public static TheoryData<string> Formats => new() { "docx", "pptx", "xlsx" };

    /// <summary>A new document of the format with one 4 × 3 cm picture.</summary>
    internal static (Document Doc, Node Image) WithPicture(string format, string png)
    {
        var doc = Adapters.ForName(format).Create();
        var parent = format switch
        {
            "docx" => doc.Root.Children.Single(),
            "pptx" => Mutations.Add(doc.Root, "slide", Props(("layout", "Blank")), null),
            _ => doc.Root.Children[0],
        };
        var size = format == "docx" ? Props(("src", png), ("width", "4cm"), ("height", "3cm")) : Props(("src", png), ("x", "2cm"), ("y", "2cm"), ("w", "4cm"), ("h", "3cm"));
        return (doc, Mutations.Add(parent, "image", size, null));
    }

    internal static byte[] Save(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return ms.ToArray();
    }

    internal static IEnumerable<string> Errors(string format, byte[] bytes)
    {
        using OpenXmlPackage package = format switch
        {
            "docx" => WordprocessingDocument.Open(new MemoryStream(bytes), false),
            "pptx" => PresentationDocument.Open(new MemoryStream(bytes), false),
            _ => SpreadsheetDocument.Open(new MemoryStream(bytes), false),
        };
        return new OpenXmlValidator(FileFormatVersions.Office2016).Validate(package).Select(e => $"{e.Part?.Uri} {e.Path?.XPath}: {e.Description}").ToList();
    }

    static string Shown(Node node, string name) => Registry.ToDisplay(node.Format, node.Kind, node.GetProps()).GetValueOrDefault(name) ?? "";

    [Theory]
    [MemberData(nameof(Formats))]
    public void Every_adjustment_round_trips_as_valid_office_markup(string format)
    {
        var (doc, image) = WithPicture(format, _png);
        using (doc)
        {
            var look = Props(("crop", "10,5,10,5"), ("rotation", "90"), ("flipH", "true"), ("flipV", "true"), ("brightness", "20"), ("contrast", "-30"),
                ("grayscale", "true"), ("transparency", "25"), ("line", "C00000"), ("lineWidth", "2.5pt"), ("shadow", "true"), ("geometry", "roundRect"));
            image = Mutations.Set(image, look);
            var expected = look.ToDictionary(p => p.Key, p => p.Value == "2.5pt" ? "2.5" : p.Value);
            foreach (var (name, value) in expected) Assert.Equal(value, image.GetProps().GetValueOrDefault(name));
            var raw = image.GetRaw();
            foreach (var tag in new[] { "a:srcRect", "rot=\"5400000\"", "flipH=\"1\"", "a:lum bright=\"20000\" contrast=\"-30000\"", "a:grayscl", "a:alphaModFix amt=\"75000\"",
                "a:ln w=\"31750\"", "a:srgbClr val=\"C00000\"", "a:outerShdw", "prst=\"roundRect\"" })
                Assert.Contains(tag, raw);

            var bytes = Save(doc);
            Assert.Empty(Errors(format, bytes));
            using var reopened = Adapters.ForName(format).Open(new MemoryStream(bytes));
            var again = PathResolver.Single(reopened.Root, image.Path);
            foreach (var (name, value) in expected) Assert.Equal(value, again.GetProps().GetValueOrDefault(name));

            // zero, false and none take each adjustment out of the file again
            var cleared = Mutations.Set(again, Props(("crop", "0,0,0,0"), ("rotation", "0"), ("flipH", "false"), ("flipV", "false"), ("brightness", "0"), ("contrast", "0"),
                ("grayscale", "false"), ("transparency", "0"), ("line", "none"), ("shadow", "false"), ("geometry", "rect")));
            foreach (var name in expected.Keys) Assert.False(cleared.GetProps().ContainsKey(name), name);
            foreach (var tag in new[] { "a:srcRect", "rot=", "a:lum", "a:grayscl", "a:alphaModFix", "a:ln", "a:effectLst" }) Assert.DoesNotContain(tag, cleared.GetRaw());
            Assert.Empty(Errors(format, Save(reopened)));
        }
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void Cropping_keeps_the_scale_and_uncropping_restores_the_frame(string format)
    {
        var (doc, image) = WithPicture(format, _png);
        using (doc)
        {
            var (w, h) = format == "docx" ? ("width", "height") : ("w", "h");
            image = Mutations.Set(image, Props(("crop", "25,0,0,10")));
            Assert.Equal(("3cm", "2.7cm"), (Shown(image, w), Shown(image, h)));
            if (format != "docx") Assert.Equal(("3cm", "2cm"), (Shown(image, "x"), Shown(image, "y"))); // what is still shown stays put
            image = Mutations.Set(image, Props(("crop", "0,0,0,0")));
            Assert.Equal(("4cm", "3cm"), (Shown(image, w), Shown(image, h)));
            if (format != "docx") Assert.Equal(("2cm", "2cm"), (Shown(image, "x"), Shown(image, "y")));
            var ex = Assert.Throws<WriterException>(() => Mutations.Set(image, Props(("crop", "60,0,50,0"))));
            Assert.Contains("left,top,right,bottom", ex.Hint);
        }
    }

    [Fact]
    public void Word_reserves_room_for_a_rotated_picture_its_border_and_shadow()
    {
        var (doc, image) = WithPicture("docx", _png);
        using (doc)
        {
            image = Mutations.Set(image, Props(("rotation", "90")));
            Assert.Contains("<wp:effectExtent l=\"-180000\" t=\"180000\" r=\"-180000\" b=\"180000\" />", image.GetRaw()); // 4 × 3 cm turned: half a cm each way
            image = Mutations.Set(image, Props(("rotation", "0"), ("lineWidth", "2pt")));
            Assert.Contains("<wp:effectExtent l=\"12700\" t=\"12700\" r=\"12700\" b=\"12700\" />", image.GetRaw());
        }
    }

    [Theory]
    [MemberData(nameof(Formats))]
    public void Replacing_keeps_the_width_and_follows_the_new_picture(string format)
    {
        var (doc, image) = WithPicture(format, _png);
        using (doc)
        {
            var tall = Path.Combine(Path.GetTempPath(), $"writer-tall-{Guid.NewGuid():N}.png");
            File.WriteAllBytes(tall, Png.Encode(10, 20, (_, _) => (0, 0, 255, 255)));
            try
            {
                image = Mutations.Set(image, Props(("crop", "10,0,0,0")));
                var before = image.GetProps()["src"];
                image = Mutations.Set(image, Props(("src", tall)));
                var (w, h) = format == "docx" ? ("width", "height") : ("w", "h");
                Assert.Equal(("3.6cm", "7.2cm"), (Shown(image, w), Shown(image, h)));
                Assert.False(image.GetProps().ContainsKey("crop"));
                Assert.NotEqual(before, image.GetProps()["src"]);
                Assert.Equal(new FileInfo(tall).Length.ToString(), image.GetProps()["bytes"]);
                Assert.Empty(Errors(format, Save(doc)));
            }
            finally { File.Delete(tall); }
        }
    }

    [Fact]
    public void Excel_pictures_are_added_found_and_removed_with_their_image()
    {
        using var doc = Adapters.ForName("xlsx").Create();
        var sheet = doc.Root.Children[0];
        var image = Mutations.Add(sheet, "image", Props(("src", _png), ("alt", "logo")), null);
        Assert.Equal("/sheet[1]/image[1]", image.Path);
        Assert.Equal(("1.058cm", "0.794cm", "logo"), (Shown(image, "w"), Shown(image, "h"), Shown(image, "alt"))); // 40 × 30 px at 96 dpi
        var id = image.GetProps()["id"];
        Assert.Same(image.Anchor, PathResolver.Single(doc.Root, $"/sheet[1]/image[@id={id}]").Anchor);
        image = Mutations.Set(image, Props(("x", "5cm"), ("y", "1cm")));
        Assert.Equal(("5cm", "1cm"), (Shown(image, "x"), Shown(image, "y")));
        Assert.Equal(File.ReadAllBytes(_png), image.GetBinary()!.Value.Data);
        var bytes = Save(doc);
        Assert.Empty(Errors("xlsx", bytes));
        using (var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false))
        {
            var anchor = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>().Single();
            Assert.Equal("oneCell", anchor.EditAs!.InnerText); // moves with its cells but keeps its size, as Excel inserts pictures
        }
        image.Remove();
        Assert.DoesNotContain(Mutations.Refresh(sheet).Children, c => c.Kind == "image");
        using var package2 = SpreadsheetDocument.Open(new MemoryStream(Save(doc)), false);
        Assert.Null(package2.WorkbookPart!.WorksheetParts.Single().DrawingsPart);
    }

    static Node Picture(Document doc, string id) => PathResolver.Single(doc.Root, $"//image[@id={id}]");

    static Document Reopen(Document doc)
    {
        var bytes = Save(doc);
        Assert.Empty(Errors("docx", bytes));
        return Adapters.ForName("docx").Open(new MemoryStream(bytes));
    }

    [Fact]
    public void A_Word_picture_floats_in_a_paragraph_at_its_place_and_comes_back_inline()
    {
        var (doc, image) = WithPicture("docx", _png);
        using (doc)
        {
            var body = doc.Root.Children.Single();
            Mutations.Add(body, "paragraph", Props(("text", "Text the picture floats over.")), null);
            image = Mutations.Set(image, Props(("crop", "10,0,0,0"), ("alt", "logo")));
            var id = image.GetProps()["id"];
            Assert.Equal("inline", Shown(image, "wrap"));

            image = Mutations.Move(image, PathResolver.Single(doc.Root, "/body/paragraph[1]"), null);
            Assert.Equal("/body/paragraph[1]/image[1]", image.Path);
            Assert.Equal(("square", "0cm", "0cm", "column", "paragraph"), (Shown(image, "wrap"), Shown(image, "x"), Shown(image, "y"), Shown(image, "xFrom"), Shown(image, "yFrom")));
            Assert.DoesNotContain(body.Children, c => c.Kind == "image"); // the block the picture was in goes with it
            Mutations.Set(image, Props(("wrap", "front"), ("x", "2.5cm"), ("y", "1cm"), ("xFrom", "page")));

            using (var again = Reopen(doc))
            {
                var floating = Picture(again, id);
                Assert.Equal("/body/paragraph[1]/image[1]", floating.Path);
                foreach (var (name, value) in new[] { ("wrap", "front"), ("x", "2.5cm"), ("y", "1cm"), ("xFrom", "page"), ("yFrom", "paragraph"), ("width", "3.6cm"), ("height", "3cm"), ("crop", "10,0,0,0"), ("alt", "logo") })
                    Assert.Equal(value, Shown(floating, name));
                Assert.Equal("Text the picture floats over.", PathResolver.Single(again.Root, "/body/paragraph[1]").Text);

                var inline = Mutations.Move(floating, again.Root.Children.Single(), 1);
                Assert.Equal("/body/image[1]", inline.Path);
                using var back = Reopen(again);
                var block = Picture(back, id);
                Assert.Equal("/body/image[1]", block.Path);
                foreach (var (name, value) in new[] { ("wrap", "inline"), ("width", "3.6cm"), ("height", "3cm"), ("crop", "10,0,0,0"), ("alt", "logo") })
                    Assert.Equal(value, Shown(block, name));
                Assert.False(block.GetProps().ContainsKey("x"));
                Assert.Equal("Text the picture floats over.", PathResolver.Single(back.Root, "/body/paragraph[1]").Text);
            }
        }
    }

    [Fact]
    public void Every_Word_wrap_and_alignment_round_trips()
    {
        using var doc = Adapters.ForName("docx").Create();
        var paragraph = Mutations.Add(doc.Root.Children.Single(), "paragraph", Props(("text", "Anchor")), null);
        var image = Mutations.Add(paragraph, "image", Props(("src", _png), ("width", "2cm"), ("wrap", "behind"), ("x", "1cm"), ("y", "-0.5cm")), null);
        var id = image.GetProps()["id"];
        Assert.Equal(("behind", "1cm", "-0.5cm"), (Shown(image, "wrap"), Shown(image, "x"), Shown(image, "y")));
        foreach (var wrap in new[] { "square", "tight", "through", "topBottom", "front", "behind" })
        {
            Mutations.Set(Picture(doc, id), Props(("wrap", wrap)));
            using var again = Reopen(doc);
            Assert.Equal(wrap, Shown(Picture(again, id), "wrap"));
        }
        Mutations.Set(Picture(doc, id), Props(("xAlign", "center"), ("xFrom", "margin"), ("yAlign", "bottom"), ("yFrom", "page")));
        using (var again = Reopen(doc))
        {
            var p = Picture(again, id).GetProps();
            Assert.Equal(("center", "margin", "bottom", "page"), (p["xAlign"], p["xFrom"], p["yAlign"], p["yFrom"]));
            Assert.False(p.ContainsKey("x") || p.ContainsKey("y"), "an alignment replaces the offset");
        }
        Mutations.Set(Picture(doc, id), Props(("x", "3cm")));
        var shown = Picture(doc, id).GetProps();
        Assert.True(shown.ContainsKey("x") && !shown.ContainsKey("xAlign"), "an offset replaces the alignment");
        Mutations.Set(Picture(doc, id), Props(("wrap", "inline")));
        using (var again = Reopen(doc))
        {
            var p = Picture(again, id).GetProps();
            Assert.Equal("inline", p["wrap"]);
            Assert.False(p.ContainsKey("x") || p.ContainsKey("xFrom"));
            Assert.Equal("Anchor", PathResolver.Single(again.Root, "/body/paragraph[1]").Text);
        }
    }

    [Fact]
    public void Floating_pictures_from_Word_keep_their_place_through_edits()
    {
        using var file = File.OpenRead(Path.Combine(TestDocs.FixtureDir("docx"), "pictures.docx"));
        using var doc = Adapters.ForName("docx").Open(file);
        var expected = new Dictionary<string, (string Wrap, string XFrom, string X, string YFrom, string Y)>
        {
            ["4"] = ("behind", "margin", "xAlign=center", "margin", "yAlign=center"),
            ["5"] = ("square", "margin", "xAlign=right", "paragraph", "0cm"),
            ["6"] = ("tight", "margin", "2cm", "paragraph", "1cm"),
        };
        void Check(Document d)
        {
            foreach (var (id, e) in expected)
            {
                var p = Registry.ToDisplay("docx", "image", Picture(d, id).GetProps());
                Assert.Equal(e, (p["wrap"], p["xFrom"], p.TryGetValue("x", out var x) ? x : "xAlign=" + p["xAlign"], p["yFrom"], p.TryGetValue("y", out var y) ? y : "yAlign=" + p["yAlign"]));
            }
        }
        Check(doc);
        foreach (var id in expected.Keys)
        {
            var paragraph = Picture(doc, id).Parent!;
            Mutations.Set(paragraph, Props(("html", "Edited <b>text</b> of the paragraph the picture floats in.")));
        }
        using var again = Adapters.ForName("docx").Open(new MemoryStream(Save(doc)));
        Check(again);
        Assert.Equal("Edited text of the paragraph the picture floats in.", Picture(again, "6").Parent!.Text);
    }

    [Fact]
    public void Office_pictures_read_their_look_from_the_file()
    {
        using var doc = Adapters.ForName("pptx").Open(File.OpenRead(Path.Combine(TestDocs.FixtureDir("pptx"), "pictures-basic.pptx")));
        var image = PathResolver.Single(doc.Root, "/slide[1]/image[3]");
        var props = image.GetProps();
        Assert.Equal("100005", props["id"]);
        Assert.True(long.Parse(props["bytes"]) > 0);
        Assert.False(props.ContainsKey("crop"));
    }
}

/// <summary>A picture's image through the writer-vision helper: a stand-in script that plays the helper's part, and — on macOS
/// with the helper built — Apple Vision itself. One collection, because the tests point WRITER_VISION at their helper.</summary>
[Collection("writer-vision")]
public class PictureHelperTests : IDisposable
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    readonly string _dir = Directory.CreateTempSubdirectory("writer-helper-").FullName;
    readonly string? _before = Environment.GetEnvironmentVariable("WRITER_VISION");

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("WRITER_VISION", _before);
        Directory.Delete(_dir, true);
    }

    /// <summary>A helper that answers every command with the given picture (and records its arguments), or fails.</summary>
    [System.Runtime.Versioning.UnsupportedOSPlatform("windows")]
    string FakeHelper(byte[]? answer, string? failure = null)
    {
        var script = Path.Combine(_dir, "writer-vision");
        var reply = Path.Combine(_dir, "reply.png");
        if (answer is not null) File.WriteAllBytes(reply, answer);
        File.WriteAllText(script, "#!/bin/sh\n" + $"echo \"$@\" > '{Path.Combine(_dir, "args")}'\n"
            + (failure is null ? $"cp '{reply}' \"$3\"\n" : $"echo '{failure}' >&2\nexit 3\n"));
        File.SetUnixFileMode(script, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        Environment.SetEnvironmentVariable("WRITER_VISION", script);
        return script;
    }

    string Args => File.ReadAllText(Path.Combine(_dir, "args")).Trim();

    string Png(string name, int w, int h, Func<int, int, (byte, byte, byte, byte)> pixel)
    {
        var path = Path.Combine(_dir, name);
        File.WriteAllBytes(path, Tests.Png.Encode(w, h, pixel));
        return path;
    }

    [Theory]
    [MemberData(nameof(PictureTests.Formats), MemberType = typeof(PictureTests))]
    public void Background_removal_keeps_the_original_and_reset_restores_it(string format)
    {
        if (OperatingSystem.IsWindows()) return; // the stand-in helper is a shell script
        var photo = Png("photo.png", 40, 30, (x, y) => (10, 200, (byte)x, 255));
        var cut = Tests.Png.Encode(40, 30, (x, y) => (10, 200, (byte)x, (byte)(x > 10 && x < 30 ? 255 : 0)));
        FakeHelper(cut);
        var (doc, image) = PictureTests.WithPicture(format, photo);
        using (doc)
        {
            image = Mutations.Set(image, Props(("brightness", "20"), ("crop", "10,0,0,0")));
            image = Mutations.Set(image, Props(("background", "remove")));
            Assert.StartsWith("cutout ", Args);
            Assert.Equal(cut, image.GetBinary()!.Value.Data);
            Assert.Equal("image/png", image.GetBinary()!.Value.ContentType);
            Assert.Equal(("20", "10,0,0,0"), (image.GetProps()["brightness"], image.GetProps()["crop"])); // adjustments stay on the cut-out
            var bytes = PictureTests.Save(doc);
            Assert.Empty(PictureTests.Errors(format, bytes));

            // a second cut-out still keeps the very first picture
            using var reopened = Adapters.ForName(format).Open(new MemoryStream(bytes));
            var again = Mutations.Set(PathResolver.Single(reopened.Root, image.Path), Props(("background", "remove")));
            again = Mutations.Set(again, Props(("reset", "true")));
            Assert.Equal(File.ReadAllBytes(photo), again.GetBinary()!.Value.Data);
            Assert.False(again.GetProps().ContainsKey("brightness"));
            Assert.False(again.GetProps().ContainsKey("crop"));
            var (w, h) = format == "docx" ? ("width", "height") : ("w", "h");
            Assert.Equal(("4cm", "3cm"), (Registry.ToDisplay(format, "image", again.GetProps())[w], Registry.ToDisplay(format, "image", again.GetProps())[h]));
            var saved = PictureTests.Save(reopened);
            Assert.Empty(PictureTests.Errors(format, saved));
            using OpenXmlPackage package = format switch
            {
                "docx" => WordprocessingDocument.Open(new MemoryStream(saved), false),
                "pptx" => PresentationDocument.Open(new MemoryStream(saved), false),
                _ => SpreadsheetDocument.Open(new MemoryStream(saved), false),
            };
            Assert.Single(package.GetAllParts().OfType<ImagePart>()); // the cut-outs and the kept copy are gone
        }
    }

    [Fact]
    public void Compress_shrinks_to_what_the_frame_needs_and_drops_the_cropped_areas()
    {
        if (OperatingSystem.IsWindows()) return;
        var big = Png("big.png", 2400, 1800, (x, y) => ((byte)(x * 7), (byte)(y * 13), (byte)((x ^ y) & 0xFF), 255));
        FakeHelper(Tests.Png.Encode(300, 200, (_, _) => (1, 2, 3, 255)));
        var (doc, image) = PictureTests.WithPicture("pptx", big);
        using (doc)
        {
            image = Mutations.Set(image, Props(("crop", "25,0,0,0"), ("w", "5.08cm"), ("h", "5.08cm")));
            var before = long.Parse(image.GetProps()["bytes"]);
            image = Mutations.Set(image, Props(("compress", "web")));
            Assert.Equal("compress", Args.Split(' ')[0]);
            Assert.Contains("--max 300x300 --crop 0.25,0,0,0", Args); // 2 in at 150 ppi, the cropped quarter cut away
            Assert.True(long.Parse(image.GetProps()["bytes"]) < before);
            Assert.False(image.GetProps().ContainsKey("crop"));
            Assert.Equal(("5.08cm", "5.08cm"), (Registry.ToDisplay("pptx", "image", image.GetProps())["w"], Registry.ToDisplay("pptx", "image", image.GetProps())["h"]));

            var small = long.Parse(image.GetProps()["bytes"]);
            File.Delete(Path.Combine(_dir, "args"));
            image = Mutations.Set(image, Props(("compress", "print")));
            Assert.False(File.Exists(Path.Combine(_dir, "args")), "a picture already no bigger than it is shown is left alone");
            Assert.Equal(small, long.Parse(image.GetProps()["bytes"]));
        }
    }

    [Fact]
    public void A_missing_or_failing_helper_leaves_the_picture_alone_with_a_clear_error()
    {
        if (OperatingSystem.IsWindows()) return;
        var photo = Png("photo.png", 40, 30, (_, _) => (9, 9, 9, 255));
        var (doc, image) = PictureTests.WithPicture("docx", photo);
        using (doc)
        {
            Environment.SetEnvironmentVariable("WRITER_VISION", Path.Combine(_dir, "nothing-here"));
            if (Writer.Formats.Common.Vision.Find() is null)
            {
                var missing = Assert.Throws<WriterException>(() => Mutations.Set(image, Props(("background", "remove"))));
                Assert.Contains("抠图需要 macOS 14", missing.Message);
            }
            FakeHelper(null, "no subject found");
            var failed = Assert.Throws<WriterException>(() => Mutations.Set(image, Props(("background", "remove"))));
            Assert.Contains("no subject found", failed.Message);
            Assert.Equal(File.ReadAllBytes(photo), Mutations.Refresh(image).GetBinary()!.Value.Data);
        }
    }

    /// <summary>The helper built by desktop/scripts/build-vision.mjs, or WRITER_VISION.</summary>
    internal static string? BuiltHelper()
    {
        if (Environment.GetEnvironmentVariable("WRITER_VISION") is { Length: > 0 } env && File.Exists(env)) return env;
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            var bins = Path.Combine(dir.FullName, "desktop", "src-tauri", "binaries");
            if (!Directory.Exists(bins)) continue;
            var host = System.Runtime.InteropServices.RuntimeInformation.OSArchitecture == System.Runtime.InteropServices.Architecture.Arm64 ? "aarch64" : "x86_64";
            return new[] { $"writer-vision-{host}-apple-darwin", "writer-vision-universal-apple-darwin" }.Select(n => Path.Combine(bins, n)).FirstOrDefault(File.Exists);
        }
        return null;
    }

    [VisionFact]
    public void Apple_Vision_cuts_a_disc_out_of_a_white_background()
    {
        Environment.SetEnvironmentVariable("WRITER_VISION", BuiltHelper());
        var disc = Png("disc.png", 400, 400, (x, y) => (x - 200) * (x - 200) + (y - 200) * (y - 200) < 120 * 120 ? ((byte)200, (byte)40, (byte)40, (byte)255) : ((byte)255, (byte)255, (byte)255, (byte)255));
        var (doc, image) = PictureTests.WithPicture("pptx", disc);
        using (doc)
        {
            image = Mutations.Set(image, Props(("background", "remove")));
            var (w, h, rgba) = Tests.Png.Decode(image.GetBinary()!.Value.Data);
            Assert.Equal((400, 400), (w, h));
            byte Alpha(int x, int y) => rgba[(y * w + x) * 4 + 3];
            Assert.All(new[] { (2, 2), (397, 2), (2, 397), (397, 397) }, p => Assert.Equal(0, Alpha(p.Item1, p.Item2)));
            Assert.Equal(255, Alpha(200, 200));
        }
    }

    [VisionFact]
    public void The_real_helper_compresses_a_large_picture()
    {
        Environment.SetEnvironmentVariable("WRITER_VISION", BuiltHelper());
        var big = Png("big.png", 3000, 2000, (x, y) => ((byte)(x * 7 + y), (byte)(y * 13), (byte)((x ^ y) & 0xFF), 255));
        var (doc, image) = PictureTests.WithPicture("docx", big);
        using (doc)
        {
            var before = long.Parse(image.GetProps()["bytes"]);
            image = Mutations.Set(image, Props(("compress", "print")));
            var after = long.Parse(image.GetProps()["bytes"]);
            Assert.True(after < before / 4, $"{before} → {after} bytes");
            var (w, h, _) = Tests.Png.Decode(image.GetBinary()!.Value.Data);
            Assert.True(w <= 347 && h <= 260, $"{w} × {h}"); // 4 × 3 cm at 220 ppi
            Assert.Equal(("4cm", "3cm"), (Registry.ToDisplay("docx", "image", image.GetProps())["width"], Registry.ToDisplay("docx", "image", image.GetProps())["height"]));
        }
    }
}

/// <summary>Runs on macOS once the writer-vision helper is built (node desktop/scripts/build-vision.mjs); skipped elsewhere.</summary>
public sealed class VisionFactAttribute : FactAttribute
{
    public VisionFactAttribute()
    {
        if (!OperatingSystem.IsMacOS()) Skip = "Apple Vision runs on macOS only.";
        else if (PictureHelperTests.BuiltHelper() is null) Skip = "Build the helper first: node desktop/scripts/build-vision.mjs";
    }
}

/// <summary>Just enough PNG for tests: 8-bit RGBA written unfiltered, and 8-bit RGB or RGBA read back.</summary>
static class Png
{
    public static byte[] Encode(int width, int height, Func<int, int, (byte R, byte G, byte B, byte A)> pixel)
    {
        var raw = new MemoryStream();
        for (var y = 0; y < height; y++)
        {
            raw.WriteByte(0);
            for (var x = 0; x < width; x++)
            {
                var (r, g, b, a) = pixel(x, y);
                raw.Write([r, g, b, a]);
            }
        }
        raw.Position = 0;
        var z = new MemoryStream();
        using (var deflate = new ZLibStream(z, CompressionLevel.Fastest, leaveOpen: true)) raw.CopyTo(deflate);
        var png = new MemoryStream();
        png.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
        var header = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(header, width);
        BinaryPrimitives.WriteInt32BigEndian(header.AsSpan(4), height);
        header[8] = 8;
        header[9] = 6;
        Chunk(png, "IHDR", header);
        Chunk(png, "IDAT", z.ToArray());
        Chunk(png, "IEND", []);
        return png.ToArray();
    }

    static void Chunk(Stream s, string type, byte[] data)
    {
        var head = new byte[8];
        BinaryPrimitives.WriteInt32BigEndian(head, data.Length);
        System.Text.Encoding.ASCII.GetBytes(type, head.AsSpan(4));
        s.Write(head);
        s.Write(data);
        var crc = new byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(crc, Crc(head.AsSpan(4).ToArray().Concat(data).ToArray()));
        s.Write(crc);
    }

    static uint Crc(byte[] bytes)
    {
        var c = 0xFFFFFFFFu;
        foreach (var b in bytes)
        {
            c ^= b;
            for (var k = 0; k < 8; k++) c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
        }
        return c ^ 0xFFFFFFFFu;
    }

    /// <summary>Pixels as RGBA, for 8-bit RGB or RGBA pictures without interlacing (what ImageIO writes).</summary>
    public static (int Width, int Height, byte[] Rgba) Decode(byte[] png)
    {
        int width = 0, height = 0, channels = 0;
        var idat = new MemoryStream();
        for (var i = 8; i + 8 <= png.Length;)
        {
            var length = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(i));
            var type = System.Text.Encoding.ASCII.GetString(png, i + 4, 4);
            var data = png.AsSpan(i + 8, length);
            if (type == "IHDR")
            {
                (width, height) = (BinaryPrimitives.ReadInt32BigEndian(data), BinaryPrimitives.ReadInt32BigEndian(data[4..]));
                Assert.Equal(8, data[8]);
                channels = data[9] switch { 6 => 4, 2 => 3, _ => throw new NotSupportedException($"PNG colour type {data[9]}") };
                Assert.Equal(0, data[12]);
            }
            if (type == "IDAT") idat.Write(data);
            i += 12 + length;
        }
        idat.Position = 0;
        var raw = new MemoryStream();
        using (var inflate = new ZLibStream(idat, CompressionMode.Decompress)) inflate.CopyTo(raw);
        var bytes = raw.ToArray();
        int stride = width * channels;
        var rows = new byte[height * stride];
        for (var y = 0; y < height; y++)
        {
            var filter = bytes[y * (stride + 1)];
            for (var x = 0; x < stride; x++)
            {
                int a = x >= channels ? rows[y * stride + x - channels] : 0, b = y > 0 ? rows[(y - 1) * stride + x] : 0, c = x >= channels && y > 0 ? rows[(y - 1) * stride + x - channels] : 0;
                int p = a + b - c, pa = Math.Abs(p - a), pb = Math.Abs(p - b), pc = Math.Abs(p - c);
                var predictor = filter switch { 0 => 0, 1 => a, 2 => b, 3 => (a + b) / 2, 4 => pa <= pb && pa <= pc ? a : pb <= pc ? b : c, _ => throw new NotSupportedException($"filter {filter}") };
                rows[y * stride + x] = (byte)(bytes[y * (stride + 1) + 1 + x] + predictor);
            }
        }
        var rgba = new byte[width * height * 4];
        for (var i = 0; i < width * height; i++)
        {
            rows.AsSpan(i * channels, 3).CopyTo(rgba.AsSpan(i * 4));
            rgba[i * 4 + 3] = channels == 4 ? rows[i * channels + 3] : (byte)255;
        }
        return (width, height, rgba);
    }
}
