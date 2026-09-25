using System.Text;
using Writer.Core;
using Writer.Formats.Compat;

namespace Writer.Tests;

/// <summary>PowerPoint 97-2003 reading. There is no PowerPoint here, so the fixture is a .ppt this file builds itself:
/// one slide with a title and body placeholder whose texts live in the SlideListWithText, a bold free text box and a PNG.</summary>
public class PptReaderTests
{
    static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    [Fact]
    public void Reads_the_built_deck()
    {
        var warnings = new List<string>();
        var deck = PptReader.Read(BuildPpt(), warnings);

        Assert.Equal(25.4, deck.WidthCm, 2);
        Assert.Equal(19.05, deck.HeightCm, 2);
        var slide = Assert.Single(deck.Slides);
        Assert.Equal(4, slide.Shapes.Count);

        var title = slide.Shapes[0];
        Assert.True(title.IsTitle);
        Assert.Equal("Hello 演示", Assert.Single(title.Paragraphs).Html);
        Assert.Equal(288 / 576.0 * 2.54, title.X, 3);
        Assert.Equal(173 / 576.0 * 2.54, title.Y, 3);

        var body = slide.Shapes[1];
        Assert.False(body.IsTitle);
        Assert.Equal(2, body.Paragraphs.Count);
        Assert.Equal(("bullet", 0, "<span style=\"font-size:24pt;\">First point</span>"), (body.Paragraphs[0].List, body.Paragraphs[0].Level, body.Paragraphs[0].Html));
        Assert.Equal(("bullet", 1, "<span style=\"font-size:24pt;\">Second point</span>"), (body.Paragraphs[1].List, body.Paragraphs[1].Level, body.Paragraphs[1].Html));
        Assert.Equal(24, body.SizePt);

        var free = slide.Shapes[2];
        Assert.Equal("<b>Free text</b>", Assert.Single(free.Paragraphs).Html);
        Assert.Null(free.Paragraphs[0].List);

        var pic = slide.Shapes[3];
        Assert.Equal(ShapeKind.Image, pic.Kind);
        Assert.Equal(Png, pic.Image);
        Assert.Equal((2880 / 576.0 * 2.54, 1440 / 576.0 * 2.54, 1152 / 576.0 * 2.54, 864 / 576.0 * 2.54), (pic.X, pic.Y, pic.W, pic.H));
        Assert.Empty(warnings);
    }

    [Fact]
    public void Encrypted_files_are_refused()
    {
        var e = Assert.Throws<WriterException>(() => PptReader.Read(BuildPpt(encrypted: true), []));
        Assert.Equal(ErrorCode.FormatError, e.Code);
        Assert.Contains("password", e.Message);
    }

    [Fact]
    public void Fixture_file_matches_the_builder()
    {
        // WRITER_UPDATE_FIXTURES=1 rewrites tests/Writer.Tests/Fixtures/compat/sample.ppt from the builder
        var bytes = BuildPpt();
        if (Environment.GetEnvironmentVariable("WRITER_UPDATE_FIXTURES") == "1")
        {
            var src = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Fixtures", "compat", "sample.ppt"));
            Directory.CreateDirectory(Path.GetDirectoryName(src)!);
            File.WriteAllBytes(src, bytes);
            return;
        }
        Assert.Equal(bytes, File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "compat", "sample.ppt")));
    }

    // ---- a minimal PowerPoint 97-2003 file ----

    static byte[] BuildPpt(bool encrypted = false)
    {
        // the document container at offset 0 (persist 1), the slide (persist 2), then the persist directory and the user edit
        var document = Rec(0x03E8, 0, 0xF,
            Rec(0x03E9, 0, 1, I32(5760), I32(4320), I32(4320), I32(5760), I32(1), I32(2), I32(0), I32(0), U16(1), U16(0), new byte[4]),
            Rec(0x03F2, 0, 0xF, Rec(0x07D5, 0, 0xF, Rec(0x0FB7, 0, 0, FontName("Arial"), new byte[] { 0, 4, 0, 0 }))),
            Rec(0x0FF0, 0, 0xF,
                Rec(0x03F3, 0, 0, I32(2), I32(0), I32(2), I32(256), I32(0)),
                Rec(0x0F9F, 0, 0, I32(0)), Rec(0x0FA0, 0, 0, Utf16("Hello 演示")),
                Rec(0x0F9F, 0, 0, I32(1)), Rec(0x0FA0, 0, 0, Utf16("First point\rSecond point")),
                Rec(0x0FA1, 0, 0,
                    // paragraph runs: 12 chars level 0 bullet + left, 13 chars level 1 bullet
                    I32(12), U16(0), I32(0x801), U16(1), U16(0),
                    I32(13), U16(1), I32(0x001), U16(1),
                    // character run: 25 chars at 24 pt
                    I32(25), I32(0x20000), U16(24))),
            Rec(0x040B, 0, 0xF, Rec(0xF000, 0, 0xF,
                Rec(0xF006, 0, 0, I32(1027), I32(2), I32(1), I32(1), I32(1), I32(1)),
                Rec(0xF001, 1, 0xF, Rec(0xF007, 6, 2, new byte[] { 6, 6 }, new byte[16], U16(0xFF), I32(8 + 17 + Png.Length), I32(1), I32(0), new byte[] { 0, 0, 0, 0 })))),
            Rec(0x03EA, 0, 0));

        var slide = Rec(0x03EE, 0, 0xF,
            Rec(0x03EF, 0, 2, I32(0), new byte[] { 15, 16, 0, 0, 0, 0, 0, 0 }, I32(0), I32(0), U16(0), U16(0)),
            Rec(0x040C, 0, 0xF, Rec(0xF002, 0, 0xF,
                Rec(0xF008, 1, 0, I32(5), I32(1028)),
                Rec(0xF003, 0, 0xF,
                    Rec(0xF004, 0, 0xF, Rec(0xF009, 0, 1, new byte[16]), Rec(0xF00A, 0, 2, I32(1024), I32(0x5))),
                    // title placeholder → outline text 0
                    Rec(0xF004, 0, 0xF, Rec(0xF00A, 1, 2, I32(1025), I32(0xA00)), Opt((0x01BF, 0x00100000), (0x01FF, 0x00080000)),
                        Anchor(173, 288, 5472, 893), Rec(0xF011, 0, 0xF, Rec(0x0BC3, 0, 0, I32(-1), new byte[] { 15, 0, 0, 0 })),
                        Rec(0xF00D, 0, 0xF, Rec(0x0F9E, 0, 0, I32(0)))),
                    // body placeholder → outline text 1
                    Rec(0xF004, 0, 0xF, Rec(0xF00A, 1, 2, I32(1026), I32(0xA00)), Opt((0x01BF, 0x00100000), (0x01FF, 0x00080000)),
                        Anchor(1008, 288, 5472, 3859), Rec(0xF011, 0, 0xF, Rec(0x0BC3, 0, 0, I32(-1), new byte[] { 13, 0, 0, 0 })),
                        Rec(0xF00D, 0, 0xF, Rec(0x0F9E, 0, 0, I32(1)))),
                    // a free text box with its text inline, bold
                    Rec(0xF004, 0, 0xF, Rec(0xF00A, 202, 2, I32(1027), I32(0xA00)), Opt((0x01BF, 0x00100000), (0x01FF, 0x00080000)),
                        Anchor(3900, 288, 2600, 4100),
                        Rec(0xF00D, 0, 0xF, Rec(0x0F9F, 0, 0, I32(4)), Rec(0x0FA0, 0, 0, Utf16("Free text")),
                            Rec(0x0FA1, 0, 0, I32(10), U16(0), I32(0), I32(10), I32(0x1), U16(1)))),
                    // a picture frame showing BLIP 1
                    Rec(0xF004, 0, 0xF, Rec(0xF00A, 75, 2, I32(1028), I32(0xA00)), Opt((0x4104, 1), (0x01BF, 0x00100000), (0x01FF, 0x00080000)),
                        Anchor(1440, 2880, 4032, 2304))))));

        var persistDir = Rec(0x1772, 0, 0, I32((2 << 20) | 1), I32(0), I32(document.Length));
        var userEditOffset = document.Length + slide.Length + persistDir.Length;
        var userEdit = Rec(0x0FF5, 0, 0, I32(256), U16(0), new byte[] { 0, 3 }, I32(0), I32(document.Length + slide.Length), I32(1), I32(3), U16(1), U16(0));
        var docStream = Pad(Concat(document, slide, persistDir, userEdit));

        var currentUser = Pad(Rec(0x0FF6, 0, 0, I32(20), I32(encrypted ? unchecked((int)0xF3D1C4DF) : unchecked((int)0xE391C05F)), I32(userEditOffset),
            U16(6), U16(0x03F4), new byte[] { 3, 0, 0, 0 }, "Auspex"u8.ToArray(), I32(8)));
        var pictures = Pad(Rec(0xF01E, 0x6E0, 0, new byte[16], new byte[] { 0xFF }, Png));

        return Cfb(("PowerPoint Document", docStream), ("Current User", currentUser), ("Pictures", pictures));
    }

    static byte[] Rec(int type, int inst, int ver, params byte[][] parts)
    {
        var body = Concat(parts);
        return Concat(U16((inst << 4) | ver), U16(type), I32(body.Length), body);
    }

    static byte[] Opt(params (int Pid, int Value)[] props)
    {
        var body = Concat(props.Select(p => Concat(U16(p.Pid), I32(p.Value))).ToArray());
        return Rec(0xF00B, props.Length, 3, body);
    }

    /// <summary>OfficeArtClientAnchor as PowerPoint writes it: top, left, right, bottom in master units.</summary>
    static byte[] Anchor(int top, int left, int right, int bottom) => Rec(0xF010, 0, 0, U16(top), U16(left), U16(right), U16(bottom));

    static byte[] FontName(string name) { var b = new byte[64]; Encoding.Unicode.GetBytes(name).CopyTo(b, 0); return b; }
    static byte[] Utf16(string s) => Encoding.Unicode.GetBytes(s);
    static byte[] U16(int v) => [(byte)v, (byte)(v >> 8)];
    static byte[] I32(int v) => BitConverter.GetBytes(v);
    static byte[] Concat(params byte[][] parts) { var ms = new MemoryStream(); foreach (var p in parts) ms.Write(p); return ms.ToArray(); }
    /// <summary>Zeros after the last record up to 4096 bytes, so the stream lives in the FAT chain rather than the mini stream.</summary>
    static byte[] Pad(byte[] b) => b.Length >= 4096 ? b : Concat(b, new byte[4096 - b.Length]);

    /// <summary>A compound file with 512-byte sectors: FAT in sector 0, the directory in sector 1, then the streams.</summary>
    static byte[] Cfb(params (string Name, byte[] Data)[] streams)
    {
        const int S = 512;
        var fat = new List<int> { unchecked((int)0xFFFFFFFD), -2 }; // sector 0 = FAT, sector 1 = directory
        var dir = new byte[S];
        var body = new MemoryStream();
        void Entry(int i, string name, byte type, int start, int size, int child)
        {
            var off = i * 128;
            var n = Encoding.Unicode.GetBytes(name);
            n.CopyTo(dir, off);
            U16(n.Length + 2).CopyTo(dir, off + 0x40);
            dir[off + 0x42] = type; dir[off + 0x43] = 1;
            I32(-1).CopyTo(dir, off + 0x44); I32(i + 1 < streams.Length + 1 && type == 2 ? i + 1 : -1).CopyTo(dir, off + 0x48); I32(child).CopyTo(dir, off + 0x4C);
            I32(start).CopyTo(dir, off + 0x74); I32(size).CopyTo(dir, off + 0x78);
        }
        Entry(0, "Root Entry", 5, -2, 0, 1);
        for (var i = 0; i < streams.Length; i++)
        {
            var (name, data) = streams[i];
            var start = fat.Count;
            var sectors = (data.Length + S - 1) / S;
            for (var s = 0; s < sectors; s++) fat.Add(s + 1 < sectors ? start + s + 1 : -2);
            body.Write(data);
            body.Write(new byte[sectors * S - data.Length]);
            Entry(i + 1, name, 2, start, data.Length, -1);
        }
        var header = new byte[S];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(header, 0);
        U16(0x3E).CopyTo(header, 0x18); U16(3).CopyTo(header, 0x1A); U16(0xFFFE).CopyTo(header, 0x1C);
        U16(9).CopyTo(header, 0x1E); U16(6).CopyTo(header, 0x20);
        I32(1).CopyTo(header, 0x2C);      // one FAT sector
        I32(1).CopyTo(header, 0x30);      // directory starts at sector 1
        I32(4096).CopyTo(header, 0x38);   // mini stream cutoff
        I32(-2).CopyTo(header, 0x3C); I32(0).CopyTo(header, 0x40); I32(-2).CopyTo(header, 0x44); I32(0).CopyTo(header, 0x48);
        I32(0).CopyTo(header, 0x4C);      // DIFAT[0] = FAT sector 0
        for (var i = 1; i < 109; i++) I32(-1).CopyTo(header, 0x4C + i * 4);
        var fatSector = new byte[S];
        for (var i = 0; i < S / 4; i++) I32(i < fat.Count ? fat[i] : -1).CopyTo(fatSector, i * 4);
        return Concat(header, fatSector, dir, body.ToArray());
    }
}
