using System.Buffers.Binary;
using System.Text;
using Writer.Core;
using Writer.Formats.Compat;

namespace Writer.Tests;

public class XlsReaderTests
{
    [Fact]
    public void ReadsCellsStylesFormulasMergesAndSizes()
    {
        var warnings = new List<string>();
        var sheets = XlsReader.Read(XlsFixture.Build(), warnings);

        Assert.Equal(["Data", "第二页"], sheets.Select(s => s.Name));
        Assert.Empty(warnings);

        var data = sheets[0];
        CellModel At(int row, int col) => data.Cells.Single(c => c.Row == row && c.Col == col);
        Assert.Equal(1.0, At(1, 1).Value);
        Assert.Equal(2.0, At(2, 1).Value);
        Assert.Equal(3.0, At(3, 1).Value);
        Assert.Equal("SUM(A1:A2)", At(3, 1).Formula);
        Assert.Equal("Hello", At(1, 2).Value);
        Assert.True(At(1, 2).Bold);
        Assert.False(At(2, 2).Bold);
        Assert.Equal("第二页文字", At(2, 2).Value);
        Assert.Equal(true, At(1, 3).Value);
        Assert.Equal(new DateTime(2025, 1, 1), At(2, 3).Value);
        Assert.Equal("m/d/yyyy", At(2, 3).Format);
        Assert.Equal("middle", At(1, 1).VAlign);
        Assert.Equal("IF(A1>1,\"yes\",B1&\"x\")", At(3, 3).Formula);
        Assert.Equal("no", At(3, 3).Value);
        Assert.Equal("ROUND(SUM($A$1:B$2)*2,0)", At(4, 3).Formula);
        Assert.Equal("A1*2", At(1, 4).Formula);
        Assert.Equal("A2*2", At(2, 4).Formula);
        Assert.Equal(4.0, At(2, 4).Value);
        Assert.Equal(["A5:B6"], data.Merges);
        Assert.Equal(20, data.ColWidths[1]);
        Assert.Equal(30, data.RowHeights[1]);

        Assert.Equal("第二页文字", sheets[1].Cells.Single().Value);
    }

    /// <summary>Writes the sample workbook next to the other fixtures (and into the checked-in Fixtures folder when run from the repo) so end-to-end tests can open a real .xls file.</summary>
    [Fact]
    public void WritesTheFixtureFile()
    {
        var bytes = XlsFixture.Build();
        var output = Path.Combine(AppContext.BaseDirectory, "Fixtures", "compat", "sample.xls");
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        File.WriteAllBytes(output, bytes);
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            if (!File.Exists(Path.Combine(dir.FullName, "Writer.Tests.csproj"))) continue;
            var source = Path.Combine(dir.FullName, "Fixtures", "compat", "sample.xls");
            if (!File.Exists(source) || !File.ReadAllBytes(source).AsSpan().SequenceEqual(bytes))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(source)!);
                File.WriteAllBytes(source, bytes);
            }
            break;
        }
        Assert.Equal(bytes, File.ReadAllBytes(output));
        Assert.Equal(2, XlsReader.Read(File.ReadAllBytes(output), []).Count);
    }

    [Fact]
    public void Biff5AndEncryptedWorkbooksAskForXlsx()
    {
        var biff5 = Assert.Throws<WriterException>(() => XlsReader.Read(XlsFixture.Cfb("Book", XlsFixture.Workbook()), []));
        Assert.Equal(ErrorCode.FormatError, biff5.Code);
        Assert.Contains(".xlsx", biff5.Hint);

        var locked = Assert.Throws<WriterException>(() => XlsReader.Read(XlsFixture.Build(encrypted: true), []));
        Assert.Equal(ErrorCode.FormatError, locked.Code);
        Assert.Contains("password", locked.Message);
    }
}

/// <summary>A tiny BIFF8 workbook inside a 512-byte-sector compound file, built byte by byte (no Excel, no NPOI on this machine).</summary>
public static class XlsFixture
{
    public static byte[] Build(bool encrypted = false) => Cfb("Workbook", Workbook(encrypted));

    public static byte[] Workbook(bool encrypted = false)
    {
        var sheet1 = Sheet1();
        var sheet2 = Sheet2();
        var globalsLength = Globals(0, 0, encrypted).Length;
        var globals = Globals(globalsLength, globalsLength + sheet1.Length, encrypted);
        return [.. globals, .. sheet1, .. sheet2];
    }

    static byte[] Globals(int sheet1Pos, int sheet2Pos, bool encrypted) =>
    [
        .. Rec(0x0809, U16(0x0600), U16(0x0005), U16(0x0DBB), U16(0x07CC), I32(0), I32(0)),
        .. (encrypted ? Rec(0x002F, U16(1), U16(1), U16(1), new byte[48]) : []),
        .. Rec(0x0042, U16(1200)),
        .. Rec(0x0022, U16(0)),
        .. Font(bold: false),
        .. Font(bold: true),
        .. Xf(font: 0, fmt: 0),
        .. Xf(font: 1, fmt: 0),
        .. Xf(font: 0, fmt: 14),
        .. Rec(0x00FC, I32(2), I32(2), Str16("Hello"), Str16W("第二页文字")),
        .. Rec(0x0085, I32(sheet1Pos), U16(0), Str8("Data")),
        .. Rec(0x0085, I32(sheet2Pos), U16(0), Str8W("第二页")),
        .. Rec(0x000A),
    ];

    static byte[] Sheet1() =>
    [
        .. Rec(0x0809, U16(0x0600), U16(0x0010), U16(0x0DBB), U16(0x07CC), I32(0), I32(0)),
        .. Rec(0x007D, U16(0), U16(0), U16(20 * 256), U16(0), U16(0), U16(0)),
        .. Rec(0x0208, U16(0), U16(0), U16(3), U16(30 * 20), U16(0), U16(0), U16(0x0140), U16(0)),
        .. Rec(0x0203, U16(0), U16(0), U16(0), F64(1.0)),
        .. Rec(0x027E, U16(1), U16(0), U16(0), I32((2 << 2) | 2)),
        .. Rec(0x0006, U16(2), U16(0), U16(0), F64(3.0), U16(0), I32(0), U16(13),
            [0x25, 0x00, 0x00, 0x01, 0x00, 0x00, 0xC0, 0x00, 0xC0, 0x19, 0x10, 0x00, 0x00]),
        .. Rec(0x00FD, U16(0), U16(1), U16(1), I32(0)),
        .. Rec(0x00FD, U16(1), U16(1), U16(0), I32(1)),
        .. Rec(0x0205, U16(0), U16(2), U16(0), [1, 0]),
        .. Rec(0x0203, U16(1), U16(2), U16(2), F64(45658.0)),
        // C3 = IF(A1>1,"yes",B1&"x") with a string result held by the STRING record that follows
        .. Rec(0x0006, U16(2), U16(2), U16(0), [0, 0, 0, 0, 0, 0, 0xFF, 0xFF], U16(0), I32(0), U16(37),
            [0x24, 0, 0, 0, 0xC0, 0x1E, 1, 0, 0x0D, 0x19, 0x02, 8, 0, 0x17, 3, 0, (byte)'y', (byte)'e', (byte)'s', 0x19, 0x08, 11, 0,
             0x24, 0, 0, 1, 0xC0, 0x17, 1, 0, (byte)'x', 0x08, 0x22, 3, 1, 0]),
        .. Rec(0x0207, Str16("no")),
        // C4 = ROUND(SUM($A$1:B$2)*2,0)
        .. Rec(0x0006, U16(3), U16(2), U16(0), F64(6.0), U16(0), I32(0), U16(23),
            [0x25, 0, 0, 1, 0, 0, 0, 1, 0x40, 0x19, 0x10, 0, 0, 0x1E, 2, 0, 0x05, 0x1E, 0, 0, 0x21, 27, 0]),
        // D1:D2 share one formula, A1*2, stored once in SHRFMLA with a RefN three columns to the left
        .. Rec(0x0006, U16(0), U16(3), U16(0), F64(2.0), U16(8), I32(0), U16(5), [0x01, 0, 0, 3, 0]),
        .. Rec(0x04BC, U16(0), U16(1), [3, 3, 0, 2], U16(9), [0x2C, 0, 0, 0xFD, 0xC0, 0x1E, 2, 0, 0x05]),
        .. Rec(0x0006, U16(1), U16(3), U16(0), F64(4.0), U16(8), I32(0), U16(5), [0x01, 0, 0, 3, 0]),
        .. Rec(0x00E5, U16(1), U16(4), U16(5), U16(0), U16(1)),
        .. Rec(0x000A),
    ];

    static byte[] Sheet2() =>
    [
        .. Rec(0x0809, U16(0x0600), U16(0x0010), U16(0x0DBB), U16(0x07CC), I32(0), I32(0)),
        .. Rec(0x00FD, U16(0), U16(0), U16(0), I32(1)),
        .. Rec(0x000A),
    ];

    static byte[] Font(bool bold) =>
        Rec(0x0031, U16(200), U16(0), U16(0x7FFF), U16(bold ? 700 : 400), U16(0), [0, 0, 0, 0], Str8("Arial"));

    static byte[] Xf(int font, int fmt) =>
        Rec(0x00E0, U16(font), U16(fmt), U16(1), [0x10, 0, 0, 0], I32(0), I32(0), U16(64 | (65 << 7)));

    static byte[] Rec(int id, params byte[][] parts)
    {
        var payload = parts.SelectMany(p => p).ToArray();
        return [.. U16(id), .. U16(payload.Length), .. payload];
    }

    static byte[] U16(int v) { var b = new byte[2]; BinaryPrimitives.WriteUInt16LittleEndian(b, (ushort)v); return b; }
    static byte[] I32(int v) { var b = new byte[4]; BinaryPrimitives.WriteInt32LittleEndian(b, v); return b; }
    static byte[] F64(double v) { var b = new byte[8]; BinaryPrimitives.WriteDoubleLittleEndian(b, v); return b; }
    static byte[] Str8(string s) => [(byte)s.Length, 0, .. Encoding.Latin1.GetBytes(s)];
    static byte[] Str8W(string s) => [(byte)s.Length, 1, .. Encoding.Unicode.GetBytes(s)];
    static byte[] Str16(string s) => [.. U16(s.Length), 0, .. Encoding.Latin1.GetBytes(s)];
    static byte[] Str16W(string s) => [.. U16(s.Length), 1, .. Encoding.Unicode.GetBytes(s)];

    /// <summary>A compound file with one stream in regular sectors: header, one FAT sector, one directory sector, then the stream padded to at least 4096 bytes.</summary>
    public static byte[] Cfb(string streamName, byte[] stream)
    {
        var padded = new byte[Math.Max(4096, (stream.Length + 511) / 512 * 512)];
        stream.CopyTo(padded, 0);
        var sectors = padded.Length / 512;
        var file = new byte[512 * (3 + sectors)];
        var h = file.AsSpan();
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(h);
        BinaryPrimitives.WriteUInt16LittleEndian(h[0x18..], 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(h[0x1A..], 3);
        BinaryPrimitives.WriteUInt16LittleEndian(h[0x1C..], 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(h[0x1E..], 9);
        BinaryPrimitives.WriteUInt16LittleEndian(h[0x20..], 6);
        BinaryPrimitives.WriteInt32LittleEndian(h[0x2C..], 1);
        BinaryPrimitives.WriteInt32LittleEndian(h[0x30..], 1);
        BinaryPrimitives.WriteInt32LittleEndian(h[0x38..], 4096);
        BinaryPrimitives.WriteInt32LittleEndian(h[0x3C..], -2);
        BinaryPrimitives.WriteInt32LittleEndian(h[0x44..], -2);
        for (var i = 0; i < 109; i++) BinaryPrimitives.WriteInt32LittleEndian(h[(0x4C + i * 4)..], i == 0 ? 0 : -1);

        var fat = file.AsSpan(512, 512);
        for (var i = 0; i < 128; i++) BinaryPrimitives.WriteInt32LittleEndian(fat[(i * 4)..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(fat, -3);
        BinaryPrimitives.WriteInt32LittleEndian(fat[4..], -2);
        for (var i = 0; i < sectors; i++) BinaryPrimitives.WriteInt32LittleEndian(fat[((2 + i) * 4)..], i == sectors - 1 ? -2 : 3 + i);

        var dir = file.AsSpan(1024, 512);
        Entry(dir, "Root Entry", 5, child: 1, start: -2, size: 0);
        Entry(dir[128..], streamName, 2, child: -1, start: 2, size: padded.Length);
        padded.CopyTo(file, 1536);
        return file;
    }

    static void Entry(Span<byte> e, string name, byte type, int child, int start, int size)
    {
        var utf16 = Encoding.Unicode.GetBytes(name);
        utf16.CopyTo(e);
        BinaryPrimitives.WriteUInt16LittleEndian(e[0x40..], (ushort)(utf16.Length + 2));
        e[0x42] = type;
        e[0x43] = 1;
        BinaryPrimitives.WriteInt32LittleEndian(e[0x44..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(e[0x48..], -1);
        BinaryPrimitives.WriteInt32LittleEndian(e[0x4C..], child);
        BinaryPrimitives.WriteInt32LittleEndian(e[0x74..], start);
        BinaryPrimitives.WriteInt32LittleEndian(e[0x78..], size);
    }
}
