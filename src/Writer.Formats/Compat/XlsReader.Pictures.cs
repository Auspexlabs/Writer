using System.Buffers.Binary;

namespace Writer.Formats.Compat;

public static partial class XlsReader
{
    /// <summary>The sheet's pictures: each drawing shape whose pib points at a PNG, JPEG or DIB in the workbook's picture store, anchored at its top-left cell.</summary>
    static void ReadPictures(SheetState st, Book b, List<string> warnings)
    {
        if (st.Drawing.Length == 0 || b.DrawingGroup.Length == 0) return;
        b.Blips ??= Blips(b.DrawingGroup.ToArray());
        var d = st.Drawing.ToArray();
        var n = 0;
        Walk(d, 0, d.Length, (type, _, start, end) =>
        {
            if (type != 0xF004) return;
            int pib = 0, col1 = 0, row1 = 0, col2 = 0, row2 = 0, dx1 = 0, dy1 = 0, dx2 = 0, dy2 = 0;
            var anchored = false;
            Walk(d, start, end, (t, inst, s, e) =>
            {
                if (t == 0xF00B)
                    for (var i = 0; i < inst && s + i * 6 + 6 <= e; i++)
                        if ((U16(d, s + i * 6) & 0x3FFF) == 0x0104) pib = I32(d, s + i * 6 + 2);
                if (t == 0xF010 && e - s >= 18)
                {
                    anchored = true;
                    col1 = U16(d, s + 2); dx1 = U16(d, s + 4); row1 = U16(d, s + 6); dy1 = U16(d, s + 8);
                    col2 = U16(d, s + 10); dx2 = U16(d, s + 12); row2 = U16(d, s + 14); dy2 = U16(d, s + 16);
                }
            });
            if (pib <= 0 || !anchored) return;
            n++;
            var blip = pib <= b.Blips.Count ? b.Blips[pib - 1] : (null, "missing");
            if (blip.Bytes is null) { warnings.Add($"Sheet {st.Sheet.Name}: picture {n} at {A1(row1, col1)} is a {blip.Kind} image and was dropped"); return; }
            var img = new SheetImage { Bytes = blip.Bytes, Col = col1 + 1, Row = row1 + 1 };
            var w = Span(col1, dx1, col2, dx2, 1024, c => (st.Sheet.ColWidths.TryGetValue(c + 1, out var cw) ? cw : st.DefaultColWidth) * 7 + 5);
            var h = Span(row1, dy1, row2, dy2, 256, r => (st.Sheet.RowHeights.TryGetValue(r + 1, out var rh) ? rh : st.DefaultRowHeight) * 96 / 72);
            if (w > 1 && h > 1) { img.WidthCm = Math.Round(w * 2.54 / 96, 2); img.HeightCm = Math.Round(h * 2.54 / 96, 2); }
            st.Sheet.Images.Add(img);
        });
    }

    /// <summary>Pixels between two anchor points; each fraction is in 1/unit of its cell's size.</summary>
    static double Span(int a, int fa, int z, int fz, double unit, Func<int, double> size)
    {
        if (z < a || z - a > 1000) return 0;
        if (a == z) return (fz - fa) / unit * size(a);
        var px = (1 - fa / unit) * size(a) + fz / unit * size(z);
        for (var i = a + 1; i < z; i++) px += size(i);
        return px;
    }

    /// <summary>Calls visit for every Office Art atom, walking into containers.</summary>
    static void Walk(byte[] d, int start, int end, Action<int, int, int, int> visit)
    {
        for (var p = start; p + 8 <= end;)
        {
            var verInst = U16(d, p);
            var type = U16(d, p + 2);
            var len = (int)Math.Min(BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(p + 4)), (uint)(end - p - 8));
            if ((verInst & 0xF) == 0xF) { visit(type, verInst >> 4, p + 8, p + 8 + len); Walk(d, p + 8, p + 8 + len, visit); }
            else visit(type, verInst >> 4, p + 8, p + 8 + len);
            p += 8 + len;
        }
    }

    /// <summary>The workbook's picture store in pib order: the image bytes (PNG, JPEG, DIB as BMP), or null with the kind that was dropped.</summary>
    static List<(byte[]? Bytes, string Kind)> Blips(byte[] g)
    {
        var list = new List<(byte[]?, string)>();
        Walk(g, 0, g.Length, (type, _, start, end) =>
        {
            if (type != 0xF007) return;
            if (end - start < 36) { list.Add((null, "empty")); return; }
            var p = start + 36 + g[start + 33];
            if (p + 8 > end) { list.Add((null, "external")); return; }
            var blipType = U16(g, p + 2);
            var inst = U16(g, p) >> 4;
            var data = p + 8 + 16 + ((inst & 1) != 0 ? 16 : 0);
            switch (blipType)
            {
                case 0xF01E or 0xF01D: list.Add((data + 1 <= end ? g[(data + 1)..end] : null, blipType == 0xF01E ? "PNG" : "JPEG")); break;
                case 0xF01F: list.Add((data + 1 + 40 <= end ? Bmp(g[(data + 1)..end]) : null, "DIB")); break;
                case 0xF01A: list.Add((null, "EMF")); break;
                case 0xF01B: list.Add((null, "WMF")); break;
                case 0xF01C: list.Add((null, "PICT")); break;
                default: list.Add((null, "TIFF or unknown")); break;
            }
        });
        return list;
    }

    /// <summary>A BMP file from a DIB: the 14-byte file header in front of the BITMAPINFOHEADER-led bytes.</summary>
    static byte[] Bmp(byte[] dib)
    {
        var biSize = I32(dib, 0);
        int bitCount = U16(dib, 14), compression = I32(dib, 16), clrUsed = I32(dib, 32);
        var colours = clrUsed != 0 ? clrUsed : bitCount <= 8 ? 1 << bitCount : 0;
        var offBits = 14 + biSize + colours * 4 + (compression == 3 && biSize == 40 ? 12 : 0);
        var bmp = new byte[14 + dib.Length];
        bmp[0] = (byte)'B'; bmp[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(2), bmp.Length);
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(10), offBits);
        dib.CopyTo(bmp, 14);
        return bmp;
    }
}
