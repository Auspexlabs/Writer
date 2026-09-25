using System.Buffers.Binary;
using System.Text;
using Writer.Core;

namespace Writer.Formats.Compat;

public static partial class PptReader
{
    // ---- records: uint16 verInst (recVer low 4 bits, 0xF = container; recInstance high 12), uint16 type, uint32 length ----

    readonly record struct Rec(int Type, int Inst, int Ver, int Start, int Len) { public int End => Start + Len; }

    static IEnumerable<Rec> Records(byte[] b, int start, int end)
    {
        end = Math.Min(end, b.Length);
        for (var off = Math.Max(0, start); off + 8 <= end;)
        {
            var r = RecAt(b, off, end);
            yield return r;
            off += 8 + r.Len;
        }
    }

    static Rec RecAt(byte[] b, int off, int end = int.MaxValue)
    {
        end = Math.Min(end, b.Length);
        int vi = U16(b, off), type = U16(b, off + 2);
        var len = (int)Math.Min(U32(b, off + 4), (uint)Math.Max(0, end - off - 8));
        return new Rec(type, vi >> 4, vi & 0xF, off + 8, len);
    }

    static IEnumerable<Rec> Children(byte[] b, Rec r) => r.Ver == 0xF ? Records(b, r.Start, r.End) : [];

    static Rec? Child(byte[] b, Rec r, int type)
    {
        foreach (var c in Children(b, r)) if (c.Type == type) return c;
        return null;
    }

    static int TypeAt(byte[] b, int off) => off >= 0 && off + 8 <= b.Length ? U16(b, off + 2) : -1;

    static int U16(byte[] b, int off) => off >= 0 && off + 2 <= b.Length ? BinaryPrimitives.ReadUInt16LittleEndian(b.AsSpan(off)) : 0;
    static int I16(byte[] b, int off) => off >= 0 && off + 2 <= b.Length ? BinaryPrimitives.ReadInt16LittleEndian(b.AsSpan(off)) : 0;
    static int I32(byte[] b, int off) => off >= 0 && off + 4 <= b.Length ? BinaryPrimitives.ReadInt32LittleEndian(b.AsSpan(off)) : 0;
    static uint U32(byte[] b, int off) => off >= 0 && off + 4 <= b.Length ? BinaryPrimitives.ReadUInt32LittleEndian(b.AsSpan(off)) : 0;
    static byte B(byte[] b, int off) => off >= 0 && off < b.Length ? b[off] : (byte)0;

    // ---- shapes ----

    /// <summary>Group space → slide space: slide = O + (child - C) × S, all in master units.</summary>
    readonly record struct Xf(double Ox, double Oy, double Sx, double Sy, double Cx, double Cy)
    {
        public static readonly Xf Identity = new(0, 0, 1, 1, 0, 0);
        public (double X, double Y, double W, double H) Map(double l, double t, double r, double b) => (Ox + (l - Cx) * Sx, Oy + (t - Cy) * Sy, (r - l) * Sx, (b - t) * Sy);
    }

    sealed class Found { public ShapeModel Shape = new(); public int Placeholder, TextType = 4; public string Raw = ""; public bool Background; }

    /// <summary>Every shape of a slide's (or notes page's) drawing, groups flattened, in z-order.</summary>
    static List<Found> Shapes(byte[] b, Rec slideRec, ListEntry entry, SlideCtx sctx)
    {
        var found = new List<Found>();
        if (Child(b, slideRec, RtDrawing) is { } dr && Child(b, dr, DgContainer) is { } dg)
            foreach (var c in Children(b, dg))
            {
                if (c.Type == SpgrContainer) Group(b, c, Xf.Identity, entry, sctx, found);
                else if (c.Type == SpContainer) Visit(b, c, Xf.Identity, entry, sctx, found); // the slide's background shape sits here
            }
        return found;
    }

    static void Group(byte[] b, Rec group, Xf xf, ListEntry entry, SlideCtx sctx, List<Found> found)
    {
        var first = true;
        foreach (var c in Children(b, group))
        {
            if (c.Type == SpgrContainer) { Group(b, c, xf, entry, sctx, found); continue; }
            if (c.Type != SpContainer) continue;
            if (first)
            {
                first = false;
                // the group's own shape: its anchor (in the parent's space) against its child bounds is the transform for the rest
                if (Child(b, c, Spgr) is { } sg && sg.Len >= 16)
                {
                    int cl = I32(b, sg.Start), ct = I32(b, sg.Start + 4), cr = I32(b, sg.Start + 8), cb = I32(b, sg.Start + 12);
                    if (Anchor(b, c, xf) is { } a) xf = new Xf(a.X, a.Y, cr > cl ? a.W / (cr - cl) : 1, cb > ct ? a.H / (cb - ct) : 1, cl, ct);
                    continue;
                }
            }
            Visit(b, c, xf, entry, sctx, found);
        }
    }

    static void Visit(byte[] b, Rec c, Xf xf, ListEntry entry, SlideCtx sctx, List<Found> found)
    {
        try { if (Shape(b, c, xf, entry, sctx) is { } f) found.Add(f); }
        catch (Exception e) when (e is not WriterException) { sctx.Ctx.Warnings.Add($"a shape could not be read ({e.Message})"); }
    }

    static (double X, double Y, double W, double H)? Anchor(byte[] b, Rec sp, Xf xf)
    {
        if (Child(b, sp, ClientAnchor) is { } ca)
        {
            // top, left, right, bottom in master units; int16 ×4 or int32 ×4
            if (ca.Len >= 16) { int t = I32(b, ca.Start), l = I32(b, ca.Start + 4), r = I32(b, ca.Start + 8), bt = I32(b, ca.Start + 12); return (l, t, r - l, bt - t); }
            if (ca.Len >= 8) { int t = I16(b, ca.Start), l = I16(b, ca.Start + 2), r = I16(b, ca.Start + 4), bt = I16(b, ca.Start + 6); return (l, t, r - l, bt - t); }
        }
        if (Child(b, sp, ChildAnchor) is { } ch && ch.Len >= 16)
            return xf.Map(I32(b, ch.Start), I32(b, ch.Start + 4), I32(b, ch.Start + 8), I32(b, ch.Start + 12));
        return null;
    }

    static Found? Shape(byte[] b, Rec c, Xf xf, ListEntry entry, SlideCtx sctx)
    {
        if (Child(b, c, Sp) is not { } sp || sp.Len < 8) return null;
        var flags = U32(b, sp.Start + 4);
        if ((flags & 0xC) != 0) return null; // deleted, or the patriarch
        var opt = OptProps(b, c);
        var f = new Found();
        var s = f.Shape;
        if ((flags & 0x400) != 0) // the slide's own background
        {
            f.Background = true;
            if (opt.TryGetValue(0x0181, out var bg) && (!opt.TryGetValue(0x0180, out var fillType) || fillType == 0)) s.Fill = sctx.OptColor(bg); // solid fills only
            return f;
        }
        var a = Anchor(b, c, xf) ?? (0, 0, 0, 0);
        (s.X, s.Y, s.W, s.H) = (Cm(a.X), Cm(a.Y), Cm(a.W), Cm(a.H));
        if (Child(b, c, ClientData) is { } cd && Child(b, cd, RtPlaceholder) is { } ph && ph.Len >= 5) f.Placeholder = B(b, ph.Start + 4);

        if (opt.TryGetValue(0x0104, out var pib))
        {
            if (Picture(sctx.Ctx, (int)pib) is not { } img) return null;
            s.Kind = ShapeKind.Image;
            s.Image = img;
            return f;
        }

        TextGroup? g = null;
        if (Child(b, c, ClientTextbox) is { } tb)
            foreach (var r in Children(b, tb))
                switch (r.Type)
                {
                    case RtOutlineTextRef when r.Len >= 4: { var i = I32(b, r.Start); if (i >= 0 && i < entry.Texts.Count) g = entry.Texts[i]; break; }
                    case RtTextHeader: g = new TextGroup { Type = I32(b, r.Start) }; break;
                    case RtTextChars or RtTextBytes when g is not null: g.Text = Text(b, r); break;
                    case RtStyleTextProp when g is not null: g.Style = r; break;
                    case RtMasterTextProp when g is not null: g.Master = r; break;
                }
        double? firstSize = null;
        if (g is not null)
        {
            f.TextType = g.Type;
            f.Raw = g.Text.Replace('\r', '\n').Replace('\v', '\n');
            s.Paragraphs = Paragraphs(b, g, sctx, out firstSize);
        }
        s.SizePt = firstSize;
        s.IsTitle = f.Placeholder is 1 or 3 or 12 or 15 or 17 || f.TextType is 0 or 6;

        var type = sp.Inst;
        var connector = type == 20 || (flags & 0x100) != 0;
        var primitive = type is not (75 or 202) && !connector; // an autoshape is filled and outlined unless told otherwise; text boxes and pictures are not
        var filled = opt.TryGetValue(0x01BF, out var ff) && (ff & 0x100000) != 0 ? (ff & 0x10) != 0 : primitive;
        var lined = opt.TryGetValue(0x01FF, out var lf) && (lf & 0x80000) != 0 ? (lf & 0x08) != 0 : primitive || connector;
        if (filled) s.Fill = opt.TryGetValue(0x0181, out var fc) ? sctx.OptColor(fc) : sctx.SchemeColor(4);
        if (lined) s.Line = opt.TryGetValue(0x01C0, out var lc) ? sctx.OptColor(lc) : sctx.SchemeColor(1);
        s.Geometry = Geometry(type);
        if (s.Paragraphs.All(p => p.Html.Length == 0) && s.Fill is null && s.Line is null) return null;
        return f;
    }

    /// <summary>The shape's OfficeArt properties by id (flag bits stripped); complex data is not needed here.</summary>
    static Dictionary<int, uint> OptProps(byte[] b, Rec sp)
    {
        var d = new Dictionary<int, uint>();
        foreach (var o in Children(b, sp))
            if (o.Type is Opt or 0xF121 or 0xF122)
                for (int i = 0, p = o.Start; i < o.Inst && p + 6 <= o.End; i++, p += 6)
                    d[U16(b, p) & 0x3FFF] = U32(b, p + 2);
        return d;
    }

    static string? Geometry(int shapeType) => shapeType switch
    {
        1 => "rect", 2 => "roundRect", 3 => "ellipse", 4 => "diamond", 5 => "triangle", 6 => "rtTriangle", 7 => "parallelogram",
        8 => "trapezoid", 9 => "hexagon", 10 => "octagon", 11 => "plus", 12 => "star5", 13 => "rightArrow", 15 => "homePlate",
        16 => "cube", 20 => "line", 22 => "can", 23 => "donut", 32 => "straightConnector1", 56 => "pentagon", 66 => "leftArrow",
        67 => "downArrow", 68 => "upArrow", 69 => "leftRightArrow", 70 => "upDownArrow", 96 => "smileyFace", _ => null,
    };

    // ---- text ----

    static readonly Encoding Cp1252 = Ansi();

    static Encoding Ansi()
    {
        try { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); return Encoding.GetEncoding(1252); }
        catch (Exception e) when (e is ArgumentException or NotSupportedException) { return Encoding.Latin1; }
    }

    static string Text(byte[] b, Rec r) => r.Type == RtTextBytes ? Cp1252.GetString(b, r.Start, r.Len) : Encoding.Unicode.GetString(b, r.Start, r.Len & ~1);

    sealed record ParaSpec(int Count, int Indent, bool? Bullet, string? Align);
    sealed record RunSpec(int Count, bool Bold, bool Italic, bool Underline, double? Size, string? Color, string? Font);

    /// <summary>The text split at '\r' into paragraphs, each with its runs from the StyleTextPropAtom; a broken atom leaves plain paragraphs.</summary>
    static List<Block> Paragraphs(byte[] b, TextGroup g, SlideCtx sctx, out double? firstSize)
    {
        var text = g.Text;
        var paras = new List<ParaSpec>(); var runs = new List<RunSpec>();
        if (g.Style is { } st) ParseStyle(b, st, text.Length + 1, paras, runs, sctx);
        if (paras.Count == 0 && g.Master is { } m) // MasterTextPropAtom: (count, indent level) pairs when the text keeps the master's styles
            for (var p = m.Start; p + 6 <= m.End && I32(b, p) > 0; p += 6) paras.Add(new ParaSpec(I32(b, p), Math.Clamp(U16(b, p + 4), 0, 8), null, null));
        firstSize = runs.FirstOrDefault(r => r.Size is not null)?.Size;
        var defaultBullet = g.Type is 1 or 5 or 7 or 8; // body placeholders are bulleted by their master
        var blocks = new List<Block>();
        int pos = 0, pi = 0, pCovered = 0, ri = 0, rStart = 0;
        while (true)
        {
            var end = text.IndexOf('\r', pos); if (end < 0) end = text.Length;
            while (pi < paras.Count && pCovered + paras[pi].Count <= pos) { pCovered += paras[pi].Count; pi++; }
            var ps = pi < paras.Count ? paras[pi] : null;
            var sb = new StringBuilder();
            for (var i = pos; i < end;)
            {
                while (ri < runs.Count && rStart + runs[ri].Count <= i) { rStart += runs[ri].Count; ri++; }
                var rs = ri < runs.Count ? runs[ri] : null;
                var to = rs is null ? end : Math.Min(end, rStart + rs.Count);
                var seg = text[i..to];
                sb.Append(rs is null ? Inline.Run(seg) : Inline.Run(seg, rs.Bold, rs.Italic, rs.Underline, color: rs.Color, sizePt: rs.Size, font: rs.Font));
                i = to;
            }
            var bullet = end > pos && (ps?.Bullet ?? defaultBullet);
            blocks.Add(new Block { Kind = BlockKind.Paragraph, Html = sb.ToString(), List = bullet ? "bullet" : null, Level = ps?.Indent ?? 0, Align = ps?.Align });
            if (end >= text.Length) break;
            pos = end + 1;
        }
        if (blocks.Count > 1 && text.EndsWith('\r') && blocks[^1].Html.Length == 0) blocks.RemoveAt(blocks.Count - 1); // a trailing '\r' closes the last paragraph
        return blocks;
    }

    /// <summary>StyleTextPropAtom: paragraph runs (count, indent, mask, fields) then character runs (count, mask, fields), each covering n characters.</summary>
    static void ParseStyle(byte[] b, Rec r, int n, List<ParaSpec> paras, List<RunSpec> runs, SlideCtx sctx)
    {
        int p = r.Start, end = r.End, covered = 0;
        while (covered < n && p + 10 <= end)
        {
            int count = I32(b, p), indent = U16(b, p + 4); var mask = U32(b, p + 6); p += 10;
            if (count <= 0) break;
            bool? bullet = null; string? align = null;
            if ((mask & 0xF) != 0) { if ((mask & 0x1) != 0) bullet = (U16(b, p) & 1) != 0; p += 2; }
            if ((mask & 0x80) != 0) p += 2;   // bulletChar
            if ((mask & 0x10) != 0) p += 2;   // bulletFontRef
            if ((mask & 0x40) != 0) p += 2;   // bulletSize
            if ((mask & 0x20) != 0) p += 4;   // bulletColor
            if ((mask & 0x800) != 0) { align = U16(b, p) switch { 0 => "left", 1 => "center", 2 => "right", 3 => "justify", _ => null }; p += 2; }
            if ((mask & 0x1000) != 0) p += 2; // lineSpacing
            if ((mask & 0x2000) != 0) p += 2; // spaceBefore
            if ((mask & 0x4000) != 0) p += 2; // spaceAfter
            if ((mask & 0x100) != 0) p += 2;  // leftMargin
            if ((mask & 0x400) != 0) p += 2;  // indent
            if ((mask & 0x8000) != 0) p += 2; // defaultTabSize
            if ((mask & 0x100000) != 0) p += 2 + U16(b, p) * 4; // tabStops
            if ((mask & 0x10000) != 0) p += 2;  // fontAlign
            if ((mask & 0xE0000) != 0) p += 2;  // wrapFlags
            if ((mask & 0x200000) != 0) p += 2; // textDirection
            if (p > end) { paras.Clear(); return; }
            paras.Add(new ParaSpec(count, Math.Clamp(indent, 0, 8), bullet, align));
            covered += count;
        }
        covered = 0;
        while (covered < n && p + 8 <= end)
        {
            var count = I32(b, p); var mask = U32(b, p + 4); p += 8;
            if (count <= 0) break;
            bool bold = false, italic = false, underline = false; double? size = null; string? color = null, font = null;
            if ((mask & 0xFFFF) != 0)
            {
                var fs = U16(b, p); p += 2;
                bold = (mask & 1) != 0 && (fs & 1) != 0; italic = (mask & 2) != 0 && (fs & 2) != 0; underline = (mask & 4) != 0 && (fs & 4) != 0;
            }
            if ((mask & 0x10000) != 0) { font = sctx.Ctx.Font(U16(b, p)); p += 2; }
            if ((mask & 0x200000) != 0) p += 2; // oldEAFontRef
            if ((mask & 0x400000) != 0) p += 2; // ansiFontRef
            if ((mask & 0x800000) != 0) p += 2; // symbolFontRef
            if ((mask & 0x20000) != 0) { var v = U16(b, p); if (v > 0) size = v; p += 2; }
            if ((mask & 0x40000) != 0) { color = sctx.Color(B(b, p), B(b, p + 1), B(b, p + 2), B(b, p + 3)); p += 4; }
            if ((mask & 0x80000) != 0) p += 2;  // position
            if (p > end) { runs.Clear(); return; }
            runs.Add(new RunSpec(count, bold, italic, underline, size, color, font));
            covered += count;
        }
    }

    // ---- pictures ----

    /// <summary>The BLIP store entries in order: where each picture sits in the Pictures stream, or the BLIP record itself when it is inline.</summary>
    static List<(int FoDelay, Rec? Inline)> Bses(byte[] doc, Rec docRec)
    {
        var list = new List<(int, Rec?)>();
        if (Child(doc, docRec, RtDrawingGroup) is { } dg && Child(doc, dg, DggContainer) is { } dgg && Child(doc, dgg, BStore) is { } bs)
            foreach (var e in Children(doc, bs))
            {
                if (e.Type != Bse) { list.Add((-1, e.Type is >= 0xF018 and <= 0xF117 ? e : null)); continue; }
                if (e.Len < 36) { list.Add((-1, null)); continue; }
                int foDelay = I32(doc, e.Start + 28), cbName = B(doc, e.Start + 33);
                list.Add((foDelay, e.Len > 36 + cbName + 8 ? RecAt(doc, e.Start + 36 + cbName, e.End) : null));
            }
        return list;
    }

    static byte[]? Picture(Ctx ctx, int pib)
    {
        var missing = $"picture {pib} is missing from the file";
        if (pib <= 0 || pib > ctx.Bses.Count) { ctx.Warnings.Add(missing); return null; }
        var (foDelay, inline) = ctx.Bses[pib - 1];
        if (inline is { } r) return Blip(ctx.Doc, r, ctx.Warnings);
        if (foDelay < 0 || foDelay + 8 > ctx.Pictures.Length) { ctx.Warnings.Add(missing); return null; }
        return Blip(ctx.Pictures, RecAt(ctx.Pictures, foDelay), ctx.Warnings);
    }

    /// <summary>The picture bytes of a BLIP record: PNG and JPEG as they are, a DIB wrapped as BMP; metafiles are dropped.</summary>
    static byte[]? Blip(byte[] b, Rec r, List<string> warnings)
    {
        if (r.Type is not (0xF01D or 0xF01E or 0xF01F))
        {
            warnings.Add(r.Type is 0xF01A or 0xF01B or 0xF01C ? "picture in WMF/EMF format dropped" : $"picture of unsupported type 0x{r.Type:X4} dropped");
            return null;
        }
        // 16 uid bytes, 16 more for the two-uid instances, then a tag byte, then the data
        var skip = r.Type switch { 0xF01E => r.Inst == 0x6E1, 0xF01D => r.Inst is 0x46B or 0x6E3, _ => r.Inst == 0x7A9 } ? 33 : 17;
        if (!Looks(b, r.Start + skip, r.Type) && Looks(b, r.Start + 50 - skip, r.Type)) skip = 50 - skip;
        if (r.Len <= skip) { warnings.Add("empty picture dropped"); return null; }
        var data = b.AsSpan(r.Start + skip, r.Len - skip).ToArray();
        return r.Type == 0xF01F ? Bmp(data) : data;
    }

    static bool Looks(byte[] b, int off, int type) => type switch
    {
        0xF01E => B(b, off) == 0x89 && B(b, off + 1) == 0x50,
        0xF01D => B(b, off) == 0xFF && B(b, off + 1) == 0xD8,
        _ => U32(b, off) is 12 or 40 or 52 or 56 or 108 or 124,
    };

    static byte[] Bmp(byte[] dib)
    {
        var biSize = (int)U32(dib, 0);
        int bpp = U16(dib, 14), compression = (int)U32(dib, 16), clrUsed = (int)U32(dib, 32);
        var table = bpp is > 0 and <= 8 ? (clrUsed > 0 ? clrUsed : 1 << bpp) * 4 : compression == 3 && biSize == 40 ? 12 : 0;
        var bmp = new byte[14 + dib.Length];
        bmp[0] = (byte)'B'; bmp[1] = (byte)'M';
        BinaryPrimitives.WriteUInt32LittleEndian(bmp.AsSpan(2), (uint)bmp.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(bmp.AsSpan(10), (uint)(14 + biSize + table));
        dib.CopyTo(bmp, 14);
        return bmp;
    }
}
