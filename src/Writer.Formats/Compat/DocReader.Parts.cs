using System.Text;

namespace Writer.Formats.Compat;

public static partial class DocReader
{
    /// <summary>Character properties, resolved from the paragraph style downwards.</summary>
    struct Chp
    {
        public bool Bold, Italic, Underline, Strike, Hidden, Spec;
        /// <summary>Size in half-points.</summary>
        public int Hps = 20;
        public int Ico, Highlight;
        /// <summary>COLORREF (R | G&lt;&lt;8 | B&lt;&lt;16 | fAuto&lt;&lt;24), or -1.</summary>
        public int Cv = -1;
        public int Ftc0 = -1, Ftc1 = -1, PicFc = -1, Symbol = -1, SymFtc = -1;
        public Chp() { }
    }

    /// <summary>Paragraph properties: what the block model needs, plus the table row definition a row-end mark carries.</summary>
    struct Pap
    {
        public int Istd, Jc, Itap, Ilvl, Ilfo;
        public int OutLvl = 9;
        public bool InTable, Ttp, PageBreakBefore;
        public byte[]? TDef;
        public List<(int First, byte[] Shd)>? Shd;
        public Pap() { }
    }

    sealed class Sty
    {
        public string Name = "";
        public int Sti = -1, Base = 0xFFF;
        public int PapxOff = -1, PapxLen, ChpxOff = -1, ChpxLen;
        /// <summary>Heading level 1-9 when the style is one, else 0.</summary>
        public int Level;
        public string? Hint;
        public Pap? Pap;
        public Chp? Chp;
    }

    static readonly string[] Ico =
    [
        "", "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00", "FFFFFF",
        "000080", "008080", "008000", "800080", "800000", "808000", "808080", "C0C0C0",
    ];

    static string? Palette(int ico) => ico is >= 1 and <= 16 ? Ico[ico] : null;

    /// <summary>The run's color as RRGGBB; black and automatic count as none.</summary>
    static string? Color(in Chp chp)
    {
        var c = chp.Cv >= 0 && chp.Cv >> 24 != 0xFF ? Inline.Bgr((uint)chp.Cv) : chp.Ico is >= 2 and <= 16 ? Ico[chp.Ico] : null;
        return c == "000000" ? null : c;
    }

    static bool Toggle(byte b, bool current) => b switch { 0 => false, 1 => true, 0x81 => !current, _ => current };

    sealed partial class Doc
    {
        Sty?[] _styles = [];
        readonly Sty _plain = new();
        string[] _fonts = [];
        readonly Dictionary<int, byte[]> _lists = [];
        readonly List<(int Lsid, int[]? Nfc)> _lfos = [];

        // ---- styles ----

        void ReadStyles()
        {
            int fc = I32(_wd, 0xA2), lcb = I32(_wd, 0xA6);
            if (fc < 0 || lcb < 4) return;
            var end = (int)Math.Min((long)fc + lcb, _tbl.Length);
            var cbStshi = U16(_tbl, fc);
            var stshi = fc + 2;
            int cstd = U16(_tbl, stshi), cbBase = U16(_tbl, stshi + 2);
            if (cstd <= 0 || cstd > 4096 || cbBase < 8) return;
            _styles = new Sty?[cstd];
            var pos = stshi + cbStshi;
            for (var i = 0; i < cstd && pos + 2 <= end; i++)
            {
                var cbStd = U16(_tbl, pos);
                pos += 2;
                if (cbStd == 0) continue;
                var std = pos;
                pos += cbStd;
                if (std + cbStd > end) break;
                var s = new Sty { Sti = U16(_tbl, std) & 0xFFF };
                var w2 = U16(_tbl, std + 2);
                int stk = w2 & 0xF, cupx = U16(_tbl, std + 4) & 0xF;
                s.Base = w2 >> 4;
                var cch = U16(_tbl, std + cbBase);
                if (std + cbBase + 2 + 2 * cch <= end) s.Name = Encoding.Unicode.GetString(_tbl, std + cbBase + 2, 2 * cch);
                var p = cbBase + 2 + 2 * cch + 2;
                if ((p & 1) == 1) p++;
                for (var u = 0; u < cupx && std + p + 2 <= std + cbStd; u++)
                {
                    var cb = U16(_tbl, std + p);
                    p += 2;
                    var off = std + p;
                    var len = Math.Min(cb, std + cbStd - off);
                    if (len < 0) break;
                    if (u == 0 && stk is 1 or 3) { if (len >= 2) { s.PapxOff = off + 2; s.PapxLen = len - 2; } }
                    else if (u == 0 && stk == 2 || u == 1 && stk is 1 or 3) { s.ChpxOff = off; s.ChpxLen = len; }
                    p += cb;
                    if ((p & 1) == 1) p++;
                }
                if (s.Sti is >= 1 and <= 9) s.Level = s.Sti;
                else if (s.Sti == 62) s.Hint = "Title";
                else if (s.Sti == 74) s.Hint = "Subtitle";
                else s.Level = NameLevel(s.Name);
                _styles[i] = s;
            }
        }

        static int NameLevel(string name)
        {
            var n = name.Trim();
            foreach (var prefix in (ReadOnlySpan<string>)["heading", "标题", "標題"])
                if (n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    var rest = n[prefix.Length..].Trim();
                    if (rest.Length == 1 && rest[0] is >= '1' and <= '9') return rest[0] - '0';
                }
            return 0;
        }

        Sty Style(int istd) => istd >= 0 && istd < _styles.Length && _styles[istd] is { } s ? s : _plain;

        Pap StylePap(int istd, int depth = 0)
        {
            var s = Style(istd);
            if (s.Pap is { } done) return done;
            var pap = s.Base != 0xFFF && s.Base != istd && depth < 12 ? StylePap(s.Base, depth + 1) : new Pap();
            pap.Istd = istd;
            pap.TDef = null;
            pap.Shd = null;
            if (s.PapxOff >= 0) ApplyParaSprms(ref pap, _tbl, s.PapxOff, s.PapxLen);
            s.Pap = pap;
            return pap;
        }

        Chp StyleChp(int istd, int depth = 0)
        {
            var s = Style(istd);
            if (s.Chp is { } done) return done;
            var chp = s.Base != 0xFFF && s.Base != istd && depth < 12 ? StyleChp(s.Base, depth + 1) : new Chp();
            if (s.ChpxOff >= 0) ApplyCharSprms(ref chp, _tbl, s.ChpxOff, s.ChpxLen, depth + 1);
            s.Chp = chp;
            return chp;
        }

        /// <summary>A character style applied on top of the current properties: its base chain first, then its own sprms.</summary>
        void ApplyCharStyle(ref Chp chp, int istd, int depth)
        {
            if (depth > 12) return;
            var s = Style(istd);
            if (s == _plain) return;
            if (s.Base != 0xFFF && s.Base != istd) ApplyCharStyle(ref chp, s.Base, depth + 1);
            if (s.ChpxOff >= 0) ApplyCharSprms(ref chp, _tbl, s.ChpxOff, s.ChpxLen, depth + 1);
        }

        void ApplyCharSprms(ref Chp chp, byte[] b, int start, int len, int depth)
        {
            var pos = start;
            var end = (int)Math.Min((long)start + len, b.Length);
            while (Next(b, ref pos, end, out var op, out var off, out _))
            {
                switch (op)
                {
                    case 0x0835: chp.Bold = Toggle(b[off], chp.Bold); break;
                    case 0x0836: chp.Italic = Toggle(b[off], chp.Italic); break;
                    case 0x0837 or 0x2A53: chp.Strike = Toggle(b[off], chp.Strike); break;
                    case 0x0855: chp.Spec = Toggle(b[off], chp.Spec); break;
                    case 0x0818: chp.Hidden = Toggle(b[off], chp.Hidden); break;
                    case 0x2A3E: chp.Underline = b[off] != 0; break;
                    case 0x2A42: chp.Ico = b[off]; chp.Cv = -1; break;
                    case 0x6870: chp.Cv = I32(b, off); break;
                    case 0x4A43: chp.Hps = U16(b, off); break;
                    case 0x4A4F: chp.Ftc0 = U16(b, off); break;
                    case 0x4A50: chp.Ftc1 = U16(b, off); break;
                    case 0x2A0C: chp.Highlight = b[off]; break;
                    case 0x6A03: chp.PicFc = I32(b, off); break;
                    case 0x6A09: chp.SymFtc = U16(b, off); chp.Symbol = U16(b, off + 2); break;
                    case 0x4A30: if (depth <= 12) ApplyCharStyle(ref chp, U16(b, off), depth + 1); break;
                }
            }
        }

        void ApplyParaSprms(ref Pap pap, byte[] b, int start, int len)
        {
            var pos = start;
            var end = (int)Math.Min((long)start + len, b.Length);
            while (Next(b, ref pos, end, out var op, out var off, out var n))
            {
                switch (op)
                {
                    case 0x2403 or 0x2461: pap.Jc = b[off]; break;
                    case 0x2416: pap.InTable = b[off] != 0; break;
                    case 0x2417: pap.Ttp = b[off] != 0; break;
                    case 0x6649: pap.Itap = I32(b, off); break;
                    case 0x664A: pap.Itap += I32(b, off); break;
                    case 0x260A: pap.Ilvl = b[off]; break;
                    case 0x460B: pap.Ilfo = U16(b, off); break;
                    case 0x2640: pap.OutLvl = b[off]; break;
                    case 0x2407: pap.PageBreakBefore = b[off] != 0; break;
                    case 0xD608 or 0xD606: if (n > 2) pap.TDef = b[(off + 2)..(off + n)]; break;
                    case 0xD609 or 0xD612 or 0xD616 or 0xD60C:
                        if (n > 1) (pap.Shd ??= []).Add((op == 0xD616 ? 22 : op == 0xD60C ? 44 : 0, b[(off + 1)..(off + n)]));
                        break;
                }
            }
        }

        // ---- fonts ----

        void ReadFonts()
        {
            int fc = I32(_wd, 0x112), lcb = I32(_wd, 0x116);
            if (fc < 0 || lcb < 4) return;
            var end = (int)Math.Min((long)fc + lcb, _tbl.Length);
            var count = U16(_tbl, fc);
            var pos = fc + 4;
            if (count == 0xFFFF) { count = U16(_tbl, fc + 2); pos = fc + 6; }
            var fonts = new List<string>();
            for (var i = 0; i < count && pos < end; i++)
            {
                int cb = _tbl[pos], nameAt = pos + 40, stop = Math.Min(pos + 1 + cb, end);
                var sb = new StringBuilder();
                for (var p = nameAt; p + 1 < stop; p += 2)
                {
                    var ch = (char)U16(_tbl, p);
                    if (ch == '\0') break;
                    sb.Append(ch);
                }
                fonts.Add(sb.ToString());
                pos += 1 + cb;
            }
            _fonts = fonts.ToArray();
        }

        string? Font(int ftc) => ftc >= 0 && ftc < _fonts.Length && _fonts[ftc].Length > 0 ? _fonts[ftc] : null;

        // ---- lists ----

        void ReadLists()
        {
            int fc = I32(_wd, 0x2E2), lcb = I32(_wd, 0x2E6);
            if (fc >= 0 && lcb >= 2)
            {
                var end = (int)Math.Min((long)fc + lcb, _tbl.Length);
                var cLst = I16(_tbl, fc);
                var pos = fc + 2;
                var lsids = new List<(int Lsid, bool Simple)>();
                for (var i = 0; i < cLst && pos + 28 <= end; i++, pos += 28) lsids.Add((I32(_tbl, pos), (_tbl[pos + 26] & 1) != 0));
                // the LVLs follow the LSTFs; Word places them right after, but the lcb only bounds the LSTFs
                var lvlEnd = _tbl.Length;
                foreach (var (lsid, simple) in lsids)
                {
                    var nfc = new byte[9];
                    for (var l = 0; l < (simple ? 1 : 9) && pos + 28 <= lvlEnd; l++)
                    {
                        nfc[l] = _tbl[pos + 4];
                        pos += 28 + _tbl[pos + 25] + _tbl[pos + 24];
                        pos += 2 + 2 * U16(_tbl, pos);
                    }
                    _lists[lsid] = nfc;
                }
            }
            fc = I32(_wd, 0x2EA); lcb = I32(_wd, 0x2EE);
            if (fc < 0 || lcb < 4) return;
            var cLfo = I32(_tbl, fc);
            var p = fc + 4;
            var clfolvl = new List<int>();
            for (var i = 0; i < cLfo && p + 16 <= _tbl.Length; i++, p += 16)
            {
                _lfos.Add((I32(_tbl, p), null));
                clfolvl.Add(_tbl[p + 12]);
            }
            for (var i = 0; i < _lfos.Count && p + 4 <= _tbl.Length; i++)
            {
                p += 4; // the cp the list first appears at
                for (var k = 0; k < clfolvl[i] && p + 8 <= _tbl.Length; k++)
                {
                    int ilvl = _tbl[p + 4] & 0xF;
                    var formatting = (_tbl[p + 4] & 0x20) != 0;
                    p += 8;
                    if (!formatting || p + 28 > _tbl.Length) continue;
                    var ovr = _lfos[i].Nfc ?? Enumerable.Repeat(-1, 9).ToArray();
                    ovr[ilvl] = _tbl[p + 4];
                    _lfos[i] = (_lfos[i].Lsid, ovr);
                    p += 28 + _tbl[p + 25] + _tbl[p + 24];
                    p += 2 + 2 * U16(_tbl, p);
                }
            }
        }

        /// <summary>bullet, number, or null when the level shows no marker.</summary>
        string? ListKind(int ilfo, int ilvl)
        {
            ilvl = Math.Clamp(ilvl, 0, 8);
            if (ilfo <= 0 || ilfo > _lfos.Count) return "number";
            var (lsid, ovr) = _lfos[ilfo - 1];
            var nfc = ovr is not null && ovr[ilvl] >= 0 ? ovr[ilvl] : _lists.TryGetValue(lsid, out var lvls) ? lvls[ilvl] : 0;
            return nfc switch { 23 => "bullet", 255 => null, _ => "number" };
        }

        // ---- table cell shading ----

        static string? Shd(byte[] shd, int k)
        {
            if (10 * k + 10 > shd.Length) return null;
            int fore = I32(shd, 10 * k), back = I32(shd, 10 * k + 4), ipat = U16(shd, 10 * k + 8);
            var c = ipat == 1 ? fore : back;
            return ipat is 0 or 1 && c >> 24 != 0xFF ? Inline.Bgr((uint)c) : null;
        }

        static string? Shd80(byte[] shd, int k)
        {
            if (2 * k + 2 > shd.Length) return null;
            var v = U16(shd, 2 * k);
            int fore = v & 0x1F, back = v >> 5 & 0x1F, ipat = v >> 10;
            return ipat == 1 ? Palette(fore) : ipat == 0 ? Palette(back) : null;
        }

        // ---- pictures ----

        /// <summary>The inline picture whose PICF sits at <paramref name="fc"/> in the Data stream, or null (with a warning).</summary>
        Block? Picture(int fc)
        {
            if (fc < 0 || fc + 0x44 > _data.Length) return null;
            var lcb = I32(_data, fc);
            int cbHeader = U16(_data, fc + 4), mm = I16(_data, fc + 6);
            int dxaGoal = I16(_data, fc + 0x1C), dyaGoal = I16(_data, fc + 0x1E), mx = U16(_data, fc + 0x20), my = U16(_data, fc + 0x22);
            var pos = fc + Math.Max(cbHeader, 0x44);
            if (mm == 0x66 && pos < _data.Length) pos += 1 + _data[pos];
            var end = (int)Math.Min((long)fc + Math.Max(lcb, 0), _data.Length);
            var bytes = FirstBlip(_data, pos, end, 0);
            if (bytes is null) { Warn("picture in an unsupported format dropped"); return null; }
            double? w = dxaGoal > 0 && mx > 0 ? dxaGoal * mx / 1000.0 / 567.0 : null;
            double? h = dyaGoal > 0 && my > 0 ? dyaGoal * my / 1000.0 / 567.0 : null;
            return Block.Picture(bytes, w, h);
        }

        byte[]? FirstBlip(byte[] b, int pos, int end, int depth)
        {
            while (pos + 8 <= end && depth < 8)
            {
                int verInst = U16(b, pos), type = U16(b, pos + 2);
                var len = (int)Math.Min(U32(b, pos + 4), end - pos - 8);
                var body = pos + 8;
                if ((verInst & 0xF) == 0xF)
                {
                    var found = FirstBlip(b, body, body + len, depth + 1);
                    if (found is not null) return found;
                }
                else if (type == 0xF007 && body + 36 <= end)
                {
                    var found = FirstBlip(b, body + 36 + b[body + 33], body + len, depth + 1);
                    if (found is not null) return found;
                }
                else if (type is >= 0xF01A and <= 0xF02A)
                {
                    var instance = verInst >> 4;
                    switch (type)
                    {
                        case 0xF01E: return Blip(b, body, len, instance == 0x6E1 ? 33 : 17, false);
                        case 0xF01D: return Blip(b, body, len, instance is 0x46B or 0x6E3 ? 33 : 17, false);
                        case 0xF01F: return Blip(b, body, len, instance == 0x7A9 ? 33 : 17, true);
                        case 0xF01A or 0xF01B or 0xF01C: Warn("picture in WMF/EMF format dropped"); return null;
                        default: Warn("picture in an unsupported format dropped"); return null;
                    }
                }
                pos = body + len;
            }
            return null;
        }

        static byte[]? Blip(byte[] b, int body, int len, int skip, bool dib)
        {
            if (len <= skip) return null;
            var payload = b[(body + skip)..(body + len)];
            if (!dib)
            {
                // some writers put TIFF or other data in a PNG/JPEG record; the model takes PNG, JPEG, GIF and BMP only
                var ok = payload.Length > 8 && (payload[0] == 0x89 && payload[1] == 'P' && payload[2] == 'N' && payload[3] == 'G'
                    || payload[0] == 0xFF && payload[1] == 0xD8 || payload[0] == 'G' && payload[1] == 'I' && payload[2] == 'F');
                return ok ? payload : null;
            }
            // a DIB is a BMP without its 14-byte file header
            uint hdr = U32(payload, 0), comp = U32(payload, 16), clrUsed = U32(payload, 32);
            if (hdr is not (12 or 40 or 52 or 56 or 108 or 124)) return null;
            int bpp = U16(payload, 14);
            var colors = bpp <= 8 ? (clrUsed != 0 ? clrUsed : 1u << bpp) * 4 : comp == 3 && hdr == 40 ? 12u : 0;
            var bmp = new byte[14 + payload.Length];
            bmp[0] = (byte)'B'; bmp[1] = (byte)'M';
            BitConverter.TryWriteBytes(bmp.AsSpan(2), (uint)bmp.Length);
            BitConverter.TryWriteBytes(bmp.AsSpan(10), 14 + hdr + colors);
            payload.CopyTo(bmp, 14);
            return bmp;
        }

        // ---- floating pictures: anchored by PlcfSpaMom, drawn by the shapes in DggInfo, stored in its BStore ----

        /// <summary>BStore entries in order: where each blip record sits (in the table stream when embedded, else in Data at foDelay).</summary>
        readonly List<(bool InTable, int Pos, int End)> _bstore = [];
        readonly Dictionary<int, byte[]?> _bstoreBytes = [];
        readonly Dictionary<int, int> _spidPib = [];
        /// <summary>Anchor cp → shape id and size in twips.</summary>
        readonly Dictionary<int, (int Spid, int W, int H)> _spa = [];
        readonly List<Block> _floating = [];

        sealed class Group { public int Spid = -1; }

        void ReadDrawings()
        {
            int fc = I32(_wd, 0x22A), lcb = I32(_wd, 0x22E);
            if (fc >= 0 && lcb > 0) WalkDrawing(fc, (int)Math.Min((long)fc + lcb, _tbl.Length), 0, null);
            fc = I32(_wd, 0x1DA); lcb = I32(_wd, 0x1DE);
            var n = (lcb - 4) / 30;
            for (var i = 0; i < n && fc >= 0; i++)
            {
                var spa = fc + 4 * (n + 1) + 26 * i;
                _spa[I32(_tbl, fc + 4 * i)] = (I32(_tbl, spa), I32(_tbl, spa + 12) - I32(_tbl, spa + 4), I32(_tbl, spa + 16) - I32(_tbl, spa + 8));
            }
        }

        void WalkDrawing(int pos, int end, int depth, Group? group)
        {
            while (pos + 8 <= end && depth < 12)
            {
                int verInst = U16(_tbl, pos), type = U16(_tbl, pos + 2);
                if (type is < 0xF000 or > 0xF1FF)
                {
                    // OfficeArtWordDrawing: a one-byte dgglbl (0 main document, 1 headers) sits before each drawing container
                    if (depth == 0) { pos++; continue; }
                    break;
                }
                var len = (int)Math.Min(U32(_tbl, pos + 4), end - pos - 8);
                var body = pos + 8;
                if (type == 0xF004)
                {
                    // a shape: its id and, for pictures, the BStore index (property pib)
                    int spid = -1, pib = 0;
                    for (var p = body; p + 8 <= body + len;)
                    {
                        int t = U16(_tbl, p + 2), l = (int)Math.Min(U32(_tbl, p + 4), body + len - p - 8);
                        if (t == 0xF00A) spid = I32(_tbl, p + 8);
                        else if (t == 0xF00B)
                            for (int k = 0, c = U16(_tbl, p) >> 4; k < c && p + 8 + 6 * k + 6 <= p + 8 + l; k++)
                                if ((U16(_tbl, p + 8 + 6 * k) & 0x3FFF) == 0x0104) pib = I32(_tbl, p + 8 + 6 * k + 2);
                        p += 8 + l;
                    }
                    if (spid >= 0 && pib > 0) _spidPib[spid] = pib;
                    if (group is not null)
                    {
                        if (group.Spid < 0) group.Spid = spid;
                        else if (pib > 0 && !_spidPib.ContainsKey(group.Spid)) _spidPib[group.Spid] = pib; // a picture inside a group: the anchor names the group
                    }
                }
                else if (type == 0xF007 && body + 36 <= end)
                {
                    // not embedded: the blip sits in Word's delay stream, which is the WordDocument stream, at foDelay
                    int size = I32(_tbl, body + 20), foDelay = I32(_tbl, body + 28);
                    if (len > 36) _bstore.Add((true, body + 36 + _tbl[body + 33], body + len));
                    else _bstore.Add((false, foDelay, foDelay >= 0 ? (int)Math.Min((long)foDelay + (size > 0 ? size : _wd.Length), _wd.Length) : 0));
                }
                else if ((verInst & 0xF) == 0xF) WalkDrawing(body, body + len, depth + 1, type == 0xF003 ? new Group() : group);
                pos = body + len;
            }
        }

        byte[]? Bstore(int pib)
        {
            if (pib < 1 || pib > _bstore.Count) return null;
            if (_bstoreBytes.TryGetValue(pib, out var cached)) return cached;
            var (inTable, pos, end) = _bstore[pib - 1];
            var bytes = inTable ? FirstBlip(_tbl, pos, end, 0) : pos >= 0 && pos < end ? FirstBlip(_wd, pos, end, 0) : null;
            return _bstoreBytes[pib] = bytes;
        }

        /// <summary>The picture anchored at <paramref name="cp"/> (a 0x08 character) becomes a block after its paragraph.</summary>
        void Floating(int cp)
        {
            if (_spa.TryGetValue(cp, out var a) && _spidPib.TryGetValue(a.Spid, out var pib) && Bstore(pib) is { } bytes)
                _floating.Add(Block.Picture(bytes, a.W > 0 ? a.W / 567.0 : null, a.H > 0 ? a.H / 567.0 : null));
            else Warn("drawn shapes dropped");
        }
    }
}
