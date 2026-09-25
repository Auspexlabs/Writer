using System.Text;
using Writer.Core;

namespace Writer.Formats.Compat;

/// <summary>Word 97-2003 binary documents (.doc, .dot, and Kingsoft WPS .wps/.wpt, which use the same layout) → blocks.
/// Keeps paragraphs, headings, runs (bold/italic/underline/strike, size, font, color, highlight), alignment, lists,
/// tables, inline pictures, page breaks and hyperlinks; drops headers, footers, footnotes, comments, text boxes,
/// hidden text, field codes and floating drawings, one warning line each.</summary>
public static partial class DocReader
{
    public static List<Block> Read(byte[] file, List<string> warnings)
    {
        var cfb = new Cfb(file);
        var wd = cfb.Stream("WordDocument")
            ?? throw new WriterException(ErrorCode.FormatError, "Not a Word document", "The file has no WordDocument stream; it may be another kind of Office 97-2003 file.");
        if (wd.Length < 0x200 || U16(wd, 0) != 0xA5EC)
            throw new WriterException(ErrorCode.FormatError, U16(wd, 0) == 0xA5DC ? "Word 6/95 documents are not supported" : "Not a Word 97-2003 document",
                "Open it in Word or WPS and save it as .docx, then open it again.");
        var flags = U16(wd, 0x0A);
        if ((flags & 0x0100) != 0)
            throw new WriterException(ErrorCode.FormatError, "This document is password-protected", "Remove the password in Word, then open it again.");
        var table = cfb.Stream((flags & 0x0200) != 0 ? "1Table" : "0Table") ?? cfb.Stream("1Table") ?? cfb.Stream("0Table")
            ?? throw new WriterException(ErrorCode.FormatError, "Damaged Word document", "The file has no table stream.");
        var counts = new Dictionary<string, int>();
        try
        {
            return new Doc(wd, table, cfb.Stream("Data") ?? [], counts).Run();
        }
        catch (WriterException) { throw; }
        catch (Exception e)
        {
            throw new WriterException(ErrorCode.FormatError, "Damaged Word document", $"The file could not be read ({e.GetType().Name}). Open and re-save it in Word or WPS, then try again.");
        }
        finally
        {
            foreach (var (m, n) in counts) warnings.Add(n > 1 ? $"{m} (×{n})" : m);
        }
    }

    // lenient little-endian readers: out of range reads as 0, so a damaged table degrades instead of throwing
    static int U16(byte[] b, long off) => off >= 0 && off + 2 <= b.Length ? b[off] | b[off + 1] << 8 : 0;
    static int I16(byte[] b, long off) => (short)U16(b, off);
    static int I32(byte[] b, long off) => off >= 0 && off + 4 <= b.Length ? b[off] | b[off + 1] << 8 | b[off + 2] << 16 | b[off + 3] << 24 : 0;
    static uint U32(byte[] b, long off) => (uint)I32(b, off);

    /// <summary>The next sprm of a grpprl: opcode, where its operand starts and how long it is (for variable-size
    /// sprms the length byte(s) are included). False at the end or when the grpprl is malformed.</summary>
    static bool Next(byte[] b, ref int pos, int end, out int op, out int off, out int len)
    {
        op = 0; off = 0; len = 0;
        if (pos + 2 > end) return false;
        op = U16(b, pos); off = pos + 2;
        switch (op >> 13)
        {
            case 0 or 1: len = 1; break;
            case 2 or 4 or 5: len = 2; break;
            case 3: len = 4; break;
            case 7: len = 3; break;
            default:
                if (op is 0xD608 or 0xD606) len = U16(b, off) + 1;
                else
                {
                    if (off >= end) return false;
                    if (op == 0xC615 && b[off] == 255) return false;
                    len = b[off] + 1;
                }
                break;
        }
        if (off + len > end) return false;
        pos = off + len;
        return true;
    }

    record struct Piece(int CpStart, int CpEnd, int Fc, bool Unicode);

    /// <summary>One run of a CHPX or PAPX FKP: the fc range it covers and where its grpprl sits in WordDocument.</summary>
    record struct Fkp(int Fc, int End, int Off, int Len, int Istd);

    sealed class Seg
    {
        public readonly StringBuilder Text = new();
        public int Chpx = -1;
        public string? Link;
    }

    sealed class Field
    {
        public bool InCode = true;
        public readonly StringBuilder Code = new();
        public string? Link;
    }

    /// <summary>How one stretch of text looks; equal props merge into one run.</summary>
    record struct Props(bool Bold, bool Italic, bool Underline, bool Strike, string? Color, double? SizePt, string? Font, string? Highlight, string? Link);

    sealed partial class Doc(byte[] wd, byte[] tbl, byte[] data, Dictionary<string, int> warn)
    {
        readonly byte[] _wd = wd, _tbl = tbl, _data = data;
        readonly Dictionary<string, int> _warn = warn;

        readonly List<Piece> _pieces = [];
        char[] _text = [];
        int _ccpText;
        readonly List<Fkp> _chpx = [], _papx = [];
        int _ci, _pi;
        readonly HashSet<int> _sectionMarks = [];

        readonly List<Block> _blocks = [];
        readonly List<Seg> _segs = [];
        Seg? _cur;
        readonly List<Field> _fields = [];
        List<List<TableCell>> _rows = [];
        List<TableCell> _cells = [];
        readonly List<string> _cellParas = [];
        readonly List<Block> _tablePics = [];

        void Warn(string message) => _warn[message] = _warn.GetValueOrDefault(message) + 1;

        void Try(Action read, string what)
        {
            try { read(); }
            catch (Exception) { Warn($"{what} could not be read; some formatting is lost"); }
        }

        public List<Block> Run()
        {
            ReadText();
            Try(ReadStyles, "the style sheet");
            Try(ReadFonts, "the font table");
            Try(ReadLists, "the list table");
            Try(ReadSections, "the section table");
            Try(ReadDrawings, "the drawings");
            Try(() => ReadFkps(0xFA, _chpx, false), "character formatting");
            Try(() => ReadFkps(0x102, _papx, true), "paragraph formatting");

            var lastFc = 0;
            foreach (var p in _pieces)
            {
                if (p.CpStart >= _ccpText) break;
                var end = Math.Min(p.CpEnd, _ccpText);
                for (var cp = p.CpStart; cp < end; cp++)
                {
                    var fc = p.Unicode ? p.Fc + 2 * (cp - p.CpStart) : p.Fc + (cp - p.CpStart);
                    lastFc = fc;
                    Handle(cp, fc, _text[cp]);
                }
            }
            if (_segs.Count > 0) EndParagraph(lastFc, '\r');
            FlushTable();
            while (_blocks.Count > 0 && _blocks[^1].Kind is BlockKind.PageBreak || _blocks.Count > 0 && _blocks[^1] is { Kind: BlockKind.Paragraph, Html.Length: 0 })
                _blocks.RemoveAt(_blocks.Count - 1);
            StoryWarnings();
            return _blocks;
        }

        // ---- text ----

        void ReadText()
        {
            _ccpText = I32(_wd, 0x4C);
            var total = _ccpText;
            for (var i = 0; i < 7; i++) total += Math.Max(0, I32(_wd, 0x50 + 4 * i));
            ParseClx(I32(_wd, 0x1A2), I32(_wd, 0x1A6));
            if (_pieces.Count == 0)
            {
                // no usable piece table: the text sits between fcMin and fcMac as one piece
                int fcMin = I32(_wd, 0x18), fcMac = I32(_wd, 0x1C);
                if (fcMin >= 0 && fcMac > fcMin && fcMac <= _wd.Length)
                {
                    var uni = total > 0 && fcMac - fcMin >= 2 * total;
                    _pieces.Add(new Piece(0, uni ? (fcMac - fcMin) / 2 : fcMac - fcMin, fcMin, uni));
                }
            }
            if (_pieces.Count == 0) throw new WriterException(ErrorCode.FormatError, "Damaged Word document", "The file's text table is missing.");
            var cpEnd = _pieces[^1].CpEnd;
            if (cpEnd <= 0 || cpEnd > 100_000_000) throw new WriterException(ErrorCode.FormatError, "Damaged Word document", "The file's text table is invalid.");
            _text = new char[cpEnd];
            foreach (var p in _pieces)
                for (int k = 0, n = p.CpEnd - p.CpStart; k < n && p.CpStart + k < cpEnd; k++)
                {
                    if (p.Unicode)
                    {
                        var at = (long)p.Fc + 2 * k;
                        _text[p.CpStart + k] = at + 1 < _wd.Length ? (char)(_wd[at] | _wd[at + 1] << 8) : '\0';
                    }
                    else
                    {
                        var at = (long)p.Fc + k;
                        _text[p.CpStart + k] = at < _wd.Length ? Cp1252(_wd[at]) : '\0';
                    }
                }
            if (_ccpText <= 0 || _ccpText > cpEnd) _ccpText = cpEnd;
        }

        void ParseClx(int fc, int lcb)
        {
            if (fc < 0 || lcb <= 0) return;
            var end = (int)Math.Min((long)fc + lcb, _tbl.Length);
            var pos = fc;
            while (pos < end && _tbl[pos] == 1) pos += 3 + U16(_tbl, pos + 1);
            if (pos >= end || _tbl[pos] != 2) return;
            var plc = pos + 5;
            var n = (I32(_tbl, pos + 1) - 4) / 12;
            if (n <= 0 || (long)plc + 4 * (n + 1) + 8L * n > _tbl.Length) return;
            var last = 0;
            for (var i = 0; i < n; i++)
            {
                int cpS = I32(_tbl, plc + 4 * i), cpE = I32(_tbl, plc + 4 * (i + 1));
                var f = I32(_tbl, plc + 4 * (n + 1) + 8 * i + 2);
                var uni = (f & 0x40000000) == 0;
                if (!uni) f = (f & 0x3FFFFFFF) / 2;
                if (cpE <= cpS || cpS != last || f < 0) { _pieces.Clear(); return; }
                _pieces.Add(new Piece(cpS, cpE, f, uni));
                last = cpE;
            }
        }

        static readonly ushort[] Cp1252High =
        [
            0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
            0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
        ];

        static char Cp1252(byte b) => b is < 0x80 or >= 0xA0 ? (char)b : (char)Cp1252High[b - 0x80];

        void ReadSections()
        {
            int fc = I32(_wd, 0xCA), lcb = I32(_wd, 0xCE);
            var n = (lcb - 4) / 16;
            for (var i = 1; i < n; i++) _sectionMarks.Add(I32(_tbl, fc + 4 * i) - 1);
        }

        void StoryWarnings()
        {
            string[] names = ["footnotes", "headers and footers", "macros", "comments", "endnotes", "text boxes", "header text boxes"];
            var start = _ccpText;
            for (var i = 0; i < names.Length; i++)
            {
                var count = Math.Max(0, I32(_wd, 0x50 + 4 * i));
                for (var cp = start; cp < start + count && cp < _text.Length; cp++)
                    if (_text[cp] > ' ') { Warn($"{names[i]} dropped"); break; }
                start += count;
            }
        }

        // ---- formatting runs ----

        void ReadFkps(int fibOff, List<Fkp> runs, bool papx)
        {
            int fc = I32(_wd, fibOff), lcb = I32(_wd, fibOff + 4);
            var n = (lcb - 4) / 8;
            for (var i = 0; i < n; i++)
            {
                var page = (long)I32(_tbl, fc + 4 * (n + 1) + 4 * i) * 512;
                if (page < 0 || page + 512 > _wd.Length) continue;
                int crun = _wd[page + 511];
                if (4 * (crun + 1) + (papx ? 13 : 1) * crun > 511) continue;
                for (var j = 0; j < crun; j++)
                {
                    int fS = I32(_wd, page + 4 * j), fE = I32(_wd, page + 4 * (j + 1));
                    if (fE <= fS) continue;
                    var pageEnd = (int)page + 512;
                    if (!papx)
                    {
                        int rgb = _wd[page + 4 * (crun + 1) + j];
                        if (rgb == 0) { runs.Add(new Fkp(fS, fE, -1, 0, 0)); continue; }
                        var off = (int)page + rgb * 2;
                        if (off >= pageEnd) continue;
                        runs.Add(new Fkp(fS, fE, off + 1, Math.Min(_wd[off], pageEnd - off - 1), 0));
                    }
                    else
                    {
                        int bo = _wd[page + 4 * (crun + 1) + 13 * j];
                        if (bo == 0) { runs.Add(new Fkp(fS, fE, -1, 0, 0)); continue; }
                        var off = (int)page + bo * 2;
                        if (off + 1 >= pageEnd) continue;
                        int cb = _wd[off], start, len;
                        if (cb == 0) { cb = _wd[off + 1]; start = off + 2; len = 2 * cb; }
                        else { start = off + 1; len = 2 * cb - 1; }
                        len = Math.Min(len, pageEnd - start);
                        runs.Add(len >= 2 ? new Fkp(fS, fE, start + 2, len - 2, U16(_wd, start)) : new Fkp(fS, fE, -1, 0, 0));
                    }
                }
            }
            runs.Sort((a, b) => a.Fc.CompareTo(b.Fc));
        }

        static int RunAt(List<Fkp> runs, int fc, ref int cur)
        {
            if (runs.Count == 0) return -1;
            if (cur >= 0 && cur < runs.Count && fc >= runs[cur].Fc && fc < runs[cur].End) return cur;
            if (cur + 1 < runs.Count && fc >= runs[cur + 1].Fc && fc < runs[cur + 1].End) return ++cur;
            int lo = 0, hi = runs.Count - 1, best = -1;
            while (lo <= hi)
            {
                var mid = (lo + hi) / 2;
                if (runs[mid].Fc <= fc) { best = mid; lo = mid + 1; } else hi = mid - 1;
            }
            if (best >= 0 && fc < runs[best].End) { cur = best; return best; }
            return -1;
        }

        Pap PapAt(int fc)
        {
            var i = RunAt(_papx, fc, ref _pi);
            if (i < 0) return StylePap(0);
            var r = _papx[i];
            var pap = StylePap(r.Istd);
            pap.Istd = r.Istd;
            if (r.Off >= 0) ApplyParaSprms(ref pap, _wd, r.Off, r.Len);
            return pap;
        }

        // ---- the character walk ----

        void Handle(int cp, int fc, char ch)
        {
            switch (ch)
            {
                case '\x13': _fields.Add(new Field()); return;
                case '\x14':
                    if (_fields.Count > 0) { var f = _fields[^1]; f.InCode = false; f.Link = Hyperlink(f.Code.ToString()); }
                    return;
                case '\x15':
                    if (_fields.Count > 0) _fields.RemoveAt(_fields.Count - 1);
                    return;
            }
            for (var i = _fields.Count - 1; i >= 0; i--)
                if (_fields[i].InCode) { _fields[i].Code.Append(ch); return; }
            switch (ch)
            {
                case '\r' or '\x07': EndParagraph(fc, ch); return;
                case '\x0c':
                    if (_sectionMarks.Contains(cp)) { EndParagraph(fc, '\r'); if (cp + 1 < _ccpText) Add(Block.PageBreak()); }
                    else if (_rows.Count > 0 || _cells.Count > 0 || _cellParas.Count > 0) return; // a break inside a table cell has nowhere to go
                    else { var pap = PapAt(fc); if (_segs.Count > 0) EmitParagraph(pap); Add(Block.PageBreak()); }
                    return;
                case '\x0b': ch = '\n'; break;
                case '\x1e': ch = '-'; break;
                case '\x1f': return;
                case '\x08': Floating(cp); return;
            }
            var ci = RunAt(_chpx, fc, ref _ci);
            string? link = null;
            for (var i = _fields.Count - 1; i >= 0 && link is null; i--) link = _fields[i].Link;
            if (_cur is null || _cur.Chpx != ci || _cur.Link != link)
            {
                _cur = new Seg { Chpx = ci, Link = link };
                _segs.Add(_cur);
            }
            _cur.Text.Append(ch);
        }

        static string? Hyperlink(string code)
        {
            var s = code.AsSpan().TrimStart();
            if (!s.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase)) return null;
            s = s["HYPERLINK".Length..];
            string? target = null, anchor = null;
            var expectAnchor = false;
            var i = 0;
            while (i < s.Length)
            {
                while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
                if (i >= s.Length) break;
                string tok;
                if (s[i] == '"')
                {
                    var j = s[(i + 1)..].IndexOf('"');
                    if (j < 0) j = s.Length - i - 1;
                    tok = s.Slice(i + 1, j).ToString();
                    i += j + 2;
                }
                else
                {
                    var j = i;
                    while (j < s.Length && !char.IsWhiteSpace(s[j])) j++;
                    tok = s[i..j].ToString();
                    i = j;
                }
                if (tok.StartsWith('\\'))
                {
                    if (tok.Length >= 2 && (tok[1] == 'l' || tok[1] == 'L')) expectAnchor = true;
                    else if (target is not null) break; // the rest are switches with their own values
                    continue;
                }
                if (expectAnchor) { anchor = tok; expectAnchor = false; }
                else target ??= tok;
            }
            var href = (target ?? "") + (anchor is not null ? "#" + anchor : "");
            return href.Length > 0 ? href : null;
        }

        void Add(Block b)
        {
            if (b.Kind == BlockKind.PageBreak && (_blocks.Count == 0 || _blocks[^1].Kind == BlockKind.PageBreak)) return;
            _blocks.Add(b);
        }

        // ---- paragraphs and tables ----

        void EndParagraph(int fcMark, char mark)
        {
            var pap = PapAt(fcMark);
            var depth = pap.Itap > 0 ? pap.Itap : pap.InTable ? 1 : 0;
            if (depth >= 1) { _tablePics.AddRange(_floating); _floating.Clear(); }
            if (mark == '\x07' && depth <= 1)
            {
                if (pap.Ttp) EndRow(pap);
                else
                {
                    var html = Html(pap, _tablePics, null, null);
                    if (html.Length > 0 || _cellParas.Count > 0) _cellParas.Add(html);
                    _cells.Add(new TableCell(string.Join("<br>", _cellParas)));
                    _cellParas.Clear();
                }
            }
            else if (depth >= 1) _cellParas.Add(Html(pap, _tablePics, null, null));
            else
            {
                FlushTable();
                EmitParagraph(pap);
            }
            _segs.Clear();
            _cur = null;
        }

        void EndRow(Pap pap)
        {
            if (pap.TDef is { Length: >= 1 } t)
            {
                int itc = t[0], tcOff = 1 + 2 * (itc + 1);
                foreach (var (first, shd) in pap.Shd ?? [])
                    for (var k = 0; first + k < _cells.Count; k++)
                    {
                        var fill = shd.Length % 10 == 0 ? Shd(shd, k) : Shd80(shd, k);
                        if (fill is not null) _cells[first + k].Fill = fill;
                    }
                var merged = new List<TableCell>();
                for (var i = 0; i < _cells.Count; i++)
                {
                    var flags = i < itc && tcOff + 20 * i + 2 <= t.Length ? U16(t, tcOff + 20 * i) : 0;
                    if ((flags & 0x0002) != 0 && merged.Count > 0)
                    {
                        merged[^1].ColSpan++;
                        if (_cells[i].Html.Length > 0) merged[^1].Html += (merged[^1].Html.Length > 0 ? "<br>" : "") + _cells[i].Html;
                    }
                    else merged.Add(_cells[i]);
                }
                _cells = merged;
            }
            if (_cells.Count > 0) _rows.Add(_cells);
            _cells = [];
        }

        void FlushTable()
        {
            if (_cells.Count > 0 || _cellParas.Count > 0)
            {
                if (_cellParas.Count > 0) { _cells.Add(new TableCell(string.Join("<br>", _cellParas))); _cellParas.Clear(); }
                _rows.Add(_cells);
                _cells = [];
            }
            if (_rows.Count > 0) Add(Block.Table(_rows));
            _rows = [];
            foreach (var p in _tablePics) Add(p);
            _tablePics.Clear();
        }

        void EmitParagraph(Pap pap)
        {
            var sty = Style(pap.Istd);
            var level = sty.Level;
            if (level == 0 && pap.OutLvl is >= 0 and <= 8) level = pap.OutLvl + 1;
            if (pap.PageBreakBefore) Add(Block.PageBreak());

            string? list = null;
            var listLevel = 0;
            if (level == 0 && pap.Ilfo != 0) { list = ListKind(pap.Ilfo, pap.Ilvl); listLevel = Math.Clamp(pap.Ilvl, 0, 8); }
            else if (level == 0)
            {
                // ponytail: a typed "•<tab>" / "1.<tab>" marker (TextEdit and Pages write lists this way) becomes a real list item
                var plain = new StringBuilder();
                foreach (var s in _segs) plain.Append(s.Text);
                var cut = TypedMarker(plain.ToString(), out var kind);
                if (cut > 0)
                {
                    list = kind;
                    foreach (var s in _segs)
                    {
                        var take = Math.Min(cut, s.Text.Length);
                        s.Text.Remove(0, take);
                        if ((cut -= take) == 0) break;
                    }
                }
            }

            var floating = new List<Block>(_floating);
            _floating.Clear();
            var pics = new List<Block>();
            var look = new Look();
            var html = Html(pap, pics, level > 0 ? StyleChp(pap.Istd) : null, look);
            if (level == 0 && list is null && TypedLevel(look, NormalHps) is var typed && typed > 0)
            {
                // ponytail: a short, all-bold paragraph clearly larger than Normal is a heading typed by hand (TextEdit,
                // Pages and most Chinese office documents never use heading styles); the size and bold move to the heading
                level = typed;
                var baseline = StyleChp(pap.Istd);
                baseline.Bold = true;
                baseline.Hps = look.MinHps;
                pics.Clear();
                html = Html(pap, pics, baseline, null);
            }
            var align = pap.Jc switch { 1 => "center", 2 => "right", >= 3 => "justify", _ => null };
            if (html.Length == 0)
            {
                if (pics.Count == 0 && _blocks.Count > 0 && _blocks[^1].Kind != BlockKind.PageBreak) _blocks.Add(Block.Paragraph(""));
            }
            else if (level > 0)
            {
                var b = Block.Heading(level, html);
                b.Align = align;
                _blocks.Add(b);
            }
            else
            {
                var b = Block.Paragraph(html);
                b.List = list;
                b.Level = list is not null ? listLevel : 0;
                b.Align = align;
                b.Style = sty.Hint;
                _blocks.Add(b);
            }
            foreach (var p in pics) Add(p);
            foreach (var p in floating) Add(p);
        }

        static int TypedMarker(string s, out string kind)
        {
            kind = "";
            var i = s.Length > 0 && s[0] == '\t' ? 1 : 0;
            if (i < s.Length && "•◦▪■●○‣∙·-–".Contains(s[i]))
            {
                if (i + 1 < s.Length && s[i + 1] == '\t') { kind = "bullet"; return i + 2; }
                return 0;
            }
            var j = i;
            while (j < s.Length && j - i < 3 && char.IsAsciiDigit(s[j])) j++;
            if (j == i && i < s.Length && char.IsAsciiLetter(s[i]) && !(i + 1 < s.Length && char.IsAsciiLetter(s[i + 1]))) j = i + 1;
            if (j == i) return 0;
            if (j < s.Length && (s[j] == '.' || s[j] == ')')) j++;
            if (j < s.Length && s[j] == '\t') { kind = "number"; return j + 1; }
            return 0;
        }

        /// <summary>What a paragraph's visible text looks like, for the typed-heading guess.</summary>
        sealed class Look
        {
            public bool AllBold = true, HasBreak;
            public int MinHps = int.MaxValue, Chars;
            public char Last;
        }

        int NormalHps => StyleChp(0).Hps > 0 ? StyleChp(0).Hps : 20;

        static int TypedLevel(Look look, int normalHps)
        {
            if (look.Chars == 0 || look.Chars > 80 || !look.AllBold || look.HasBreak || ".。;；,，:：".Contains(look.Last)) return 0;
            var ratio = look.MinHps / (double)Math.Max(normalHps, 2);
            return ratio >= 1.9 ? 1 : ratio >= 1.4 ? 2 : ratio >= 1.15 ? 3 : 0;
        }

        /// <summary>The paragraph's inline html from its segments; pictures met on the way go to <paramref name="pics"/>.
        /// With a <paramref name="baseline"/> (headings) only formatting that differs from it is written, so the target's
        /// own heading look applies; <paramref name="look"/>, when given, collects what the text looks like.</summary>
        string Html(Pap pap, List<Block> pics, Chp? baseline, Look? look)
        {
            var baseChp = StyleChp(pap.Istd);
            var relative = baseline.HasValue;
            var bl = baseline ?? default;
            var sb = new StringBuilder();
            var buf = new StringBuilder();
            Props cur = default;
            var have = false;
            void Flush()
            {
                if (buf.Length == 0) return;
                sb.Append(Inline.Run(buf.ToString(), cur.Bold, cur.Italic, cur.Underline, cur.Strike, cur.Color, cur.SizePt, cur.Font, cur.Highlight, cur.Link));
                buf.Clear();
            }
            foreach (var seg in _segs)
            {
                var chp = baseChp;
                if (seg.Chpx >= 0 && _chpx[seg.Chpx].Off >= 0) ApplyCharSprms(ref chp, _wd, _chpx[seg.Chpx].Off, _chpx[seg.Chpx].Len, 0);
                if (chp.Hidden) continue;
                foreach (var ch0 in seg.Text.ToString())
                {
                    var ch = ch0;
                    string? font;
                    if (ch < ' ' && ch != '\t' && ch != '\n' || ch == '￼')
                    {
                        if (!chp.Spec) continue;
                        if (ch == '\x01') { var p = Picture(chp.PicFc); if (p is not null) pics.Add(p); }
                        continue;
                    }
                    if (chp.Spec && ch == '(' && chp.Symbol >= 0) { ch = (char)chp.Symbol; font = Font(chp.SymFtc); }
                    else font = Font(IsEastAsian(ch) ? chp.Ftc1 : chp.Ftc0);
                    Props p2;
                    if (relative)
                    {
                        var baseFont = Font(IsEastAsian(ch) ? bl.Ftc1 : bl.Ftc0);
                        p2 = new Props(chp.Bold && !bl.Bold, chp.Italic && !bl.Italic, chp.Underline && !bl.Underline, chp.Strike && !bl.Strike,
                            Color(chp) is { } c && c != Color(bl) ? c : null, chp.Hps != bl.Hps && chp.Hps > 0 ? chp.Hps / 2.0 : null,
                            font != baseFont ? font : null, Palette(chp.Highlight), seg.Link);
                    }
                    else p2 = new Props(chp.Bold, chp.Italic, chp.Underline, chp.Strike, Color(chp), chp.Hps > 0 ? chp.Hps / 2.0 : null, font, Palette(chp.Highlight), seg.Link);
                    if (look is not null)
                    {
                        if (ch == '\n') look.HasBreak = true;
                        else if (!char.IsWhiteSpace(ch))
                        {
                            look.Chars++;
                            look.Last = ch;
                            look.AllBold &= chp.Bold;
                            look.MinHps = Math.Min(look.MinHps, chp.Hps > 0 ? chp.Hps : 20);
                        }
                    }
                    if (!have || p2 != cur) { Flush(); cur = p2; have = true; }
                    buf.Append(ch);
                }
            }
            Flush();
            return sb.ToString();
        }

        static bool IsEastAsian(char ch) => ch >= 0x2E80 && (ch < 0xE000 || ch >= 0xF900);
    }
}
