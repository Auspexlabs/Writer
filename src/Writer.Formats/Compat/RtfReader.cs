using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Writer.Formats.Compat;

/// <summary>Rich Text Format → blocks: paragraphs with character formatting, headings by style or outline level,
/// bullet and numbered lists, tables (with horizontal and vertical merges), embedded PNG/JPEG pictures, hyperlink
/// fields and page breaks. Headers, footers, footnotes, text boxes and metafile pictures are dropped.</summary>
public static class RtfReader
{
    static readonly Regex HeadingName = new(@"(?:heading|标题)\s*([1-9])", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static List<Block> Read(string rtf, List<string> warnings) => new Parser(warnings).Run(rtf);

    struct CharState
    {
        public bool Bold, Italic, Underline, Strike;
        public int Size, Color, Highlight, Font, Uc;
        public string? Link;
        /// <summary>Inside a destination whose text is not document text.</summary>
        public bool Skip;
        /// <summary>The destination this group is in; "*" right after {\* until the next control word names it.</summary>
        public string? Dest;
    }

    sealed class Parser(List<string> warnings)
    {
        readonly List<Block> _blocks = [];
        readonly Stack<CharState> _stack = new();
        CharState _s = new() { Uc = 1, Size = 24 };
        int _depth;
        Encoding _cp = Encoding.Latin1;
        // header tables
        readonly Dictionary<int, Encoding> _fontCp = [];
        readonly Dictionary<int, string> _fontName = [];
        int _fontBeingRead = -1; readonly StringBuilder _fontNameBuf = new();
        readonly List<string?> _colors = [];
        int _colorR, _colorG, _colorB; bool _colorSeen;
        readonly Dictionary<int, int> _styleHeading = [];   // style index → heading level (0 = Title)
        int _styleBeingRead = -1; readonly StringBuilder _styleNameBuf = new();
        readonly Dictionary<int, string> _listType = [];    // list id → bullet | number
        readonly Dictionary<int, string> _lsType = [];      // list override (\ls) → bullet | number
        int _curListId; string? _curListType; bool _firstLevel;
        // the run and paragraph being built
        readonly StringBuilder _text = new();
        readonly List<byte> _bytes = [];
        CharState _runState;
        readonly StringBuilder _para = new();
        string? _align, _list, _style; int _level, _heading; bool _inTable, _pageBefore;
        // tables
        List<List<TableCell>>? _rows; List<TableCell>? _row; readonly StringBuilder _cell = new();
        readonly List<(int Span, bool Away, bool VMerged, string? Fill)> _cellDefs = [];
        int _cellIndex; bool _cellMergedAway, _cellVMerged; string? _cellFill;
        // pictures and fields
        StringBuilder? _hex; string? _picType; int _picW, _picH, _picDepth;
        StringBuilder? _fldinst; string? _pendingLink; int _fieldDepth = -1;

        public List<Block> Run(string rtf)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var i = 0;
            var n = rtf.Length;
            var skipChars = 0; // after \uN: the fallback characters to swallow
            while (i < n)
            {
                var c = rtf[i];
                if (c == '{') { _stack.Push(_s); _depth++; i++; continue; }
                if (c == '}') { Pop(); i++; continue; }
                if (c == '\\')
                {
                    if (i + 1 >= n) break;
                    var d = rtf[i + 1];
                    if (d == '\'')
                    {
                        if (i + 3 < n && Uri.IsHexDigit(rtf[i + 2]) && Uri.IsHexDigit(rtf[i + 3]))
                        {
                            var b = byte.Parse(rtf.AsSpan(i + 2, 2), NumberStyles.HexNumber);
                            i += 4;
                            if (skipChars > 0) { skipChars--; continue; }
                            _bytes.Add(b);
                            continue;
                        }
                        i += 2;
                        continue;
                    }
                    if (char.IsAsciiLetter(d))
                    {
                        var j = i + 1;
                        while (j < n && char.IsAsciiLetter(rtf[j])) j++;
                        var word = rtf[(i + 1)..j];
                        int? param = null;
                        var k = j;
                        if (k < n && (rtf[k] == '-' || char.IsAsciiDigit(rtf[k])))
                        {
                            var neg = rtf[k] == '-';
                            if (neg) k++;
                            var start = k;
                            while (k < n && char.IsAsciiDigit(rtf[k])) k++;
                            if (k > start && int.TryParse(rtf.AsSpan(start, Math.Min(k - start, 9)), out var v)) param = neg ? -v : v;
                        }
                        if (k < n && rtf[k] == ' ') k++;
                        i = k;
                        if (word == "u" && param is { } code)
                        {
                            FlushBytes();
                            Append(((char)(code < 0 ? code + 65536 : code)).ToString());
                            skipChars = _s.Uc;
                            continue;
                        }
                        if (word == "bin" && param is { } len) { i = Math.Min(n, i + Math.Max(0, len)); continue; }
                        Control(word, param);
                        continue;
                    }
                    i += 2;
                    switch (d)
                    {
                        case '\\' or '{' or '}': Append(d.ToString()); break;
                        case '~': Append(" "); break;
                        case '_': Append("-"); break;
                        case '*': _s.Dest = "*"; break;
                        case '\r' or '\n': Control("par", null); break;
                    }
                    continue;
                }
                i++;
                if (c is '\r' or '\n') continue;
                if (skipChars > 0) { skipChars--; continue; }
                if (_hex is not null && _s.Dest == "pict") { if (Uri.IsHexDigit(c)) _hex.Append(c); continue; }
                Append(c.ToString());
            }
            EndParagraphIfText();
            EndTable();
            return _blocks;
        }

        void Pop()
        {
            if (_stack.Count == 0) return;
            FlushBytes();
            var leaving = _s;
            if (leaving.Dest == "pict" && _hex is not null && _depth == _picDepth) EndPicture();
            if (leaving.Dest == "fldinst" && _fldinst is not null) { ReadField(_fldinst.ToString()); _fldinst = null; }
            if (_depth == _fieldDepth) { _fieldDepth = -1; _pendingLink = null; }
            _s = _stack.Pop();
            _depth--;
            if (leaving.Dest == "fonttbl" && _s.Dest != "fonttbl" && _fontBeingRead >= 0) EndFont();
            if (leaving.Dest == "fonttbl" && _s.Dest == "fonttbl" && _fontBeingRead >= 0 && _fontNameBuf.Length > 0) EndFont();
            if (leaving.Dest == "stylesheet" && _styleBeingRead >= 0 && (_s.Dest != "stylesheet" || _styleNameBuf.Length > 0)) EndStyle();
            if (_text.Length > 0 && !SameRun(_s, _runState)) FlushRun();
        }

        static bool SameRun(CharState a, CharState b) =>
            a.Bold == b.Bold && a.Italic == b.Italic && a.Underline == b.Underline && a.Strike == b.Strike && a.Size == b.Size
            && a.Color == b.Color && a.Highlight == b.Highlight && a.Font == b.Font && a.Link == b.Link;

        // ---- text ----

        void Append(string text)
        {
            FlushBytes();
            switch (_s.Dest)
            {
                case "fonttbl": _fontNameBuf.Append(text); return;
                case "stylesheet": _styleNameBuf.Append(text); return;
                case "fldinst": _fldinst?.Append(text); return;
                case "colortbl": foreach (var ch in text) if (ch == ';') EndColor(); return;
            }
            if (_s.Skip) return;
            if (_text.Length == 0) _runState = _s;
            else if (!SameRun(_s, _runState)) { FlushRun(); _runState = _s; }
            _text.Append(text);
        }

        void FlushBytes()
        {
            if (_bytes.Count == 0) return;
            var enc = (_s.Dest == "fonttbl" ? _fontCp.GetValueOrDefault(_fontBeingRead) : _fontCp.GetValueOrDefault(_s.Font)) ?? _cp;
            var bytes = _bytes.ToArray();
            _bytes.Clear();
            string text;
            try { text = enc.GetString(bytes); } catch (DecoderFallbackException) { text = Encoding.Latin1.GetString(bytes); }
            Append(text);
        }

        void FlushRun()
        {
            if (_text.Length == 0) return;
            var st = _runState;
            _para.Append(Inline.Run(_text.ToString(), st.Bold, st.Italic, st.Underline, st.Strike,
                st.Color > 0 && st.Color < _colors.Count ? _colors[st.Color] : null, st.Size > 0 && st.Size != 24 ? st.Size / 2.0 : null,
                _fontName.GetValueOrDefault(st.Font), st.Highlight > 0 && st.Highlight < _colors.Count ? _colors[st.Highlight] : null, st.Link));
            _text.Clear();
        }

        void EndParagraph()
        {
            FlushBytes();
            FlushRun();
            if (_inTable)
            {
                _rows ??= [];
                if (_cell.Length > 0 || _para.Length > 0) { if (_cell.Length > 0) _cell.Append("<br>"); _cell.Append(_para); }
                _para.Clear();
                return;
            }
            if (_rows is not null) EndTable();
            var html = _para.ToString();
            _para.Clear();
            if (_pageBefore) { _blocks.Add(Block.PageBreak()); _pageBefore = false; }
            if (_heading > 0) _blocks.Add(new Block { Kind = BlockKind.Heading, Level = Math.Clamp(_heading, 1, 6), Html = html, Align = _align });
            else _blocks.Add(new Block { Kind = BlockKind.Paragraph, Html = html, List = _list, Level = _list is null ? 0 : Math.Clamp(_level, 0, 8), Align = _align, Style = _style });
        }

        void EndParagraphIfText() { if (_text.Length > 0 || _para.Length > 0 || _bytes.Count > 0) EndParagraph(); }

        // ---- tables ----

        void EndCell()
        {
            FlushBytes();
            FlushRun();
            _rows ??= [];
            _row ??= [];
            if (_para.Length > 0) { if (_cell.Length > 0) _cell.Append("<br>"); _cell.Append(_para); _para.Clear(); }
            var def = _cellIndex < _cellDefs.Count ? _cellDefs[_cellIndex] : (1, false, false, null);
            _cellIndex++;
            if (def.VMerged)
            {
                // continues the cell above: that one grows a row; this one goes
                var above = _rows.LastOrDefault();
                var target = above?.ElementAtOrDefault(Math.Min(_row.Count, above.Count - 1));
                if (target is not null) target.RowSpan++;
            }
            else if (!def.Away) _row.Add(new TableCell(_cell.ToString()) { ColSpan = def.Span, Fill = def.Fill });
            _cell.Clear();
        }

        void EndRow()
        {
            if (_row is { Count: > 0 }) _rows!.Add(_row);
            _row = null;
            _cellDefs.Clear();
            _cellIndex = 0;
        }

        void EndTable()
        {
            if (_row is { Count: > 0 }) EndRow();
            if (_rows is { Count: > 0 }) _blocks.Add(Block.Table(_rows));
            _rows = null; _row = null; _cell.Clear();
        }

        // ---- pictures and fields ----

        void EndPicture()
        {
            var hex = _hex!.ToString();
            _hex = null;
            if (_picType is null) { warnings.Add("picture in WMF/EMF format dropped"); return; }
            try
            {
                var bytes = Convert.FromHexString(hex);
                if (_inTable) { warnings.Add("picture inside a table dropped"); return; }
                EndParagraphIfText();
                _blocks.Add(Block.Picture(bytes, _picW > 0 ? _picW / 567.0 : null, _picH > 0 ? _picH / 567.0 : null));
            }
            catch (FormatException) { warnings.Add("damaged picture dropped"); }
        }

        void ReadField(string inst)
        {
            var m = Regex.Match(inst, @"HYPERLINK\s+(?:\\l\s+)?""?([^""\s]+)""?", RegexOptions.IgnoreCase);
            if (m.Success) _pendingLink = m.Groups[1].Value;
        }

        // ---- header tables ----

        void EndFont()
        {
            var name = _fontNameBuf.ToString().Trim().TrimEnd(';').Trim();
            if (name.Length > 0 && !_fontName.ContainsKey(_fontBeingRead)) _fontName[_fontBeingRead] = name;
            _fontNameBuf.Clear();
            _fontBeingRead = -1;
        }

        void EndStyle()
        {
            var name = _styleNameBuf.ToString().Trim().TrimEnd(';').Trim();
            if (HeadingName.Match(name) is { Success: true } m) _styleHeading[_styleBeingRead] = int.Parse(m.Groups[1].Value);
            else if (name.Equals("title", StringComparison.OrdinalIgnoreCase)) _styleHeading[_styleBeingRead] = 0;
            _styleNameBuf.Clear();
            _styleBeingRead = -1;
        }

        void EndColor()
        {
            _colors.Add(_colorSeen ? Inline.Rgb(_colorR, _colorG, _colorB) : null);
            _colorR = _colorG = _colorB = 0;
            _colorSeen = false;
        }

        static readonly HashSet<string> Skipped =
        [
            "info", "header", "footer", "headerl", "headerr", "headerf", "footerl", "footerr", "footerf", "footnote", "annotation", "xe", "tc",
            "object", "pntext", "nonshppict", "themedata", "colorschememapping", "latentstyles", "rsidtbl", "generator", "xmlnstbl", "mmathPr",
            "userprops", "docvar", "background", "filetbl", "revtbl", "protusertbl", "ftnsep", "ftnsepc", "ftncn", "aftnsep", "aftnsepc", "aftncn",
            "datastore", "factoidname", "atnid", "atnauthor", "atndate", "bkmkstart", "bkmkend", "pgdsctbl", "template", "wgrffmtfilter",
            "listpicture", "passwordhash", "objdata", "objclass", "objname", "result", "datafield", "formfield", "shpinst", "shprslt", "sp", "sn",
            "sv", "svb", "oldcprops", "oldpprops", "oldsprops", "oldtprops", "mmconnectstr", "mmquery", "mmodso", "operator", "author", "title",
            "subject", "keywords", "comment", "doccomm", "company", "manager", "category", "creatim", "revtim", "printim", "buptim", "hlinkbase",
            "nesttableprops", "panose", "falt", "fname", "listname", "leveltext", "levelnumbers", "lsdlockedexcept", "listtext",
        ];

        void Control(string word, int? p)
        {
            // {\*\word : a destination; the ones read are named here, every other one is skipped whole
            if (_s.Dest == "*")
            {
                _s.Dest = word;
                _s.Skip = true;
                switch (word)
                {
                    case "shppict": _s.Skip = false; break;
                    case "fldinst": _fldinst = new StringBuilder(); break;
                }
                return;
            }
            switch (_s.Dest)
            {
                case "fonttbl":
                    if (word == "f") { if (_fontBeingRead >= 0) EndFont(); _fontBeingRead = p ?? 0; }
                    else if (word == "fcharset" && p is { } cs && _fontBeingRead >= 0 && Charset(cs) is { } page)
                        try { _fontCp[_fontBeingRead] = Encoding.GetEncoding(page); } catch (Exception ex) when (ex is ArgumentException or NotSupportedException) { }
                    return;
                case "colortbl":
                    if (word == "red") { _colorR = p ?? 0; _colorSeen = true; }
                    else if (word == "green") { _colorG = p ?? 0; _colorSeen = true; }
                    else if (word == "blue") { _colorB = p ?? 0; _colorSeen = true; }
                    return;
                case "stylesheet":
                    if (word == "s") { if (_styleBeingRead >= 0) EndStyle(); _styleBeingRead = p ?? 0; }
                    else if (word is "cs" or "ds" or "ts") _styleBeingRead = -1;
                    return;
                case "listtable" or "list" or "listlevel":
                    switch (word)
                    {
                        case "list": _s.Dest = "list"; _curListId = 0; _curListType = null; _firstLevel = true; break;
                        case "listlevel": _s.Dest = "listlevel"; break;
                        case "levelnfc": if (_firstLevel) { _curListType = p == 23 ? "bullet" : "number"; _firstLevel = false; } break;
                        case "listid": _curListId = p ?? 0; if (_curListType is not null) _listType[_curListId] = _curListType; break;
                    }
                    return;
                case "listoverridetable" or "listoverride":
                    switch (word)
                    {
                        case "listoverride": _s.Dest = "listoverride"; _curListId = 0; break;
                        case "listid": _curListId = p ?? 0; break;
                        case "ls": _lsType[p ?? 0] = _listType.GetValueOrDefault(_curListId, "bullet"); break;
                    }
                    return;
                case "pn":
                    if (word == "pnlvlblt") _list = "bullet";
                    else if (word is "pnlvlbody" or "pnlvl" or "pndec" or "pnlcltr" or "pnucltr" or "pnlcrm" or "pnucrm") _list ??= "number";
                    return;
                case "pict":
                    switch (word)
                    {
                        case "pngblip": _picType = "png"; break;
                        case "jpegblip": _picType = "jpeg"; break;
                        case "picwgoal": _picW = p ?? 0; break;
                        case "pichgoal": _picH = p ?? 0; break;
                    }
                    return;
            }
            if (_s.Skip && word != "v") return;
            switch (word)
            {
                case "fonttbl": _s.Dest = "fonttbl"; _s.Skip = true; return;
                case "colortbl": _s.Dest = "colortbl"; _s.Skip = true; return;
                case "stylesheet": _s.Dest = "stylesheet"; _s.Skip = true; return;
                case "listtable": _s.Dest = "listtable"; _s.Skip = true; return;
                case "listoverridetable": _s.Dest = "listoverridetable"; _s.Skip = true; return;
                case "pn": _s.Dest = "pn"; _s.Skip = true; return;
                case "ansicpg": if (p is { } cpg) try { _cp = Encoding.GetEncoding(cpg); } catch (Exception ex) when (ex is ArgumentException or NotSupportedException) { } return;
                case "field": _fieldDepth = _depth; return;
                case "fldrslt": _s.Link = _pendingLink; return;
                case "pict": _s.Dest = "pict"; _s.Skip = true; _hex = new StringBuilder(); _picType = null; _picW = _picH = 0; _picDepth = _depth; return;
                // paragraphs
                case "par": EndParagraph(); return;
                case "sect": EndParagraph(); return;
                case "pard":
                    FlushBytes(); FlushRun();
                    _align = null; _list = null; _level = 0; _heading = 0; _style = null; _inTable = false;
                    return;
                case "intbl": _inTable = true; return;
                case "line": Append("\n"); return;
                case "tab": Append("\t"); return;
                case "page": EndParagraphIfText(); _pageBefore = true; return;
                case "pagebb": _pageBefore = true; return;
                case "qc": _align = "center"; return;
                case "qr": _align = "right"; return;
                case "qj": _align = "justify"; return;
                case "ql": _align = "left"; return;
                case "s": if (p is { } si && _styleHeading.TryGetValue(si, out var lvl)) { if (lvl == 0) _style = "Title"; else _heading = lvl; } return;
                case "outlinelevel": if (p is >= 0 and <= 8 && _heading == 0) _heading = p.Value + 1; return;
                case "ls": _list ??= _lsType.GetValueOrDefault(p ?? 0, "bullet"); return;
                case "ilvl": _level = p ?? 0; return;
                // tables
                case "trowd": if (_rows is null) EndParagraphIfText(); _rows ??= []; _cellDefs.Clear(); _cellIndex = 0; ResetCellDef(); return;
                case "clmgf": _cellMergedAway = false; return;
                case "clmrg":
                    _cellMergedAway = true;
                    for (var k = _cellDefs.Count - 1; k >= 0; k--)
                        if (!_cellDefs[k].Away) { _cellDefs[k] = _cellDefs[k] with { Span = _cellDefs[k].Span + 1 }; break; }
                    return;
                case "clvmrg": _cellVMerged = true; return;
                case "clcbpat": _cellFill = p is { } ci && ci > 0 && ci < _colors.Count ? _colors[ci] : null; return;
                case "cellx": _cellDefs.Add((1, _cellMergedAway, _cellVMerged, _cellFill)); ResetCellDef(); return;
                case "cell": _inTable = true; EndCell(); return;
                case "row": EndRow(); return;
                case "nestcell": Append(" "); return;
                case "nestrow": Append("\n"); return;
                // character formatting
                case "plain": _s.Bold = _s.Italic = _s.Underline = _s.Strike = false; _s.Size = 24; _s.Color = 0; _s.Highlight = 0; return;
                case "b": _s.Bold = p != 0; return;
                case "i": _s.Italic = p != 0; return;
                case "ul" or "uld" or "uldb" or "ulw" or "ulth" or "ulwave" or "uldash": _s.Underline = p != 0; return;
                case "ulnone": _s.Underline = false; return;
                case "strike" or "striked": _s.Strike = p != 0; return;
                case "fs": _s.Size = p ?? 24; return;
                case "cf": _s.Color = p ?? 0; return;
                case "cb" or "highlight" or "chcbpat": _s.Highlight = p ?? 0; return;
                case "f": _s.Font = p ?? 0; return;
                case "uc": _s.Uc = p ?? 1; return;
                case "v": _s.Skip = p != 0; return; // hidden text
                // special characters
                case "emdash": Append("—"); return;
                case "endash": Append("–"); return;
                case "lquote": Append("‘"); return;
                case "rquote": Append("’"); return;
                case "ldblquote": Append("“"); return;
                case "rdblquote": Append("”"); return;
                case "bullet": Append("•"); return;
                case "emspace" or "enspace" or "qmspace": Append(" "); return;
            }
            if (Skipped.Contains(word)) { _s.Dest = word; _s.Skip = true; }
        }

        void ResetCellDef() { _cellMergedAway = false; _cellVMerged = false; _cellFill = null; }

        static int? Charset(int cs) => cs switch
        {
            0 => 1252, 128 => 932, 129 => 949, 134 => 936, 136 => 950, 161 => 1253, 162 => 1254, 163 => 1258, 177 => 1255, 178 => 1256,
            186 => 1257, 204 => 1251, 222 => 874, 238 => 1250, 255 => 437, _ => null,
        };
    }
}
