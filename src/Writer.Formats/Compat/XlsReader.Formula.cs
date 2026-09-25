using System.Globalization;
using System.Text;
using Writer.Core;

namespace Writer.Formats.Compat;

public static partial class XlsReader
{
    /// <summary>Turns every collected FORMULA record's tokens into Excel text; a formula that will not decode keeps its cached value and adds a warning.</summary>
    static void DecodeFormulas(SheetState st, Book b, List<string> warnings)
    {
        var masters = new Dictionary<(int, int), int>();
        for (var i = 0; i < st.Shared.Count; i++) masters.TryAdd((st.Shared[i].R1, st.Shared[i].C1), i);
        var failed = 0;
        foreach (var (cell, rgce, extra) in st.Formulas)
        {
            try { cell.Formula = new Fx(b, st, masters, cell.Row - 1, cell.Col - 1).Decode(rgce, extra); }
            catch (Exception e) when (e is not WriterException)
            {
                if (++failed <= 10) warnings.Add($"{st.Sheet.Name}!{A1(cell.Row - 1, cell.Col - 1)}: formula dropped, its value kept ({e.Message})");
            }
        }
        if (failed > 10) warnings.Add($"{st.Sheet.Name}: {failed - 10} more formulas dropped, their values kept");
    }

    static readonly string[] Ops = ["+", "-", "*", "/", "^", "&", "<", "<=", "=", ">=", ">", "<>", " ", ",", ":"];

    /// <summary>Function names by iftab; /n = the fixed argument count ptgFunc relies on.</summary>
    static readonly Dictionary<int, (string Name, int Args)> Funcs = ParseFuncs(
        "0=COUNT 1=IF 2=ISNA/1 3=ISERROR/1 4=SUM 5=AVERAGE 6=MIN 7=MAX 8=ROW 9=COLUMN 10=NA/0 11=NPV 12=STDEV 13=DOLLAR 14=FIXED 15=SIN/1 16=COS/1 17=TAN/1 18=ATAN/1 19=PI/0 " +
        "20=SQRT/1 21=EXP/1 22=LN/1 23=LOG10/1 24=ABS/1 25=INT/1 26=SIGN/1 27=ROUND/2 28=LOOKUP 29=INDEX 30=REPT/2 31=MID/3 32=LEN/1 33=VALUE/1 34=TRUE/0 35=FALSE/0 36=AND 37=OR 38=NOT/1 39=MOD/2 " +
        "40=DCOUNT/3 41=DSUM/3 42=DAVERAGE/3 43=DMIN/3 44=DMAX/3 45=DSTDEV/3 46=VAR 47=DVAR/3 48=TEXT/2 49=LINEST 50=TREND 51=LOGEST 52=GROWTH 56=PV 57=FV 58=NPER 59=PMT 60=RATE 61=MIRR/3 62=IRR 63=RAND/0 " +
        "64=MATCH 65=DATE/3 66=TIME/3 67=DAY/1 68=MONTH/1 69=YEAR/1 70=WEEKDAY 71=HOUR/1 72=MINUTE/1 73=SECOND/1 74=NOW/0 75=AREAS/1 76=ROWS/1 77=COLUMNS/1 78=OFFSET 82=SEARCH 83=TRANSPOSE/1 86=TYPE/1 " +
        "97=ATAN2/2 98=ASIN/1 99=ACOS/1 100=CHOOSE 101=HLOOKUP 102=VLOOKUP 105=ISREF/1 109=LOG 111=CHAR/1 112=LOWER/1 113=UPPER/1 114=PROPER/1 115=LEFT 116=RIGHT 117=EXACT/2 118=TRIM/1 119=REPLACE/4 " +
        "120=SUBSTITUTE 121=CODE/1 124=FIND 125=CELL 126=ISERR/1 127=ISTEXT/1 128=ISNUMBER/1 129=ISBLANK/1 130=T/1 131=N/1 140=DATEVALUE/1 141=TIMEVALUE/1 142=SLN/3 143=SYD/4 144=DDB 148=INDIRECT " +
        "162=CLEAN/1 163=MDETERM/1 164=MINVERSE/1 165=MMULT/2 167=IPMT 168=PPMT 169=COUNTA 183=PRODUCT 184=FACT/1 189=DPRODUCT/3 190=ISNONTEXT/1 193=STDEVP 194=VARP 195=DSTDEVP/3 196=DVARP/3 197=TRUNC " +
        "198=ISLOGICAL/1 199=DCOUNTA/3 204=USDOLLAR 205=FINDB 206=SEARCHB 207=REPLACEB/4 208=LEFTB 209=RIGHTB 210=MIDB/3 211=LENB/1 212=ROUNDUP/2 213=ROUNDDOWN/2 214=ASC/1 215=DBCS/1 216=RANK 219=ADDRESS " +
        "220=DAYS360 221=TODAY/0 222=VDB 227=MEDIAN 228=SUMPRODUCT 229=SINH/1 230=COSH/1 231=TANH/1 232=ASINH/1 233=ACOSH/1 234=ATANH/1 235=DGET/3 244=INFO/1 247=DB 252=FREQUENCY/2 261=ERROR.TYPE/1 " +
        "269=AVEDEV 270=BETADIST 271=GAMMALN/1 272=BETAINV 273=BINOMDIST/4 274=CHIDIST/2 275=CHIINV/2 276=COMBIN/2 277=CONFIDENCE/3 278=CRITBINOM/3 279=EVEN/1 280=EXPONDIST/3 281=FDIST/3 282=FINV/3 " +
        "283=FISHER/1 284=FISHERINV/1 285=FLOOR/2 286=GAMMADIST/4 287=GAMMAINV/3 288=CEILING/2 289=HYPGEOMDIST/4 290=LOGNORMDIST/3 291=LOGINV/3 292=NEGBINOMDIST/3 293=NORMDIST/4 294=NORMSDIST/1 " +
        "295=NORMINV/3 296=NORMSINV/1 297=STANDARDIZE/3 298=ODD/1 299=PERMUT/2 300=POISSON/3 301=TDIST/3 302=WEIBULL/4 303=SUMXMY2/2 304=SUMX2MY2/2 305=SUMX2PY2/2 306=CHITEST/2 307=CORREL/2 308=COVAR/2 " +
        "309=FORECAST/3 310=FTEST/2 311=INTERCEPT/2 312=PEARSON/2 313=RSQ/2 314=STEYX/2 315=SLOPE/2 316=TTEST/4 317=PROB 318=DEVSQ 319=GEOMEAN 320=HARMEAN 321=SUMSQ 322=KURT 323=SKEW 324=ZTEST 325=LARGE/2 " +
        "326=SMALL/2 327=QUARTILE/2 328=PERCENTILE/2 329=PERCENTRANK 330=MODE 331=TRIMMEAN/2 332=TINV/2 336=CONCATENATE 337=POWER/2 342=RADIANS/1 343=DEGREES/1 344=SUBTOTAL 345=SUMIF 346=COUNTIF/2 " +
        "347=COUNTBLANK/1 350=ISPMT/4 351=DATEDIF/3 352=DATESTRING/1 353=NUMBERSTRING/2 354=ROMAN 358=GETPIVOTDATA 359=HYPERLINK 360=PHONETIC/1 361=AVERAGEA 362=MAXA 363=MINA 364=STDEVPA 365=VARPA 366=STDEVA 367=VARA");

    static Dictionary<int, (string, int)> ParseFuncs(string table)
    {
        var d = new Dictionary<int, (string, int)>();
        foreach (var e in table.Split(' '))
        {
            var eq = e.IndexOf('=');
            var slash = e.IndexOf('/');
            d[int.Parse(e[..eq], CultureInfo.InvariantCulture)] = slash < 0 ? (e[(eq + 1)..], -1) : (e[(eq + 1)..slash], int.Parse(e[(slash + 1)..], CultureInfo.InvariantCulture));
        }
        return d;
    }

    /// <summary>One formula's token stream (rgce) plus its trailing constants (rgcb) → Excel text, for the cell at row/col (0-based).</summary>
    sealed class Fx(Book b, SheetState st, Dictionary<(int, int), int> masters, int row, int col)
    {
        Cur _c = new([]), _x = new([]);
        readonly Stack<string> _s = new();

        public string Decode(byte[] rgce, byte[] extra)
        {
            _c = new Cur(rgce);
            _x = new Cur(extra);
            while (!_c.End) Token();
            if (_s.Count != 1) throw new InvalidDataException($"{_s.Count} operands left");
            return _s.Pop();
        }

        string Pop() => _s.Count > 0 ? _s.Pop() : throw new InvalidDataException("operand missing");

        void Token()
        {
            var ptg = _c.U8();
            var t = ptg >= 0x20 ? 0x20 | (ptg & 0x1F) : ptg;
            switch (t)
            {
                case 0x01: Shared(); break;
                case 0x02: throw new InvalidDataException("data table");
                case >= 0x03 and <= 0x11: { var r = Pop(); var l = Pop(); _s.Push(l + Ops[t - 3] + r); break; }
                case 0x12: _s.Push("+" + Pop()); break;
                case 0x13: _s.Push("-" + Pop()); break;
                case 0x14: _s.Push(Pop() + "%"); break;
                case 0x15: _s.Push("(" + Pop() + ")"); break;
                case 0x16: _s.Push(""); break;
                case 0x17: _s.Push("\"" + _c.Str(shortLen: true).Replace("\"", "\"\"") + "\""); break;
                case 0x19: Attr(); break;
                case 0x1C: _s.Push(ErrorText(_c.U8())); break;
                case 0x1D: _s.Push(_c.U8() != 0 ? "TRUE" : "FALSE"); break;
                case 0x1E: _s.Push(_c.U16().ToString(CultureInfo.InvariantCulture)); break;
                case 0x1F: _s.Push(Num(_c.F64())); break;
                case 0x20: _c.P += 7; _s.Push(Array()); break;
                case 0x21: Func(_c.U16(), -1); break;
                case 0x22: { var n = _c.U8() & 0x7F; Func(_c.U16() & 0x7FFF, n); break; }
                case 0x23: { var i = _c.U16(); _c.P += 2; _s.Push(Name(i)); break; }
                case 0x24: _s.Push(Ref(false)); break;
                case 0x25: _s.Push(Area(false)); break;
                case 0x26: _c.P += 6; if (!_x.End) _x.P += 2 + _x.U16() * 8; break;
                case 0x27 or 0x28: _c.P += 6; break;
                case 0x29 or 0x2E or 0x2F: _c.P += 2; break;
                case 0x2A: _c.P += 4; _s.Push("#REF!"); break;
                case 0x2B: _c.P += 8; _s.Push("#REF!"); break;
                case 0x2C: _s.Push(Ref(true)); break;
                case 0x2D: _s.Push(Area(true)); break;
                case 0x39: { var ixti = _c.U16(); var i = _c.U16(); _c.P += 2; _s.Push(NameX(ixti, i)); break; }
                case 0x3A: { var sh = Sheet3d(_c.U16()); _s.Push(sh + Ref(false)); break; }
                case 0x3B: { var sh = Sheet3d(_c.U16()); _s.Push(sh + Area(false)); break; }
                case 0x3C: { var sh = Sheet3d(_c.U16()); _c.P += 4; _s.Push(sh + "#REF!"); break; }
                case 0x3D: { var sh = Sheet3d(_c.U16()); _c.P += 8; _s.Push(sh + "#REF!"); break; }
                default: throw new InvalidDataException($"token 0x{ptg:X2}");
            }
        }

        void Shared()
        {
            int r = _c.U16(), c = _c.U16();
            if (!masters.TryGetValue((r, c), out var i))
            {
                i = st.Shared.FindIndex(x => x.R1 <= row && row <= x.R2 && x.C1 <= col && col <= x.C2);
                if (i < 0) throw new InvalidDataException("shared formula not found");
            }
            var (saveC, saveX) = (_c, _x);
            _c = new Cur(st.Shared[i].Rgce);
            _x = new Cur(st.Shared[i].Extra);
            while (!_c.End) Token();
            (_c, _x) = (saveC, saveX);
        }

        void Attr()
        {
            var grbit = _c.U8();
            var w = _c.U16();
            if ((grbit & 0x04) != 0) _c.P += (w + 1) * 2;
            else if ((grbit & 0x10) != 0) _s.Push("SUM(" + Pop() + ")");
        }

        void Func(int iftab, int nargs)
        {
            if (iftab == 255)
            {
                if (nargs < 1) throw new InvalidDataException("add-in function without a name");
                var args = PopArgs(nargs - 1);
                _s.Push(Pop() + "(" + args + ")");
                return;
            }
            if (!Funcs.TryGetValue(iftab, out var f)) throw new InvalidDataException($"function {iftab}");
            if (nargs < 0) nargs = f.Args;
            if (nargs < 0) throw new InvalidDataException($"{f.Name} argument count");
            _s.Push(f.Name + "(" + PopArgs(nargs) + ")");
        }

        string PopArgs(int n)
        {
            var a = new string[n];
            for (var i = n - 1; i >= 0; i--) a[i] = Pop();
            return string.Join(",", a);
        }

        string Array()
        {
            if (_x.End) throw new InvalidDataException("array constant missing");
            int cols = _x.U8() + 1, rows = _x.U16() + 1;
            var sb = new StringBuilder("{");
            for (var r = 0; r < rows; r++)
            {
                if (r > 0) sb.Append(';');
                for (var c = 0; c < cols; c++)
                {
                    if (c > 0) sb.Append(',');
                    switch (_x.U8())
                    {
                        case 0: _x.P += 8; break;
                        case 1: sb.Append(Num(_x.F64())); break;
                        case 2: sb.Append('"').Append(_x.Str().Replace("\"", "\"\"")).Append('"'); break;
                        case 4: sb.Append(_x.U8() != 0 ? "TRUE" : "FALSE"); _x.P += 7; break;
                        case 16: sb.Append(ErrorText(_x.U8())); _x.P += 7; break;
                        default: throw new InvalidDataException("array constant");
                    }
                }
            }
            return sb.Append('}').ToString();
        }

        string Ref(bool rel)
        {
            int r = _c.U16(), cf = _c.U16();
            return RefText(r, cf, rel);
        }

        string RefText(int r, int cf, bool rel)
        {
            var c = cf & 0x3FFF;
            bool colRel = (cf & 0x4000) != 0, rowRel = (cf & 0x8000) != 0;
            if (rel)
            {
                if (rowRel) r = ((short)r + row) & 0xFFFF;
                if (colRel) c = ((sbyte)(c & 0xFF) + col) & 0xFF;
            }
            return (colRel ? "" : "$") + ColName(c) + (rowRel ? "" : "$") + (r + 1);
        }

        string Area(bool rel)
        {
            int r1 = _c.U16(), r2 = _c.U16(), cf1 = _c.U16(), cf2 = _c.U16();
            if (!rel && r1 == 0 && r2 == 0xFFFF) return Col(cf1) + ":" + Col(cf2);
            if (!rel && (cf1 & 0x3FFF) == 0 && (cf2 & 0x3FFF) == 0xFF) return Row(r1, cf1) + ":" + Row(r2, cf2);
            return RefText(r1, cf1, rel) + ":" + RefText(r2, cf2, rel);

            static string Col(int cf) => ((cf & 0x4000) != 0 ? "" : "$") + ColName(cf & 0x3FFF);
            static string Row(int r, int cf) => ((cf & 0x8000) != 0 ? "" : "$") + (r + 1);
        }

        string Name(int ilbl) => ilbl >= 1 && ilbl <= b.Names.Count ? b.Names[ilbl - 1] : throw new InvalidDataException($"name {ilbl}");

        string NameX(int ixti, int ilbl)
        {
            var sb = ixti < b.Xti.Count && b.Xti[ixti].Book < b.Supbooks.Count ? b.Supbooks[b.Xti[ixti].Book] : null;
            if (sb is null || sb.Internal) return Name(ilbl);
            if (ilbl < 1 || ilbl > sb.Names.Count) throw new InvalidDataException($"external name {ilbl}");
            var name = sb.Names[ilbl - 1];
            return name.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase) ? name[6..] : name;
        }

        string Sheet3d(int ixti)
        {
            if (ixti >= b.Xti.Count) throw new InvalidDataException($"sheet reference {ixti}");
            var (ib, first, last) = b.Xti[ixti];
            var sb = ib < b.Supbooks.Count ? b.Supbooks[ib] : null;
            if (sb is null || sb.Internal)
            {
                if (first < 0 || first >= b.Sheets.Count) return "#REF!";
                var name = b.Sheets[first].Name;
                if (last != first && last >= 0 && last < b.Sheets.Count) name += ":" + b.Sheets[last].Name;
                return Quote(name) + "!";
            }
            var s1 = first >= 0 && first < sb.Sheets.Count ? sb.Sheets[first] : "";
            var s2 = last != first && last >= 0 && last < sb.Sheets.Count ? ":" + sb.Sheets[last] : "";
            return Quote($"[{sb.Doc}]{s1}{s2}") + "!";
        }

        static string Quote(string name)
        {
            var plain = name.Length > 0 && !char.IsDigit(name[0]) && name.All(ch => char.IsLetterOrDigit(ch) || ch is '_' or '.' or ':');
            return plain ? name : "'" + name.Replace("'", "''") + "'";
        }

        static string Num(double v) => v.ToString(CultureInfo.InvariantCulture);
    }
}
