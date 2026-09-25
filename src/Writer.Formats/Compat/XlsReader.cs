using System.Buffers.Binary;
using System.Text;
using Writer.Core;

namespace Writer.Formats.Compat;

/// <summary>Excel 97-2003 BIFF8 workbooks (.xls, .xlt; Kingsoft WPS .et/.ett share the layout) → sheets.</summary>
public static partial class XlsReader
{
    const string SaveAsXlsx = "Save it as .xlsx in Excel, then open it again.";

    /// <summary>Every worksheet in order (hidden ones too). warnings gets one line per thing dropped.</summary>
    public static List<SheetModel> Read(byte[] file, List<string> warnings)
    {
        var cfb = new Cfb(file);
        var s = cfb.Stream("Workbook");
        if (s is null)
        {
            if (cfb.Has("Book")) throw new WriterException(ErrorCode.FormatError, "This is an Excel 5.0/95 workbook, which Writer cannot read", SaveAsXlsx);
            throw new WriterException(ErrorCode.FormatError, "No workbook inside the file", "The file is not an Excel 97-2003 workbook.");
        }
        Book book;
        try { book = ReadGlobals(s); }
        catch (Exception e) when (e is not WriterException) { throw new WriterException(ErrorCode.FormatError, $"Damaged workbook: {e.Message}", "The file is damaged."); }

        var sheets = new List<SheetModel>();
        foreach (var bs in book.Sheets)
        {
            if (bs.Type != 0) { warnings.Add($"Sheet {bs.Name} is a {(bs.Type == 2 ? "chart" : "macro")} sheet and was skipped"); continue; }
            if (bs.Hidden) warnings.Add($"Sheet {bs.Name} is hidden");
            var sheet = new SheetModel { Name = bs.Name };
            sheets.Add(sheet);
            try { ReadSheet(s, bs.Pos, book, sheet, warnings); }
            catch (Exception e) when (e is not WriterException) { warnings.Add($"Sheet {bs.Name}: stopped at a damaged record ({e.GetType().Name}); kept what came before"); }
        }
        if (sheets.Count == 0) throw new WriterException(ErrorCode.FormatError, "The workbook has no worksheets", "The file is damaged or holds only charts.");
        return sheets;
    }

    // ---- records ----------------------------------------------------------------------------------------------

    /// <summary>One BIFF record with its CONTINUE records appended; Bounds = offsets in Data where each CONTINUE payload starts.</summary>
    sealed class Rec
    {
        public ushort Id;
        public byte[] Data = [];
        public List<int> Bounds = [];
        public int Next;
    }

    static ushort U16(byte[] b, int off) => BinaryPrimitives.ReadUInt16LittleEndian(b.AsSpan(off));
    static int I32(byte[] b, int off) => BinaryPrimitives.ReadInt32LittleEndian(b.AsSpan(off));

    static Rec ReadRec(byte[] s, int off)
    {
        var id = U16(s, off);
        var end = Math.Min(off + 4 + U16(s, off + 2), s.Length);
        var rec = new Rec { Id = id, Next = end };
        if (id == 0x003C || end + 4 > s.Length || U16(s, end) != 0x003C) { rec.Data = s[(off + 4)..end]; return rec; }
        var ms = new MemoryStream();
        ms.Write(s, off + 4, end - off - 4);
        while (rec.Next + 4 <= s.Length && U16(s, rec.Next) == 0x003C)
        {
            var cend = Math.Min(rec.Next + 4 + U16(s, rec.Next + 2), s.Length);
            rec.Bounds.Add((int)ms.Length);
            ms.Write(s, rec.Next + 4, cend - rec.Next - 4);
            rec.Next = cend;
        }
        rec.Data = ms.ToArray();
        return rec;
    }

    /// <summary>Sequential reader over a record's payload; BIFF strings re-read their flags byte where a CONTINUE boundary splits the characters.</summary>
    sealed class Cur(byte[] data, List<int>? bounds = null)
    {
        public readonly byte[] D = data;
        readonly List<int> _b = bounds ?? [];
        int _bi;
        public int P;
        public bool End => P >= D.Length;
        public byte U8() => D[P++];
        public ushort U16() { var v = BinaryPrimitives.ReadUInt16LittleEndian(D.AsSpan(P)); P += 2; return v; }
        public int I32() { var v = BinaryPrimitives.ReadInt32LittleEndian(D.AsSpan(P)); P += 4; return v; }
        public double F64() { var v = BinaryPrimitives.ReadDoubleLittleEndian(D.AsSpan(P)); P += 8; return v; }

        /// <summary>XLUnicodeString (uint16 cch) or ShortXLUnicodeString (byte cch), rich and extended trailers skipped.</summary>
        public string Str(bool shortLen = false)
        {
            int cch = shortLen ? U8() : U16();
            var flags = U8();
            int runs = (flags & 8) != 0 ? U16() : 0;
            int ext = (flags & 4) != 0 ? I32() : 0;
            var wide = (flags & 1) != 0;
            var sb = new StringBuilder(cch);
            for (var i = 0; i < cch && P < D.Length; i++)
            {
                while (_bi < _b.Count && _b[_bi] < P) _bi++;
                if (_bi < _b.Count && _b[_bi] == P) { wide = (U8() & 1) != 0; _bi++; }
                if (P >= D.Length) break;
                sb.Append(wide ? (char)U16() : (char)D[P++]);
            }
            P += runs * 4 + ext;
            return sb.ToString();
        }
    }

    // ---- workbook globals -------------------------------------------------------------------------------------

    sealed class Font { public string Name = ""; public double Size; public bool Bold, Italic, Underline, Strike; public int Icv; }
    sealed class Xf { public int Font, Fmt, HAlign, VAlign, Fls, IcvFore; public bool Wrap; }
    sealed class Supbook { public bool Internal; public string Doc = ""; public List<string> Sheets = [], Names = []; }

    sealed class Book
    {
        public bool Date1904;
        public List<Font> Fonts = [];
        public Dictionary<int, string> Formats = [];
        public List<Xf> Xfs = [];
        public string[] Palette = (string[])DefaultPalette.Clone();
        public List<string> Sst = [];
        public List<(string Name, int Pos, bool Hidden, int Type)> Sheets = [];
        public List<(int Book, int First, int Last)> Xti = [];
        public List<Supbook> Supbooks = [];
        public List<string> Names = [];
        public MemoryStream DrawingGroup = new();
        public List<(byte[]? Bytes, string Kind)>? Blips;

        public Font? FontAt(int ifnt) { if (ifnt >= 4) ifnt--; return ifnt < Fonts.Count ? Fonts[ifnt] : null; }
        public string? Color(int icv) => icv is >= 0 and < 64 && Palette[icv] != "000000" ? Palette[icv] : null;

        public string? FormatCode(int ifmt)
        {
            if (ifmt == 0) return null;
            var code = Formats.TryGetValue(ifmt, out var c) ? c : ifmt < BuiltinFormats.Length ? BuiltinFormats[ifmt] : null;
            return code is null or "General" or "" ? null : code;
        }

        public bool IsDate(int ifmt)
        {
            if (ifmt is >= 14 and <= 22 or >= 45 and <= 47) return true;
            if (ifmt is >= 27 and <= 36 or >= 50 and <= 58) return !Formats.TryGetValue(ifmt, out var c) || IsDateCode(c);
            return Formats.TryGetValue(ifmt, out var code) && IsDateCode(code);
        }
    }

    /// <summary>A custom number format is a date/time format when y, m, d, h or s survive outside quotes, [sections] and escapes.</summary>
    static bool IsDateCode(string code)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < code.Length; i++)
        {
            var ch = code[i];
            if (ch == '"') { var j = code.IndexOf('"', i + 1); i = j < 0 ? code.Length : j; }
            else if (ch == '[') { var j = code.IndexOf(']', i + 1); i = j < 0 ? code.Length : j; }
            else if (ch is '\\' or '_' or '*') i++;
            else sb.Append(ch);
        }
        var s = sb.ToString();
        if (s.Equals("General", StringComparison.OrdinalIgnoreCase)) return false;
        return s.IndexOfAny(['y', 'Y', 'm', 'M', 'd', 'D', 'h', 'H', 's', 'S']) >= 0;
    }

    static Book ReadGlobals(byte[] s)
    {
        var b = new Book();
        var depth = 0;
        Supbook? supbook = null;
        for (var off = 0; off + 4 <= s.Length;)
        {
            var rec = ReadRec(s, off);
            off = rec.Next;
            var d = rec.Data;
            if (rec.Id == 0x0809)
            {
                var vers = d.Length >= 2 ? U16(d, 0) : 0;
                if (depth == 0 && vers is > 0 and < 0x0600)
                    throw new WriterException(ErrorCode.FormatError, "This is an Excel 5.0/95 workbook, which Writer cannot read", SaveAsXlsx);
                depth++;
                continue;
            }
            if (rec.Id == 0x000A) { if (--depth <= 0) break; continue; }
            if (depth != 1) continue;
            switch (rec.Id)
            {
                case 0x002F: throw new WriterException(ErrorCode.FormatError, "The workbook is password protected, which Writer cannot open", "Remove the password in Excel and save it as .xlsx, then open it again.");
                case 0x0022: b.Date1904 = d.Length >= 2 && U16(d, 0) == 1; break;
                case 0x0031:
                {
                    var c = new Cur(d, rec.Bounds);
                    var f = new Font { Size = c.U16() / 20.0 };
                    var grbit = c.U16();
                    f.Icv = c.U16();
                    f.Bold = c.U16() >= 700;
                    c.U16();
                    f.Underline = c.U8() != 0;
                    c.P += 3;
                    f.Name = c.Str(shortLen: true);
                    f.Italic = (grbit & 2) != 0;
                    f.Strike = (grbit & 8) != 0;
                    b.Fonts.Add(f);
                    break;
                }
                case 0x041E:
                {
                    var c = new Cur(d, rec.Bounds);
                    var ifmt = c.U16();
                    b.Formats[ifmt] = c.Str();
                    break;
                }
                case 0x00E0:
                {
                    var x = new Xf { Font = U16(d, 0), Fmt = U16(d, 2) };
                    var alc = d[6];
                    x.HAlign = alc & 7;
                    x.Wrap = (alc & 8) != 0;
                    x.VAlign = (alc >> 4) & 7;
                    if (d.Length >= 20)
                    {
                        x.Fls = (int)(BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(14)) >> 26) & 0x3F;
                        x.IcvFore = U16(d, 18) & 0x7F;
                    }
                    b.Xfs.Add(x);
                    break;
                }
                case 0x0092:
                {
                    var ccv = U16(d, 0);
                    for (var i = 0; i < ccv && 8 + i < 64 && 2 + i * 4 + 3 <= d.Length; i++)
                        b.Palette[8 + i] = Inline.Rgb(d[2 + i * 4], d[3 + i * 4], d[4 + i * 4]);
                    break;
                }
                case 0x00FC:
                {
                    var c = new Cur(d, rec.Bounds);
                    c.I32();
                    var unique = c.I32();
                    for (var i = 0; i < unique && !c.End; i++) b.Sst.Add(c.Str());
                    break;
                }
                case 0x0085:
                {
                    var c = new Cur(d, rec.Bounds);
                    var pos = c.I32();
                    var grbit = c.U16();
                    b.Sheets.Add((c.Str(shortLen: true), pos, (grbit & 3) != 0, grbit >> 8));
                    break;
                }
                case 0x01AE:
                {
                    supbook = new Supbook();
                    b.Supbooks.Add(supbook);
                    var c = new Cur(d, rec.Bounds);
                    var ctab = c.U16();
                    var cch = U16(d, 2);
                    if (cch is 0x0401) supbook.Internal = true;
                    else if (cch is not 0x3A01 && d.Length > 4)
                    {
                        var doc = c.Str();
                        var name = new StringBuilder();
                        foreach (var ch in doc) name.Append(ch < ' ' ? (ch == '\u0003' ? '/' : ' ') : ch);
                        supbook.Doc = name.ToString().Trim();
                        for (var i = 0; i < ctab && !c.End; i++) supbook.Sheets.Add(c.Str());
                    }
                    break;
                }
                case 0x0023 when supbook is not null:
                {
                    var c = new Cur(d, rec.Bounds) { P = 6 };
                    supbook.Names.Add(c.Str(shortLen: true));
                    break;
                }
                case 0x0017:
                {
                    var n = U16(d, 0);
                    for (var i = 0; i < n && 2 + i * 6 + 6 <= d.Length; i++)
                        b.Xti.Add((U16(d, 2 + i * 6), (short)U16(d, 4 + i * 6), (short)U16(d, 6 + i * 6)));
                    break;
                }
                case 0x0018:
                {
                    var grbit = U16(d, 0);
                    var c = new Cur(d, rec.Bounds) { P = 14 };
                    var cch = d[3];
                    var flags = c.U8();
                    string name;
                    if ((grbit & 0x20) != 0 && cch == 1)
                    {
                        var code = (flags & 1) != 0 ? c.U16() : c.U8();
                        name = code < BuiltinNames.Length ? BuiltinNames[code] : "Name" + code;
                    }
                    else
                    {
                        var sb = new StringBuilder();
                        for (var i = 0; i < cch && !c.End; i++) sb.Append((flags & 1) != 0 ? (char)c.U16() : (char)c.U8());
                        name = sb.ToString();
                    }
                    b.Names.Add(name);
                    break;
                }
                case 0x00EB: b.DrawingGroup.Write(d); break;
            }
        }
        if (b.Sheets.Count == 0) throw new WriterException(ErrorCode.FormatError, "The workbook lists no sheets", "The file is damaged.");
        return b;
    }

    static readonly string[] BuiltinNames =
        ["Consolidate_Area", "Auto_Open", "Auto_Close", "Extract", "Database", "Criteria", "Print_Area", "Print_Titles", "Recorder", "Data_Form", "Auto_Activate", "Auto_Deactivate", "Sheet_Title", "_FilterDatabase"];

    static readonly string[] DefaultPalette =
        ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
         "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 " +
         "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF " +
         "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 " +
         "003366 339966 003300 333300 993300 993366 333399 333333").Split(' ');

    /// <summary>Built-in number formats by ifmt; 27-36 and 50-58 as the Chinese Excel defines them. null = not built in.</summary>
    static readonly string?[] BuiltinFormats =
    [
        "General", "0", "0.00", "#,##0", "#,##0.00", "$#,##0_);($#,##0)", "$#,##0_);[Red]($#,##0)", "$#,##0.00_);($#,##0.00)", "$#,##0.00_);[Red]($#,##0.00)",
        "0%", "0.00%", "0.00E+00", "# ?/?", "# ??/??", "m/d/yyyy", "d-mmm-yy", "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yyyy h:mm",
        null, null, null, null, "yyyy\"年\"m\"月\"", "m\"月\"d\"日\"", "m\"月\"d\"日\"", "m-d-yy", "yyyy\"年\"m\"月\"d\"日\"", "h\"时\"mm\"分\"", "h\"时\"mm\"分\"ss\"秒\"",
        "上午/下午h\"时\"mm\"分\"", "上午/下午h\"时\"mm\"分\"ss\"秒\"", "yyyy\"年\"m\"月\"", "#,##0_);(#,##0)", "#,##0_);[Red](#,##0)", "#,##0.00_);(#,##0.00)", "#,##0.00_);[Red](#,##0.00)",
        "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)", "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)", "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
        "mm:ss", "[h]:mm:ss", "mmss.0", "##0.0E+0", "@", "yyyy\"年\"m\"月\"", "m\"月\"d\"日\"", "yyyy\"年\"m\"月\"", "m\"月\"d\"日\"", "m\"月\"d\"日\"",
        "上午/下午h\"时\"mm\"分\"", "上午/下午h\"时\"mm\"分\"ss\"秒\"", "yyyy\"年\"m\"月\"", "m\"月\"d\"日\"",
    ];

    // ---- one worksheet ----------------------------------------------------------------------------------------

    sealed class SheetState
    {
        public SheetModel Sheet = new();
        public List<(CellModel Cell, byte[] Rgce, byte[] Extra)> Formulas = [];
        public List<(int R1, int C1, int R2, int C2, byte[] Rgce, byte[] Extra)> Shared = [];
        public CellModel? PendingString;
        public MemoryStream Drawing = new();
        public double DefaultColWidth = 8.43, DefaultRowHeight = 15;
    }

    static void ReadSheet(byte[] s, int start, Book b, SheetModel sheet, List<string> warnings)
    {
        var st = new SheetState { Sheet = sheet };
        var depth = 0;
        for (var off = start; off + 4 <= s.Length;)
        {
            var rec = ReadRec(s, off);
            off = rec.Next;
            if (rec.Id == 0x0809) { depth++; continue; }
            if (rec.Id == 0x000A) { if (--depth <= 0) break; continue; }
            if (depth != 1) continue;
            var d = rec.Data;
            switch (rec.Id)
            {
                case 0x0208 when d.Length >= 14: // ROW
                    if ((U16(d, 12) & 0x40) != 0) sheet.RowHeights[U16(d, 0) + 1] = (U16(d, 6) & 0x7FFF) / 20.0;
                    break;
                case 0x007D when d.Length >= 10: // COLINFO
                {
                    int first = U16(d, 0), last = Math.Min((int)U16(d, 2), 255);
                    var width = Math.Round(U16(d, 4) / 256.0, 2);
                    var hidden = (U16(d, 8) & 1) != 0;
                    for (var c = first; c <= last; c++) sheet.ColWidths[c + 1] = hidden ? 0 : width;
                    break;
                }
                case 0x0055 when d.Length >= 2: st.DefaultColWidth = U16(d, 0); break;
                case 0x0225 when d.Length >= 4: st.DefaultRowHeight = U16(d, 2) / 20.0; break;
                case 0x0201 when d.Length >= 6: // BLANK
                    Blank(st, b, U16(d, 0), U16(d, 2), U16(d, 4));
                    break;
                case 0x00BE when d.Length >= 8: // MULBLANK
                {
                    int row = U16(d, 0), col = U16(d, 2);
                    for (var i = 0; 4 + i * 2 + 2 <= d.Length - 2; i++) Blank(st, b, row, col + i, U16(d, 4 + i * 2));
                    break;
                }
                case 0x0203 when d.Length >= 14: // NUMBER
                    Number(st, b, U16(d, 0), U16(d, 2), U16(d, 4), BinaryPrimitives.ReadDoubleLittleEndian(d.AsSpan(6)));
                    break;
                case 0x027E when d.Length >= 10: // RK
                    Number(st, b, U16(d, 0), U16(d, 2), U16(d, 4), Rk(I32(d, 6)));
                    break;
                case 0x00BD when d.Length >= 12: // MULRK
                {
                    int row = U16(d, 0), col = U16(d, 2);
                    for (var i = 0; 4 + i * 6 + 6 <= d.Length - 2; i++) Number(st, b, row, col + i, U16(d, 4 + i * 6), Rk(I32(d, 6 + i * 6)));
                    break;
                }
                case 0x00FD when d.Length >= 10: // LABELSST
                {
                    var isst = I32(d, 6);
                    Cell(st, b, U16(d, 0), U16(d, 2), U16(d, 4)).Value = isst >= 0 && isst < b.Sst.Count ? b.Sst[isst] : "";
                    break;
                }
                case 0x0204 or 0x00D6 when d.Length >= 9: // LABEL, RSTRING
                    Cell(st, b, U16(d, 0), U16(d, 2), U16(d, 4)).Value = new Cur(d, rec.Bounds) { P = 6 }.Str();
                    break;
                case 0x0205 when d.Length >= 8: // BOOLERR
                    Cell(st, b, U16(d, 0), U16(d, 2), U16(d, 4)).Value = d[7] == 0 ? d[6] != 0 : ErrorText(d[6]);
                    break;
                case 0x0006 when d.Length >= 22: // FORMULA
                {
                    var cell = Cell(st, b, U16(d, 0), U16(d, 2), U16(d, 4));
                    if (U16(d, 12) == 0xFFFF)
                    {
                        cell.Value = d[6] switch { 0 => "", 1 => d[8] != 0, 2 => ErrorText(d[8]), _ => "" };
                        if (d[6] == 0) st.PendingString = cell;
                    }
                    else cell.Value = Typed(b, U16(d, 4), BinaryPrimitives.ReadDoubleLittleEndian(d.AsSpan(6)));
                    var cce = U16(d, 20);
                    if (22 + cce <= d.Length) st.Formulas.Add((cell, d[22..(22 + cce)], d[(22 + cce)..]));
                    break;
                }
                case 0x0207 when st.PendingString is not null: // STRING
                    st.PendingString.Value = new Cur(d, rec.Bounds).Str();
                    st.PendingString = null;
                    break;
                case 0x04BC when d.Length >= 10: // SHRFMLA
                {
                    var cce = U16(d, 8);
                    if (10 + cce <= d.Length) st.Shared.Add((U16(d, 0), d[4], U16(d, 2), d[5], d[10..(10 + cce)], d[(10 + cce)..]));
                    break;
                }
                case 0x0221 when d.Length >= 14: // ARRAY
                {
                    var cce = U16(d, 12);
                    if (14 + cce <= d.Length) st.Shared.Add((U16(d, 0), d[4], U16(d, 2), d[5], d[14..(14 + cce)], d[(14 + cce)..]));
                    break;
                }
                case 0x00E5 when d.Length >= 2: // MERGECELLS
                {
                    var n = U16(d, 0);
                    for (var i = 0; i < n && 2 + i * 8 + 8 <= d.Length; i++)
                    {
                        int r1 = U16(d, 2 + i * 8), r2 = U16(d, 4 + i * 8), c1 = U16(d, 6 + i * 8), c2 = U16(d, 8 + i * 8);
                        if (r2 > r1 || c2 > c1) sheet.Merges.Add($"{A1(r1, c1)}:{A1(r2, c2)}");
                    }
                    break;
                }
                case 0x00EC: st.Drawing.Write(d); break;
            }
        }
        DecodeFormulas(st, b, warnings);
        ReadPictures(st, b, warnings);
    }

    static CellModel Cell(SheetState st, Book b, int row, int col, int ixfe)
    {
        var cell = new CellModel { Row = row + 1, Col = col + 1 };
        Style(b, cell, ixfe);
        st.Sheet.Cells.Add(cell);
        return cell;
    }

    static void Blank(SheetState st, Book b, int row, int col, int ixfe)
    {
        var cell = new CellModel { Row = row + 1, Col = col + 1 };
        Style(b, cell, ixfe);
        if (cell.Fill is not null) st.Sheet.Cells.Add(cell);
    }

    static void Number(SheetState st, Book b, int row, int col, int ixfe, double v)
    {
        Cell(st, b, row, col, ixfe).Value = Typed(b, ixfe, v);
    }

    /// <summary>A number, or the date it stands for when the cell's number format is a date/time format.</summary>
    static object Typed(Book b, int ixfe, double v)
    {
        if (ixfe < 0 || ixfe >= b.Xfs.Count || !b.IsDate(b.Xfs[ixfe].Fmt) || double.IsNaN(v) || double.IsInfinity(v)) return v;
        if (b.Date1904) return v is >= 0 and < 2957004 ? new DateTime(1904, 1, 1).AddDays(v) : v;
        if (v is < 0 or >= 2958466) return v;
        // 1900 system: serial 60 is Excel's fake 29 Feb 1900; below it the serials run one day early
        return v < 60 ? new DateTime(1899, 12, 31).AddDays(v) : new DateTime(1899, 12, 30).AddDays(v < 61 ? v - 1 : v);
    }

    static void Style(Book b, CellModel cell, int ixfe)
    {
        if (ixfe < 0 || ixfe >= b.Xfs.Count) return;
        var x = b.Xfs[ixfe];
        var f = b.FontAt(x.Font);
        var normal = b.Fonts.Count > 0 ? b.Fonts[0] : null;
        if (f is not null)
        {
            cell.Bold = f.Bold; cell.Italic = f.Italic; cell.Underline = f.Underline; cell.Strike = f.Strike;
            if (f.Name.Length > 0 && f.Name != normal?.Name) cell.Font = f.Name;
            if (f.Size > 0 && f.Size != normal?.Size) cell.SizePt = f.Size;
            cell.Color = b.Color(f.Icv);
        }
        if (x.Fls == 1 && x.IcvFore is not (64 or 65 or 0x7F)) cell.Fill = b.Palette[x.IcvFore & 63];
        cell.Format = b.FormatCode(x.Fmt);
        cell.Wrap = x.Wrap;
        cell.Align = x.HAlign switch { 1 => "left", 2 or 6 => "center", 3 => "right", 5 or 7 => "justify", _ => null };
        cell.VAlign = x.VAlign switch { 0 => "top", 1 => "middle", _ => null }; // bottom is Excel's default
    }

    static double Rk(int rk)
    {
        var v = (rk & 2) != 0 ? rk >> 2 : BitConverter.Int64BitsToDouble((long)(rk & ~3) << 32);
        return (rk & 1) != 0 ? v / 100 : v;
    }

    static string ErrorText(byte code) => code switch
    {
        0x00 => "#NULL!", 0x07 => "#DIV/0!", 0x0F => "#VALUE!", 0x17 => "#REF!", 0x1D => "#NAME?", 0x24 => "#NUM!", 0x2A => "#N/A", 0x2B => "#GETTING_DATA", _ => "#VALUE!",
    };

    static string ColName(int col)
    {
        var s = "";
        for (var c = col; c >= 0; c = c / 26 - 1) s = (char)('A' + c % 26) + s;
        return s;
    }

    static string A1(int row, int col) => ColName(col) + (row + 1);
}
