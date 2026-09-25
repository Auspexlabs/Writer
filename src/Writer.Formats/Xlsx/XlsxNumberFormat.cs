using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Writer.Formats.Xlsx;

/// <summary>A number as Excel shows it under a format code: sections (positive;negative;zero), literals ("x", \x, _x), [$¥-804]
/// currency and [Red] colours, digit masks (0 # ? , . % E+) and dates and times (y m d h s AM/PM, [h] elapsed) on 1900 serials.
/// The rules of the editor's formatter (ui/sheet-engine.js fmtCode). ponytail: no [&gt;100] conditions.</summary>
static class XlsxNumberFormat
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    static readonly string[] Months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    static readonly string[] Days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const string ChineseDays = "日一二三四五六";
    static readonly Dictionary<string, string> Named = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Black"] = "000000", ["Blue"] = "0000FF", ["Cyan"] = "00FFFF", ["Green"] = "00FF00", ["Magenta"] = "FF00FF", ["Red"] = "FF0000", ["White"] = "FFFFFF", ["Yellow"] = "FFFF00",
    };
    static readonly Regex DateToken = new("^(y+|d+|h+|s+|m+|a+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    static readonly Regex AmPm = new("^(AM/PM|A/P|上午/下午)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static string General(double v) => v == Math.Floor(v) && Math.Abs(v) < 1e15 ? v.ToString("0", Inv) : v.ToString("G10", Inv);

    /// <summary>The text, and the colour (RRGGBB) when the section names one.</summary>
    public static (string Text, string? Color) Format(double v, string? code)
    {
        if (string.IsNullOrEmpty(code) || code.Equals("General", StringComparison.OrdinalIgnoreCase)) return (General(v), null);
        var sections = Sections(code);
        var sec = sections[0];
        var neg = v < 0;
        if (neg && sections.Count > 1) { sec = sections[1]; v = -v; neg = false; }
        else if (v == 0 && sections.Count > 2) sec = sections[2];
        string? color = null;
        var toks = new List<(char K, string V)>(); // l literal, n digit mask, d date/time, a AM/PM
        for (var i = 0; i < sec.Length;)
        {
            var ch = sec[i];
            if (ch == '"')
            {
                var j = sec.IndexOf('"', i + 1);
                if (j < 0) j = sec.Length;
                toks.Add(('l', sec[(i + 1)..j]));
                i = j + 1;
                continue;
            }
            if (ch == '\\') { toks.Add(('l', i + 1 < sec.Length ? sec[i + 1].ToString() : "")); i += 2; continue; }
            if (ch == '_') { toks.Add(('l', " ")); i += 2; continue; }
            if (ch == '*') { i += 2; continue; }
            if (ch == '[')
            {
                var j = sec.IndexOf(']', i);
                if (j < 0) j = sec.Length;
                var b = sec[(i + 1)..j];
                i = j + 1;
                if (b.StartsWith('$')) toks.Add(('l', b[1..].Split('-')[0]));
                else if (Regex.IsMatch(b, "^(h+|m+|s+)$", RegexOptions.IgnoreCase)) toks.Add(('d', "[" + b.ToLowerInvariant() + "]"));
                else if (Named.TryGetValue(b, out var named)) color = named;
                continue;
            }
            var rest = sec[i..];
            if (rest.StartsWith("General", StringComparison.OrdinalIgnoreCase)) { toks.Add(('n', "@")); i += 7; continue; }
            Match m;
            if ((m = AmPm.Match(rest)).Success) { toks.Add(('a', m.Value)); i += m.Length; continue; }
            if (rest.Length > 1 && ch is 'E' or 'e' && rest[1] is '+' or '-') { toks.Add(('n', "E")); i += 2; continue; }
            if ((m = DateToken.Match(rest)).Success) { toks.Add(('d', m.Value.ToLowerInvariant())); i += m.Length; continue; }
            toks.Add(("0#?,.%@".Contains(ch) ? 'n' : 'l', ch.ToString()));
            i++;
        }
        var isDate = toks.Any(t => t.K == 'a' || t.K == 'd' && "ydhsa[".Contains(t.V[0])) || toks.Any(t => t.K == 'd') && !toks.Any(t => t.K == 'n' && "0#?".Contains(t.V));
        if (isDate) return (v is < 0 or > 2958465 ? "#####" : Date(v, toks), color);
        var fraction = Regex.Match(Regex.Replace(sec, "\"[^\"]*\"", ""), @"(\?+)\s*/\s*(\?+|\d+)");
        if (fraction.Success) return (Fraction(v, Regex.IsMatch(sec, @"[#0]\s*\?"), fraction.Groups[2].Value), color); // # ?/?, # ??/??, ?/4
        if (toks.Any(t => t is ('n', "@"))) return (string.Concat(toks.Select(t => t is ('n', "@") ? General(v) : t.K == 'l' ? t.V : "")), color);
        var mask = string.Concat(toks.Where(t => t.K == 'n').Select(t => t.V));
        if (mask.IndexOfAny(['0', '#', '?']) < 0) return ((neg ? "-" : "") + string.Concat(toks.Where(t => t.K == 'l' || t.V == "%").Select(t => t.V)), color);
        var x = Math.Abs(v);
        foreach (var _ in mask.Where(c => c == '%')) x *= 100;
        var sci = mask.Contains('E');
        var body = (sci ? mask[..mask.IndexOf('E')] : mask).Replace("%", "");
        var dot = body.IndexOf('.');
        var ip = dot < 0 ? body : body[..dot];
        var dp = dot < 0 ? "" : body[(dot + 1)..];
        var scaled = ip.Length - ip.TrimEnd(',').Length;
        for (var k = 0; k < scaled; k++) x /= 1000;
        ip = ip.TrimEnd(',');
        var grouping = Regex.IsMatch(ip, "[0#?],[0#?]");
        int minInt = ip.Count(c => c == '0'), minDec = dp.Count(c => c == '0'), maxDec = dp.Count(c => c is '0' or '#' or '?');
        var exp = "";
        if (sci)
        {
            var digits = mask[(mask.IndexOf('E') + 1)..].Count(c => c == '0');
            var e = x == 0 ? 0 : (int)Math.Floor(Math.Log10(x));
            var mant = x / Math.Pow(10, e);
            if (Round(mant, maxDec) >= 10) { mant /= 10; e++; }
            x = mant;
            exp = "E" + (e < 0 ? "-" : "+") + Math.Abs(e).ToString(Inv).PadLeft(Math.Max(1, digits), '0');
        }
        var fixedText = Round(x, maxDec).ToString("F" + maxDec.ToString(Inv), Inv);
        var point = fixedText.IndexOf('.');
        var intS = point < 0 ? fixedText : fixedText[..point];
        var decS = point < 0 ? "" : fixedText[(point + 1)..];
        while (decS.Length > minDec && decS.EndsWith('0')) decS = decS[..^1];
        intS = intS == "0" && minInt == 0 ? "" : intS.PadLeft(minInt, '0');
        if (grouping && intS.Length > 3) intS = Regex.Replace(intS, @"\B(?=(\d{3})+(?!\d))", ",");
        var number = intS + (decS.Length > 0 ? "." + decS : "") + exp;
        var sb = new StringBuilder(neg ? "-" : "");
        var placed = false;
        foreach (var (k, t) in toks)
        {
            if (k == 'l') sb.Append(t);
            else if (k == 'n' && t == "%") sb.Append('%');
            else if (k == 'n' && !placed) { sb.Append(number); placed = true; }
        }
        return (sb.ToString(), color);
    }

    /// <summary>A fraction as Excel writes it: a whole part when the code has one, the nearest fraction with at most as many
    /// denominator digits as the code's ?s (or its fixed denominator).</summary>
    static string Fraction(double v, bool whole, string den)
    {
        var x = Math.Abs(v);
        var ip = whole ? Math.Floor(x) : 0;
        var f = x - ip;
        long n = 0, d = 1;
        if (den.All(char.IsAsciiDigit)) { d = long.Parse(den, Inv); n = (long)Math.Round(f * d); }
        else
        {
            var err = double.MaxValue;
            for (long q = 1; q < Math.Pow(10, den.Length); q++)
            {
                var p = (long)Math.Round(f * q);
                var e = Math.Abs(f - (double)p / q);
                if (e < err - 1e-12) (err, n, d) = (e, p, q);
            }
        }
        if (whole && n == d) (ip, n) = (ip + 1, 0);
        var text = whole ? n != 0 ? (ip != 0 ? ip.ToString(Inv) + " " : "") + n.ToString(Inv) + "/" + d.ToString(Inv) : ip.ToString(Inv) : n.ToString(Inv) + "/" + d.ToString(Inv);
        return (v < 0 ? "-" : "") + text;
    }

    static decimal Round(double x, int digits) =>
        Math.Abs(x) < 7.9e27 ? Math.Round((decimal)x, Math.Min(digits, 28), MidpointRounding.AwayFromZero) : (decimal)Math.Round(x);

    /// <summary>Splits the code at ; outside quotes, brackets and escapes.</summary>
    static List<string> Sections(string code)
    {
        var list = new List<string>();
        var start = 0;
        var quoted = false;
        var bracket = false;
        for (var i = 0; i < code.Length; i++)
        {
            var ch = code[i];
            if (ch == '\\' && !quoted) { i++; continue; }
            if (ch == '"') quoted = !quoted;
            else if (!quoted && ch == '[') bracket = true;
            else if (!quoted && ch == ']') bracket = false;
            else if (!quoted && !bracket && ch == ';') { list.Add(code[start..i]); start = i + 1; }
        }
        list.Add(code[start..]);
        return list;
    }

    static string Date(double v, List<(char K, string V)> toks)
    {
        var s = (int)Math.Floor(v);
        var (y, mo, d) = s switch
        {
            0 => (1900, 1, 0),
            60 => (1900, 2, 29), // Excel's 1900-02-29
            _ => new DateTime(1899, 12, 30).AddDays(s < 60 ? s + 1 : s) is var dt ? (dt.Year, dt.Month, dt.Day) : default,
        };
        var t = (int)Math.Min(86399, Math.Round((v - s) * 86400));
        int h = t / 3600, mi = t / 60 % 60, se = t % 60, wd = ((s - 1) % 7 + 7) % 7;
        var ap = toks.Any(x => x.K == 'a');
        var sb = new StringBuilder();
        for (var i = 0; i < toks.Count; i++)
        {
            var (k, c) = toks[i];
            if (k == 'a')
            {
                sb.Append(c.StartsWith('上') ? (h < 12 ? "上午" : "下午") : c.Length == 3 ? (h < 12 ? "A" : "P") : (h < 12 ? "AM" : "PM"));
                continue;
            }
            if (k != 'd') { sb.Append(c); continue; }
            var h12 = ap ? (h % 12 == 0 ? 12 : h % 12) : h;
            switch (c[0])
            {
                case '[':
                    var secs = Math.Round(v * 86400);
                    sb.Append((c[1] == 'h' ? Math.Floor(secs / 3600) : c[1] == 'm' ? Math.Floor(secs / 60) : secs).ToString(Inv));
                    continue;
                case 'y': sb.Append(c.Length <= 2 ? (y % 100).ToString("00", Inv) : y.ToString(Inv)); continue;
                case 'd': sb.Append(c.Length switch { 1 => d.ToString(Inv), 2 => d.ToString("00", Inv), 3 => Days[wd][..3], _ => Days[wd] }); continue;
                case 'h': sb.Append(c.Length == 1 ? h12.ToString(Inv) : h12.ToString("00", Inv)); continue;
                case 's': sb.Append(c.Length == 1 ? se.ToString(Inv) : se.ToString("00", Inv)); continue;
                case 'a': sb.Append(c.Length >= 4 ? "星期" : "周").Append(ChineseDays[wd]); continue;
            }
            // m: minutes right after an hour or right before seconds, else the month
            var prev = toks.Take(i).LastOrDefault(x => x.K == 'd').V;
            var next = toks.Skip(i + 1).FirstOrDefault(x => x.K == 'd').V;
            if (prev is not null && (prev.StartsWith('h') || prev.StartsWith("[h")) || next is not null && next.StartsWith('s'))
                sb.Append(c.Length == 1 ? mi.ToString(Inv) : mi.ToString("00", Inv));
            else sb.Append(c.Length switch { 1 => mo.ToString(Inv), 2 => mo.ToString("00", Inv), 3 => Months[mo - 1][..3], 4 => Months[mo - 1], _ => Months[mo - 1][..1] });
        }
        return sb.ToString();
    }
}
