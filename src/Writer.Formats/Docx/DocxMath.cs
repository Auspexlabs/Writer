using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using Writer.Core;
using M = DocumentFormat.OpenXml.Math;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>An equation in a paragraph: Word's m:oMath (in an m:oMathPara when it stands on a line of its own), at a character offset of
/// the paragraph's text like a note's mark, read and written as LaTeX.</summary>
sealed class DocxEquation(W.Paragraph p, M.OfficeMath math) : Node
{
    public override string Kind => "equation";
    public override object Anchor => math;
    OpenXmlElement Holder => math.Parent is M.Paragraph para ? para : math;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string>
        {
            ["latex"] = OfficeMathTex.ToLatex(math),
            ["at"] = DocxFootnotes.OffsetOf(p, Holder).ToString(CultureInfo.InvariantCulture),
        };
        if (math.Parent is M.Paragraph) props["display"] = "true";
        return props;
    }

    public override string GetRaw() => Holder.OuterXml;

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "latex":
                var fresh = OfficeMathTex.ToOmml(value);
                math.RemoveAllChildren();
                foreach (var c in fresh.ChildElements.ToList()) { c.Remove(); math.Append(c); }
                break;
            case "display" when value == "true" && math.Parent is not M.Paragraph:
                var holder = new M.Paragraph();
                math.InsertBeforeSelf(holder);
                math.Remove();
                holder.Append(math);
                break;
            case "display" when value != "true" && math.Parent is M.Paragraph para:
                math.Remove();
                para.InsertBeforeSelf(math);
                para.Remove();
                break;
            case "at":
                var h = Take();
                DocxFootnotes.Place(p, h, int.Parse(value, CultureInfo.InvariantCulture));
                break;
        }
    }

    /// <summary>Out of the paragraph, the runs it split joined again.</summary>
    OpenXmlElement Take()
    {
        var h = Holder;
        var (previous, next) = (h.PreviousSibling(), h.NextSibling());
        h.Remove();
        DocxRuns.Rejoin(previous, next);
        return h;
    }

    public override void Remove() => Take();

    /// <summary>The paragraph's own equations (not those in its text boxes), in reading order.</summary>
    public static IEnumerable<Node> In(W.Paragraph p) =>
        p.Descendants<M.OfficeMath>().Where(m => !m.Ancestors<M.OfficeMath>().Any() && m.Ancestors<W.Paragraph>().First() == p).Select(m => (Node)new DocxEquation(p, m));

    public static Node Add(W.Paragraph p, IReadOnlyDictionary<string, string> props)
    {
        var math = OfficeMathTex.ToOmml(props.GetValueOrDefault("latex") ?? "");
        OpenXmlElement holder = props.GetValueOrDefault("display") == "true" ? new M.Paragraph(math) : math;
        DocxFootnotes.Place(p, holder, props.TryGetValue("at", out var at) ? int.Parse(at, CultureInfo.InvariantCulture) : int.MaxValue);
        return new DocxEquation(p, math);
    }
}

/// <summary>LaTeX ⇄ Office Math (OMML) for the common ground both share: fractions, scripts, roots, big operators with limits,
/// \left…\right delimiters, matrices and cases, accents, over/underlines and braces, function names, text and letter styles, Greek
/// letters and symbols. LaTeX outside it is kept as text; OMML outside it is read as its content. Writing then reading gives back
/// LaTeX that reads the same again.</summary>
static class OfficeMathTex
{
    const string MNs = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace Mx = MNs;

    // command → character; the first command for a character is the one read back
    static readonly (string Cmd, string Chr)[] SymbolList =
    [
        ("alpha", "α"), ("beta", "β"), ("gamma", "γ"), ("delta", "δ"), ("epsilon", "ϵ"), ("varepsilon", "ε"), ("zeta", "ζ"), ("eta", "η"), ("theta", "θ"), ("vartheta", "ϑ"),
        ("iota", "ι"), ("kappa", "κ"), ("lambda", "λ"), ("mu", "μ"), ("nu", "ν"), ("xi", "ξ"), ("pi", "π"), ("varpi", "ϖ"), ("rho", "ρ"), ("varrho", "ϱ"), ("sigma", "σ"),
        ("varsigma", "ς"), ("tau", "τ"), ("upsilon", "υ"), ("phi", "ϕ"), ("varphi", "φ"), ("chi", "χ"), ("psi", "ψ"), ("omega", "ω"),
        ("Gamma", "Γ"), ("Delta", "Δ"), ("Theta", "Θ"), ("Lambda", "Λ"), ("Xi", "Ξ"), ("Pi", "Π"), ("Sigma", "Σ"), ("Upsilon", "Υ"), ("Phi", "Φ"), ("Psi", "Ψ"), ("Omega", "Ω"),
        ("infty", "∞"), ("pm", "±"), ("mp", "∓"), ("times", "×"), ("div", "÷"), ("cdot", "⋅"), ("ast", "∗"), ("star", "⋆"), ("circ", "∘"), ("bullet", "∙"),
        ("leq", "≤"), ("le", "≤"), ("geq", "≥"), ("ge", "≥"), ("neq", "≠"), ("ne", "≠"), ("approx", "≈"), ("equiv", "≡"), ("sim", "∼"), ("simeq", "≃"), ("cong", "≅"),
        ("propto", "∝"), ("ll", "≪"), ("gg", "≫"), ("to", "→"), ("rightarrow", "→"), ("leftarrow", "←"), ("gets", "←"), ("leftrightarrow", "↔"), ("Rightarrow", "⇒"),
        ("Leftarrow", "⇐"), ("Leftrightarrow", "⇔"), ("implies", "⟹"), ("iff", "⟺"), ("mapsto", "↦"), ("uparrow", "↑"), ("downarrow", "↓"),
        ("in", "∈"), ("notin", "∉"), ("ni", "∋"), ("subset", "⊂"), ("supset", "⊃"), ("subseteq", "⊆"), ("supseteq", "⊇"), ("cup", "∪"), ("cap", "∩"),
        ("emptyset", "∅"), ("varnothing", "∅"), ("setminus", "∖"), ("forall", "∀"), ("exists", "∃"), ("nexists", "∄"), ("neg", "¬"), ("lnot", "¬"),
        ("land", "∧"), ("wedge", "∧"), ("lor", "∨"), ("vee", "∨"), ("oplus", "⊕"), ("otimes", "⊗"), ("partial", "∂"), ("nabla", "∇"), ("degree", "°"),
        ("angle", "∠"), ("perp", "⊥"), ("parallel", "∥"), ("mid", "∣"), ("triangle", "△"), ("ldots", "…"), ("dots", "…"), ("cdots", "⋯"), ("vdots", "⋮"),
        ("ddots", "⋱"), ("prime", "′"), ("hbar", "ℏ"), ("ell", "ℓ"), ("Re", "ℜ"), ("Im", "ℑ"), ("aleph", "ℵ"), ("therefore", "∴"), ("because", "∵"),
        ("langle", "⟨"), ("rangle", "⟩"), ("lfloor", "⌊"), ("rfloor", "⌋"), ("lceil", "⌈"), ("rceil", "⌉"), ("|", "‖"), ("lvert", "|"), ("rvert", "|"),
        ("{", "{"), ("}", "}"), ("%", "%"), ("$", "$"), ("#", "#"), ("&", "&"), ("_", "_"),
        (",", " "), (":", " "), (";", " "), (" ", " "), ("quad", " "), ("qquad", "  "), ("!", ""),
    ];
    static readonly Dictionary<string, string> Symbols = SymbolList.GroupBy(s => s.Cmd).ToDictionary(g => g.Key, g => g.First().Chr);
    static readonly Dictionary<string, string> Back = SymbolList.Where(s => s.Chr.Length == 1 && s.Cmd is not (" " or "lvert" or "rvert")).GroupBy(s => s.Chr).ToDictionary(g => g.Key, g => g.First().Cmd);
    static readonly Dictionary<string, string> Operators = new()
    {
        ["sum"] = "∑", ["prod"] = "∏", ["coprod"] = "∐", ["int"] = "∫", ["iint"] = "∬", ["iiint"] = "∭", ["oint"] = "∮", ["bigcup"] = "⋃", ["bigcap"] = "⋂",
        ["bigoplus"] = "⨁", ["bigotimes"] = "⨂", ["bigvee"] = "⋁", ["bigwedge"] = "⋀",
    };
    static readonly HashSet<string> Functions = ["sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh", "coth", "log", "ln", "lg", "exp",
        "ker", "dim", "deg", "arg", "hom", "det", "gcd", "lim", "liminf", "limsup", "max", "min", "sup", "inf", "Pr"];
    static readonly HashSet<string> Limits = ["lim", "liminf", "limsup", "max", "min", "sup", "inf", "det", "gcd", "Pr"]; // their subscript goes underneath
    static readonly (string Cmd, string Chr)[] AccentList = [("hat", "̂"), ("widehat", "̂"), ("check", "̌"), ("tilde", "̃"), ("widetilde", "̃"), ("acute", "́"),
        ("grave", "̀"), ("dot", "̇"), ("ddot", "̈"), ("breve", "̆"), ("bar", "̅"), ("vec", "⃗"), ("overrightarrow", "⃗")];
    static readonly Dictionary<string, string> Accents = AccentList.ToDictionary(a => a.Cmd, a => a.Chr);
    static readonly Dictionary<string, (string Open, string Close)> Matrices = new()
    {
        ["matrix"] = ("", ""), ["pmatrix"] = ("(", ")"), ["bmatrix"] = ("[", "]"), ["Bmatrix"] = ("{", "}"), ["vmatrix"] = ("|", "|"), ["Vmatrix"] = ("‖", "‖"), ["cases"] = ("{", ""),
    };
    static readonly Dictionary<string, string> Styles = new()
    {
        ["mathrm"] = "p", ["operatorname"] = "p", ["mathbf"] = "b", ["mathit"] = "i", ["boldsymbol"] = "bi", ["mathbb"] = "double-struck", ["mathcal"] = "script",
        ["mathscr"] = "script", ["mathfrak"] = "fraktur", ["mathsf"] = "sans-serif", ["mathtt"] = "monospace",
    };
    static readonly Dictionary<string, string> Delimiters = new() { ["("] = "(", [")"] = ")", ["["] = "[", ["]"] = "]", ["\\{"] = "{", ["\\}"] = "}", ["|"] = "|", ["\\|"] = "‖",
        ["\\langle"] = "⟨", ["\\rangle"] = "⟩", ["\\lfloor"] = "⌊", ["\\rfloor"] = "⌋", ["\\lceil"] = "⌈", ["\\rceil"] = "⌉", ["."] = "", ["\\lvert"] = "|", ["\\rvert"] = "|", ["\\vert"] = "|" };

    // ---- LaTeX → OMML ----

    public static M.OfficeMath ToOmml(string latex) =>
        new($"<m:oMath xmlns:m=\"{MNs}\" xmlns:w=\"{WNs}\">{new Reader(latex).All()}</m:oMath>");

    static string X(string s) => s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    /// <summary>A math run: its style (p plain, b bold, i, bi, nor for text, or a script such as double-struck) and Word's math font.</summary>
    static string Run(string text, string? style = null)
    {
        var rPr = style switch
        {
            null => "",
            "nor" => "<m:rPr><m:nor/></m:rPr>",
            "p" or "b" or "i" or "bi" => $"<m:rPr><m:sty m:val=\"{style}\"/></m:rPr>",
            _ => $"<m:rPr><m:scr m:val=\"{style}\"/><m:sty m:val=\"p\"/></m:rPr>",
        };
        return $"<m:r>{rPr}<w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr><m:t xml:space=\"preserve\">{X(text)}</m:t></m:r>";
    }

    static string E(string inner, string tag = "e") => inner.Length == 0 ? $"<m:{tag}/>" : $"<m:{tag}>{inner}</m:{tag}>";
    static string Chr(string tag, string c) => $"<m:{tag} m:val=\"{X(c)}\"/>";
    static string Delimited(string open, string close, string inner) => $"<m:d><m:dPr>{Chr("begChr", open)}{Chr("endChr", close)}</m:dPr>{E(inner)}</m:d>";

    sealed class Reader(string s)
    {
        int i;
        bool End => i >= s.Length;
        char Peek => s[i];
        void Ws() { while (!End && char.IsWhiteSpace(Peek)) i++; }
        bool AtCmd(string name) => string.CompareOrdinal(s, i, "\\" + name, 0, name.Length + 1) == 0 && (i + name.Length + 1 >= s.Length || !char.IsLetter(s[i + name.Length + 1]));

        /// <summary>The whole formula; a stray closing brace is skipped.</summary>
        public string All()
        {
            var sb = new StringBuilder(Expr());
            while (!End) { i++; sb.Append(Expr()); }
            return sb.ToString();
        }

        /// <summary>Items up to a closing brace, a cell or row end, \right or \end; plain characters in a row go in one run.</summary>
        string Expr()
        {
            var sb = new StringBuilder();
            var text = new StringBuilder();
            void Flush() { if (text.Length > 0) { sb.Append(Run(text.ToString())); text.Clear(); } }
            while (true)
            {
                Ws();
                if (End || Peek == '}' || Peek == '&' || (Peek == '\\' && i + 1 < s.Length && s[i + 1] == '\\') || AtCmd("right") || AtCmd("end")) break;
                var item = Atom();
                var (sub, sup) = Scripts();
                if (item.Nary is { } op)
                {
                    Flush();
                    Ws();
                    var operand = End || Peek is '}' or '&' || AtCmd("right") || AtCmd("end") || (Peek == '\\' && i + 1 < s.Length && s[i + 1] == '\\') ? "" : Item();
                    var integral = op is "∫" or "∬" or "∭" or "∮";
                    sb.Append($"<m:nary><m:naryPr>{Chr("chr", op)}{Chr("limLoc", integral ? "subSup" : "undOvr")}{(sub is null ? "<m:subHide m:val=\"1\"/>" : "")}{(sup is null ? "<m:supHide m:val=\"1\"/>" : "")}</m:naryPr>{E(sub ?? "", "sub")}{E(sup ?? "", "sup")}{E(operand)}</m:nary>");
                    continue;
                }
                if (item.Limit is { } name && sub is not null)
                {
                    Flush();
                    var under = $"<m:limLow><m:e>{Run(name, "p")}</m:e>{E(sub, "lim")}</m:limLow>";
                    sb.Append(sup is null ? under : $"<m:sSup><m:e>{under}</m:e>{E(sup, "sup")}</m:sSup>");
                    continue;
                }
                if (sub is null && sup is null && item.Plain is { } plain) { text.Append(plain); continue; }
                Flush();
                var b = item.Plain is { } one ? Run(one) : item.Xml;
                sb.Append(sub is null && sup is null ? b
                    : sub is null ? $"<m:sSup>{E(b)}{E(sup!, "sup")}</m:sSup>"
                    : sup is null ? $"<m:sSub>{E(b)}{E(sub, "sub")}</m:sSub>"
                    : $"<m:sSubSup>{E(b)}{E(sub, "sub")}{E(sup, "sup")}</m:sSubSup>");
            }
            Flush();
            return sb.ToString();
        }

        /// <summary>One item with its scripts, as OMML (a big operator's operand).</summary>
        string Item()
        {
            var item = Atom();
            var (sub, sup) = Scripts();
            var b = item.Plain is { } one ? Run(one) : item.Xml;
            return sub is null && sup is null ? b
                : sub is null ? $"<m:sSup>{E(b)}{E(sup!, "sup")}</m:sSup>"
                : sup is null ? $"<m:sSub>{E(b)}{E(sub, "sub")}</m:sSub>"
                : $"<m:sSubSup>{E(b)}{E(sub, "sub")}{E(sup, "sup")}</m:sSubSup>";
        }

        (string? Sub, string? Sup) Scripts()
        {
            string? sub = null, sup = null;
            while (true)
            {
                Ws();
                if (End) break;
                if (Peek == '^') { i++; sup = Arg(); }
                else if (Peek == '_') { i++; sub = Arg(); }
                else if (Peek == '\'') { i++; sup = (sup ?? "") + Run("′"); }
                else break;
            }
            return (sub, sup);
        }

        /// <summary>A command's or a script's argument: a braced group or one item.</summary>
        string Arg()
        {
            Ws();
            if (End) return "";
            if (Peek == '{') { i++; var x = Expr(); if (!End && Peek == '}') i++; return x; }
            var item = Atom();
            return item.Plain is { } p ? Run(p) : item.Xml;
        }

        /// <summary>The raw text of a braced argument (\text{…}), braces inside kept balanced.</summary>
        string Raw()
        {
            Ws();
            if (End || Peek != '{') return End ? "" : s[i++].ToString();
            var depth = 0;
            var start = ++i;
            for (; !End; i++)
            {
                if (Peek == '\\') { i++; continue; }
                if (Peek == '{') depth++;
                else if (Peek == '}' && depth-- == 0) break;
            }
            var raw = s[start..Math.Min(i, s.Length)];
            if (!End) i++;
            return raw;
        }

        string Command()
        {
            i++; // the backslash
            if (End) return "";
            var start = i;
            if (char.IsLetter(Peek)) while (!End && char.IsLetter(Peek)) i++;
            else i++;
            return s[start..i];
        }

        string Delimiter()
        {
            Ws();
            if (End) return "";
            if (Peek == '\\')
            {
                var start = i;
                var name = "\\" + Command();
                return Delimiters.TryGetValue(name, out var d) ? d : Symbols.GetValueOrDefault(name[1..]) ?? s[start..i];
            }
            var c = s[i++].ToString();
            return Delimiters.GetValueOrDefault(c) ?? c;
        }

        record struct Piece(string Xml, string? Plain = null, string? Nary = null, string? Limit = null);

        Piece Atom()
        {
            var c = Peek;
            if (c == '{') { i++; var x = Expr(); if (!End && Peek == '}') i++; return new(x); }
            if (c is '^' or '_') return new("");
            if (c == '\'') { i++; return new("", "′"); }
            if (c != '\\') { var len = char.IsHighSurrogate(c) && i + 1 < s.Length ? 2 : 1; var t = s.Substring(i, len); i += len; return new("", t); }
            var name = Command();
            if (Symbols.TryGetValue(name, out var sym)) return new("", sym);
            if (Operators.TryGetValue(name, out var op)) return new("", Nary: op);
            if (Functions.Contains(name)) { Ws(); return Limits.Contains(name) && !End && Peek == '_' ? new("", Limit: name) : new(Run(name, "p")); }
            if (Accents.TryGetValue(name, out var acc)) return new($"<m:acc><m:accPr>{Chr("chr", acc)}</m:accPr>{E(Arg())}</m:acc>");
            if (Styles.TryGetValue(name, out var style)) return new(Run(Raw().Replace(" ", ""), style));
            switch (name)
            {
                case "frac" or "dfrac" or "tfrac" or "cfrac":
                    var num = Arg();
                    return new($"<m:f>{E(num, "num")}{E(Arg(), "den")}</m:f>");
                case "binom" or "dbinom" or "tbinom":
                    var n = Arg();
                    return new(Delimited("(", ")", $"<m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr>{E(n, "num")}{E(Arg(), "den")}</m:f>"));
                case "sqrt":
                    Ws();
                    string? degree = null;
                    if (!End && Peek == '[')
                    {
                        var close = s.IndexOf(']', i);
                        if (close < 0) close = s.Length;
                        degree = new Reader(s[(i + 1)..close]).All();
                        i = Math.Min(close + 1, s.Length);
                    }
                    return new($"<m:rad>{(degree is null ? "<m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/>" : E(degree, "deg"))}{E(Arg())}</m:rad>");
                case "overline" or "underline":
                    return new($"<m:bar><m:barPr><m:pos m:val=\"{(name == "overline" ? "top" : "bot")}\"/></m:barPr>{E(Arg())}</m:bar>");
                case "overbrace" or "underbrace":
                    var top = name == "overbrace";
                    return new($"<m:groupChr><m:groupChrPr>{Chr("chr", top ? "⏞" : "⏟")}<m:pos m:val=\"{(top ? "top" : "bot")}\"/><m:vertJc m:val=\"{(top ? "bot" : "top")}\"/></m:groupChrPr>{E(Arg())}</m:groupChr>");
                case "text" or "textrm" or "mbox" or "textit" or "textbf":
                    return new(Run(Raw(), "nor"));
                case "left":
                    var open = Delimiter();
                    var inner = Expr();
                    var closing = "";
                    if (AtCmd("right")) { Command(); closing = Delimiter(); }
                    return new(Delimited(open, closing, inner));
                case "begin":
                    return new(Environment(Raw().Trim()));
                case "displaystyle" or "textstyle" or "limits" or "nolimits" or "big" or "Big" or "bigg" or "Bigg" or "bigl" or "bigr" or "Bigl" or "Bigr":
                    return new("");
                default:
                    return new(Run(name, "p")); // a command neither side knows: its name as text
            }
        }

        /// <summary>\begin{…} … \end{…}: matrices (in their brackets), cases, and aligned rows as an equation array.</summary>
        string Environment(string env)
        {
            var rows = new List<List<string>> { new() };
            while (true)
            {
                rows[^1].Add(Expr());
                if (End) break;
                if (Peek == '&') { i++; continue; }
                if (Peek == '\\' && i + 1 < s.Length && s[i + 1] == '\\') { i += 2; rows.Add([]); continue; }
                if (AtCmd("end")) { Command(); Raw(); break; }
                if (Peek == '}') { i++; continue; }
                break;
            }
            if (rows.Count > 1 && rows[^1].All(c => c.Length == 0)) rows.RemoveAt(rows.Count - 1);
            if (!Matrices.TryGetValue(env, out var brackets))
                return "<m:eqArr>" + string.Concat(rows.Select(r => E(string.Concat(r)))) + "</m:eqArr>";
            var cols = rows.Max(r => r.Count);
            var m = $"<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val=\"{cols}\"/><m:mcJc m:val=\"{(env == "cases" ? "left" : "center")}\"/></m:mcPr></m:mc></m:mcs></m:mPr>"
                + string.Concat(rows.Select(r => "<m:mr>" + string.Concat(Enumerable.Range(0, cols).Select(k => E(k < r.Count ? r[k] : ""))) + "</m:mr>")) + "</m:m>";
            return brackets.Open.Length == 0 && brackets.Close.Length == 0 ? m : Delimited(brackets.Open, brackets.Close, m);
        }
    }

    // ---- OMML → LaTeX ----

    public static string ToLatex(M.OfficeMath math) => Kids(XElement.Parse(math.OuterXml)).Trim();

    static string Kids(XElement? e) => e is null ? "" : Join(e.Elements().Where(x => x.Name.Namespace == Mx && !x.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)).Select(Tex));

    /// <summary>Pieces side by side, a space where a command word would run into a letter.</summary>
    static string Join(IEnumerable<string> parts)
    {
        var sb = new StringBuilder();
        foreach (var part in parts)
        {
            if (part.Length == 0) continue;
            if (sb.Length > 0 && char.IsLetter(part[0]) && EndsInCommand(sb)) sb.Append(' ');
            sb.Append(part);
        }
        return sb.ToString();
    }

    static bool EndsInCommand(StringBuilder sb)
    {
        var k = sb.Length - 1;
        while (k >= 0 && char.IsLetter(sb[k])) k--;
        return k >= 0 && k < sb.Length - 1 && sb[k] == '\\';
    }

    static XElement? Child(XElement e, string name) => e.Element(Mx + name);
    static string? Val(XElement? e, string pr, string name) => e?.Element(Mx + pr)?.Element(Mx + name)?.Attribute(Mx + "val")?.Value;
    static bool On(XElement e, string pr, string name) => e.Element(Mx + pr)?.Element(Mx + name) is { } x && (x.Attribute(Mx + "val")?.Value ?? "1") is "1" or "on" or "true";
    static string Group(XElement? e) => "{" + Kids(e) + "}";

    /// <summary>A script's base: bare when it is one letter or one command, else braced.</summary>
    static string Base(XElement? e)
    {
        var t = Kids(e);
        return t.Length == 1 || (t.Length > 1 && t[0] == '\\' && t.Skip(1).All(char.IsLetter)) ? t : "{" + t + "}";
    }

    static string Tex(XElement e)
    {
        switch (e.Name.LocalName)
        {
            case "r": return RunTex(e);
            case "f":
                return $"\\frac{Group(Child(e, "num"))}{Group(Child(e, "den"))}"; // a stacked fraction without its bar reads as \binom in its parentheses
            case "sSup": return Base(Child(e, "e")) + "^" + Group(Child(e, "sup"));
            case "sSub": return Base(Child(e, "e")) + "_" + Group(Child(e, "sub"));
            case "sSubSup": return Base(Child(e, "e")) + "_" + Group(Child(e, "sub")) + "^" + Group(Child(e, "sup"));
            case "sPre": return "{}_" + Group(Child(e, "sub")) + "^" + Group(Child(e, "sup")) + Base(Child(e, "e"));
            case "rad":
                return On(e, "radPr", "degHide") || Kids(Child(e, "deg")).Length == 0 ? "\\sqrt" + Group(Child(e, "e")) : $"\\sqrt[{Kids(Child(e, "deg"))}]" + Group(Child(e, "e"));
            case "nary":
                var chr = Val(e, "naryPr", "chr") ?? "∫";
                var op = Operators.FirstOrDefault(o => o.Value == chr).Key;
                var sb = new StringBuilder(op is null ? chr : "\\" + op);
                if (!On(e, "naryPr", "subHide") && Kids(Child(e, "sub")) is { Length: > 0 } sub) sb.Append("_{").Append(sub).Append('}');
                if (!On(e, "naryPr", "supHide") && Kids(Child(e, "sup")) is { Length: > 0 } sup) sb.Append("^{").Append(sup).Append('}');
                if (Kids(Child(e, "e")) is { Length: > 0 } operand) sb.Append('{').Append(operand).Append('}');
                return sb.ToString();
            case "d": return DelimiterTex(e);
            case "m": return "\\begin{matrix}" + MatrixRows(e) + "\\end{matrix}";
            case "eqArr": return "\\begin{aligned}" + string.Join("\\\\", e.Elements(Mx + "e").Select(Kids)) + "\\end{aligned}";
            case "acc":
                var mark = Val(e, "accPr", "chr") ?? "̂";
                return "\\" + (AccentList.FirstOrDefault(a => a.Chr == mark).Cmd ?? "hat") + Group(Child(e, "e"));
            case "bar": return (Val(e, "barPr", "pos") == "top" ? "\\overline" : "\\underline") + Group(Child(e, "e"));
            case "groupChr": return ((Val(e, "groupChrPr", "chr") ?? "⏟") == "⏞" ? "\\overbrace" : "\\underbrace") + Group(Child(e, "e"));
            case "limLow": return Base(Child(e, "e")) + "_" + Group(Child(e, "lim"));
            case "limUpp": return Base(Child(e, "e")) + "^" + Group(Child(e, "lim"));
            case "func": return Join([Kids(Child(e, "fName")), Base(Child(e, "e"))]);
            default: return Kids(e); // e, box, borderBox, phant, and what else holds content
        }
    }

    static string MatrixRows(XElement m) => string.Join("\\\\", m.Elements(Mx + "mr").Select(r => string.Join("&", r.Elements(Mx + "e").Select(Kids))));

    static string DelimiterTex(XElement d)
    {
        var open = Val(d, "dPr", "begChr") ?? "(";
        var close = Val(d, "dPr", "endChr") ?? ")";
        var parts = d.Elements(Mx + "e").ToList();
        if (open == "(" && close == ")" && parts.Count == 1 && parts[0].Elements().Where(x => x.Name.Namespace == Mx).ToList() is [{ Name.LocalName: "f" } f] && Val(f, "fPr", "type") == "noBar")
            return $"\\binom{Group(Child(f, "num"))}{Group(Child(f, "den"))}";
        if (parts.Count == 1 && parts[0].Elements().Where(x => x.Name.Namespace == Mx).ToList() is [{ Name.LocalName: "m" } matrix])
        {
            var env = Matrices.FirstOrDefault(x => x.Value.Open == open && x.Value.Close == close && x.Key != "matrix").Key;
            if (env is not null) return $"\\begin{{{env}}}" + MatrixRows(matrix) + $"\\end{{{env}}}";
        }
        var sep = Val(d, "dPr", "sepChr") ?? "|";
        return "\\left" + DelimTex(open) + string.Join(sep == "|" ? "|" : sep, parts.Select(Kids)) + "\\right" + DelimTex(close);
    }

    static string DelimTex(string c) => c switch { "" => ".", "{" => "\\{", "}" => "\\}", "‖" => "\\|", _ => Back.TryGetValue(c, out var cmd) && c is not ("(" or ")" or "[" or "]" or "|") ? "\\" + cmd : c };

    static string RunTex(XElement r)
    {
        var text = string.Concat(r.Elements(Mx + "t").Select(t => t.Value));
        if (text.Length == 0) return "";
        var rPr = r.Element(Mx + "rPr");
        if (rPr?.Element(Mx + "nor") is not null) return "\\text{" + text + "}";
        var scr = rPr?.Element(Mx + "scr")?.Attribute(Mx + "val")?.Value;
        var sty = rPr?.Element(Mx + "sty")?.Attribute(Mx + "val")?.Value;
        if (scr is not null && scr != "roman" && Styles.FirstOrDefault(x => x.Value == scr).Key is { } font) return "\\" + font + "{" + text + "}";
        if (sty == "p" && Functions.Contains(text)) return "\\" + text;
        if (sty is "p" or "b" or "bi" && text.Any(char.IsLetter)) return "\\" + Styles.First(x => x.Value == sty).Key + "{" + text + "}";
        return Join(text.EnumerateRunes().Select(rune =>
        {
            var c = rune.ToString();
            if (Back.TryGetValue(c, out var cmd)) return "\\" + cmd;
            return c switch { "^" => "\\^{}", "~" => "\\sim", "\\" => "\\backslash", _ => c };
        }));
    }
}
