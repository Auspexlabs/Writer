using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>Page setup and the headers and footers of the last section: paper, orientation, margins, columns, the default
/// header and footer, and the first page's own (titlePg). Header and footer html may carry {page} and {pages}, which become PAGE
/// and NUMPAGES fields.</summary>
static partial class DocxSection
{
    const int Twip = 635; // EMU per twentieth of a point
    static readonly (string? Name, uint W, uint H)[] Papers = [("A4", 11906, 16838), ("Letter", 12240, 15840), ("Legal", 12240, 20160), ("A3", 16838, 23811), ("A5", 8391, 11906), ("B5", 10319, 14571)];
    static readonly (string? Name, uint Top, uint Side)[] Margins = [("normal", 1440, 1440), ("narrow", 720, 720), ("moderate", 1440, 1080), ("wide", 1440, 2880)];

    static W.SectionProperties Section(DocxDocument doc) =>
        doc.Main.Document!.Body!.GetFirstChild<W.SectionProperties>() ?? doc.Main.Document.Body.AppendChild(new W.SectionProperties());

    public static void Read(DocxDocument doc, Dictionary<string, string> props)
    {
        var section = doc.Main.Document!.Body!.GetFirstChild<W.SectionProperties>();
        if (section is null) return;
        ReadPage(section, props);
        foreach (var (name, header, first) in Slots)
            if (Part(doc, section, header, first) is { } part && Html(doc, part) is { Length: > 0 } html) props[name] = html;
        if (DocxRun.On(section.GetFirstChild<W.TitlePage>())) props["titlePg"] = "true";
        if (section.GetFirstChild<W.LineNumberType>() is not null) props["lineNumbers"] = "true";
    }

    /// <summary>Paper, orientation, margins and columns of one section: the document's last (the body's sectPr) or one a paragraph ends.</summary>
    public static void ReadPage(W.SectionProperties section, Dictionary<string, string> props)
    {
        if (section.GetFirstChild<W.PageSize>() is { } size && size.Width?.Value is { } w && size.Height?.Value is { } h)
        {
            var landscape = IsLandscape(size);
            var (pw, ph) = landscape ? (h, w) : (w, h);
            var paper = Papers.FirstOrDefault(p => Near(p.W, pw) && Near(p.H, ph));
            props["page"] = paper.Name ?? $"{Cm(pw)} x {Cm(ph)}";
            props["orientation"] = landscape ? "landscape" : "portrait";
        }
        if (section.GetFirstChild<W.PageMargin>() is { } m && m.Top?.Value is { } top && m.Left?.Value is { } left)
        {
            long bottom = m.Bottom?.Value ?? top;
            long right = m.Right?.Value ?? left;
            var preset = Margins.FirstOrDefault(x => Near(x.Top, top) && Near(x.Side, right) && Near(x.Top, bottom) && Near(x.Side, left));
            props["margin"] = preset.Name ?? $"{Cm(top)} {Cm(right)} {Cm(bottom)} {Cm(left)}";
        }
        props["columns"] = (section.GetFirstChild<W.Columns>()?.ColumnCount?.Value ?? 1).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>A section's page as lengths, whatever preset its props name, for drawing it: pageSize (width x height as laid out),
    /// pageMargin (top right bottom left), headerDistance and footerDistance from the page's edges, columnGap.</summary>
    public static void ReadGeometry(W.SectionProperties section, Dictionary<string, string> props)
    {
        if (section.GetFirstChild<W.PageSize>() is { } size && size.Width?.Value is { } w && size.Height?.Value is { } h) props["pageSize"] = $"{Cm(w)} x {Cm(h)}";
        if (section.GetFirstChild<W.PageMargin>() is { } m)
        {
            // a negative top or bottom margin lets the text run under the header or footer: the text still starts that far in
            props["pageMargin"] = $"{Cm(Math.Abs(m.Top?.Value ?? 1440))} {Cm(m.Right?.Value ?? 1440)} {Cm(Math.Abs(m.Bottom?.Value ?? 1440))} {Cm(m.Left?.Value ?? 1440)}";
            if (m.Header?.Value is { } header) props["headerDistance"] = Cm(header);
            if (m.Footer?.Value is { } footer) props["footerDistance"] = Cm(footer);
        }
        if (section.GetFirstChild<W.Columns>() is { } columns && columns.ColumnCount?.Value > 1)
            props["columnGap"] = Cm(long.TryParse(columns.Space?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var space) ? space : 720);
    }

    /// <summary>Line numbers in the margin, every line, counted through the document (Word's 行号 › 连续).</summary>
    public static void SetLineNumbers(DocxDocument doc, bool on)
    {
        var section = Section(doc);
        section.RemoveAllChildren<W.LineNumberType>();
        if (on) section.AddChild(new W.LineNumberType { CountBy = 1, Restart = W.LineNumberRestartValues.Continuous });
    }

    /// <summary>The section a paragraph ends (w:pPr/w:sectPr), made when missing as a copy of the document's last section (its paper,
    /// margins, columns, headers and footers), so the text before the break keeps its look; the type says how the next section starts.</summary>
    public static void SetSectionBreak(DocxDocument doc, W.Paragraph p, string type)
    {
        var pp = p.ParagraphProperties ??= new W.ParagraphProperties();
        if (type is "none" or "")
        {
            pp.SectionProperties?.Remove();
            if (!pp.HasChildren) pp.Remove();
            return;
        }
        var section = pp.SectionProperties ??= (W.SectionProperties)Section(doc).CloneNode(true);
        section.RemoveAllChildren<W.SectionType>();
        var kind = type switch
        {
            "continuous" => W.SectionMarkValues.Continuous,
            "evenPage" => W.SectionMarkValues.EvenPage,
            "oddPage" => W.SectionMarkValues.OddPage,
            _ => W.SectionMarkValues.NextPage,
        };
        section.AddChild(new W.SectionType { Val = kind });
    }

    public static string? SectionBreakOf(W.Paragraph p) => p.ParagraphProperties?.SectionProperties is { } section
        ? section.GetFirstChild<W.SectionType>()?.Val?.InnerText switch { "continuous" => "continuous", "evenPage" => "evenPage", "oddPage" => "oddPage", _ => "nextPage" }
        : null;

    /// <summary>The document props for the section's header and footer references.</summary>
    static readonly (string Name, bool Header, bool First)[] Slots = [("header", true, false), ("footer", false, false), ("firstHeader", true, true), ("firstFooter", false, true)];

    /// <summary>titlePg: the first page shows the first-page header and footer (empty when there are none).</summary>
    public static void SetTitlePage(DocxDocument doc, bool on)
    {
        var section = Section(doc);
        section.RemoveAllChildren<W.TitlePage>();
        if (on) section.AddChild(new W.TitlePage());
    }

    static bool Near(long a, long b) => Math.Abs(a - b) <= 20;
    static string Cm(long twips) => (twips * Twip / 360000.0).ToString("0.##", CultureInfo.InvariantCulture) + "cm";
    static bool IsLandscape(W.PageSize size) => size.Orient?.Value == W.PageOrientationValues.Landscape || size.Width?.Value > size.Height?.Value;

    public static void SetPage(DocxDocument doc, string value) => SetPage(Section(doc), value);

    public static void SetPage(W.SectionProperties section, string value)
    {
        var size = PageSize(section);
        uint w, h;
        var paper = Papers.FirstOrDefault(p => string.Equals(p.Name, value.Trim(), StringComparison.OrdinalIgnoreCase));
        if (paper.Name is not null) (w, h) = (paper.W, paper.H);
        else
        {
            var parts = value.Split(['x', 'X', '×', '*'], 2);
            if (parts.Length != 2 || !value.Any(char.IsDigit))
                throw new WriterException(ErrorCode.Validation, $"Unknown paper '{value}'", "Use A3, A4, A5, B5, Letter, Legal, or a size like 21cm x 29.7cm.");
            (w, h) = (Twips(parts[0]), Twips(parts[1]));
        }
        if (IsLandscape(size)) (w, h) = (h, w);
        size.Width = w;
        size.Height = h;
    }

    public static void SetOrientation(DocxDocument doc, string value) => SetOrientation(Section(doc), value);

    public static void SetOrientation(W.SectionProperties section, string value)
    {
        var size = PageSize(section);
        var landscape = value == "landscape";
        if (landscape != IsLandscape(size))
        {
            var (w, h) = (size.Width?.Value ?? 11906U, size.Height?.Value ?? 16838U);
            size.Width = h;
            size.Height = w;
        }
        size.Orient = landscape ? W.PageOrientationValues.Landscape : W.PageOrientationValues.Portrait;
    }

    public static void SetMargin(DocxDocument doc, string value) => SetMargin(Section(doc), value);

    public static void SetMargin(W.SectionProperties section, string value)
    {
        var m = section.GetFirstChild<W.PageMargin>()
            ?? Add(section, new W.PageMargin { Top = 1440, Right = 1440U, Bottom = 1440, Left = 1440U, Header = 720U, Footer = 720U, Gutter = 0U });
        var preset = Margins.FirstOrDefault(x => string.Equals(x.Name, value.Trim(), StringComparison.OrdinalIgnoreCase));
        uint[] sides;
        if (preset.Name is not null) sides = [preset.Top, preset.Side, preset.Top, preset.Side];
        else
        {
            var parts = value.Any(char.IsDigit) ? value.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries).Select(Twips).ToArray() : [];
            sides = parts.Length switch
            {
                1 => [parts[0], parts[0], parts[0], parts[0]],
                2 => [parts[0], parts[1], parts[0], parts[1]],
                4 => parts,
                _ => throw new WriterException(ErrorCode.Validation, $"Cannot read margins '{value}'", "Use narrow, normal, moderate, wide, one length, or top right bottom left."),
            };
        }
        m.Top = (int)sides[0];
        m.Right = sides[1];
        m.Bottom = (int)sides[2];
        m.Left = sides[3];
    }

    public static void SetColumns(DocxDocument doc, string value) => SetColumns(Section(doc), value);

    public static void SetColumns(W.SectionProperties section, string value)
    {
        var n = int.Parse(value, CultureInfo.InvariantCulture);
        var cols = section.GetFirstChild<W.Columns>();
        if (n <= 1)
        {
            cols?.Remove();
            return;
        }
        cols ??= Add(section, new W.Columns { Space = "720" });
        cols.RemoveAllChildren();
        cols.EqualWidth = null;
        cols.ColumnCount = (short)n;
    }

    static W.PageSize PageSize(W.SectionProperties section) =>
        section.GetFirstChild<W.PageSize>() ?? Add(section, new W.PageSize { Width = 11906U, Height = 16838U });

    /// <summary>Adds a singleton child (page size, margins, columns) where the schema puts it.</summary>
    static T Add<T>(W.SectionProperties section, T child) where T : OpenXmlElement
    {
        section.AddChild(child);
        return child;
    }

    /// <summary>References go first, headers before footers; AddChild would replace the other reference in the same slot.</summary>
    static void AddReference(W.SectionProperties section, W.HeaderFooterReferenceType reference)
    {
        var next = section.ChildElements.FirstOrDefault(c => c is not W.HeaderFooterReferenceType || (reference is W.HeaderReference && c is W.FooterReference));
        if (next is null) section.Append(reference);
        else section.InsertBefore(reference, next);
    }

    static uint Twips(string length)
    {
        var emu = Units.ParseLength(length);
        if (emu < 0) throw new WriterException(ErrorCode.Validation, $"'{length}' is negative", "Use a positive length like 2cm.");
        return (uint)Math.Round(emu / (double)Twip);
    }

    static W.HeaderFooterReferenceType? Reference(W.SectionProperties section, bool header, bool first) =>
        section.Elements<W.HeaderFooterReferenceType>().FirstOrDefault(r => (r is W.HeaderReference) == header
            && (r.Type?.Value ?? W.HeaderFooterValues.Default) == (first ? W.HeaderFooterValues.First : W.HeaderFooterValues.Default));

    static OpenXmlPart? Part(DocxDocument doc, W.SectionProperties section, bool header, bool first) =>
        Reference(section, header, first)?.Id?.Value is { } id && doc.Main.TryGetPartById(id, out var part) ? part : null;

    const char Keep = '\uE000', KeepEnd = '\uE001'; // a placeholder's number in the text while html becomes runs

    /// <summary>Html of a header or footer: lines joined by br, or one p per paragraph with its alignment when any is not left;
    /// fields as {page} and {pages}; what html cannot say (pictures, other fields, links, content controls) as numbered read-only
    /// placeholders — img data-keep for a picture, span data-keep with its text otherwise — that a write puts back in place.</summary>
    static string Html(DocxDocument doc, OpenXmlPart part)
    {
        OpenXmlCompositeElement? container = part is HeaderPart hp ? hp.Header : (part as FooterPart)?.Footer;
        if (container is null) return "";
        var lines = new List<(string? Align, string Html)>();
        var keep = 0;
        foreach (var p in container.Elements<W.Paragraph>())
        {
            var html = new StringBuilder();
            var runs = new List<RunSpec>();
            foreach (var (spec, group) in Items(doc, p))
            {
                if (spec is null)
                {
                    html.Append(InlineHtml.Render(runs)).Append(Placeholder(part, group!, keep++));
                    runs.Clear();
                }
                else if (runs.Count > 0 && runs[^1] with { Text = "" } == spec with { Text = "" }) runs[^1] = runs[^1] with { Text = runs[^1].Text + spec.Text };
                else runs.Add(spec);
            }
            html.Append(InlineHtml.Render(runs));
            lines.Add((DocxParagraph.AlignOf(p.ParagraphProperties?.Justification ?? doc.Styles.Inherited(p, s => s.Justification)), html.ToString()));
        }
        while (lines.Count > 0 && lines[^1].Html.Length == 0) lines.RemoveAt(lines.Count - 1);
        return lines.Any(l => l.Align is not (null or "left"))
            ? string.Concat(lines.Select(l => $"<p style=\"text-align:{l.Align ?? "left"}\">{l.Html}</p>"))
            : string.Join("<br>", lines.Select(l => l.Html));
    }

    /// <summary>A paragraph's content in order: text runs, and page fields as {page} / {pages} with the field's formatting, as run
    /// specs; anything else as a group of elements kept whole.</summary>
    static IEnumerable<(RunSpec? Spec, List<OpenXmlElement>? Keep)> Items(DocxDocument doc, W.Paragraph p)
    {
        List<OpenXmlElement>? field = null;
        var depth = 0;
        foreach (var child in p.ChildElements)
        {
            var fc = (child as W.Run)?.GetFirstChild<W.FieldChar>()?.FieldCharType?.Value;
            if (field is not null || fc == W.FieldCharValues.Begin)
            {
                (field ??= []).Add(child);
                if (fc == W.FieldCharValues.Begin) depth++;
                else if (fc == W.FieldCharValues.End && --depth == 0)
                {
                    var token = Token(string.Concat(field.SelectMany(e => e.Descendants<W.FieldCode>()).Select(c => c.Text)));
                    yield return token.Length > 0 ? (Spec(doc, field.OfType<W.Run>().FirstOrDefault(r => r.GetFirstChild<W.FieldCode>() is not null), token), null) : (null, field);
                    field = null;
                }
                continue;
            }
            switch (child)
            {
                case W.ParagraphProperties or W.BookmarkStart or W.BookmarkEnd or W.ProofError:
                    break;
                case W.Run r when r.ChildElements.All(e => e is W.RunProperties or W.LastRenderedPageBreak || DocxRuns.IsTextElement(e)):
                    if (DocxRuns.HasText(r)) yield return (RunSpec.FromProps(new DocxRun(doc, r, null).GetProps()), null);
                    break;
                case W.SimpleField f when Token(f.Instruction?.Value) is { Length: > 0 } token:
                    yield return (Spec(doc, f.GetFirstChild<W.Run>(), token), null);
                    break;
                default:
                    yield return (null, [child]);
                    break;
            }
        }
        if (field is not null) yield return (null, field);
    }

    static RunSpec Spec(DocxDocument doc, W.Run? run, string text) => (run is null ? new RunSpec("") : RunSpec.FromProps(new DocxRun(doc, run, null).GetProps())) with { Text = text };

    static string Token(string? instruction) => (instruction ?? "").Trim().Split(' ', 2)[0].ToUpperInvariant() switch
    {
        "NUMPAGES" => "{pages}",
        "PAGE" => "{page}",
        _ => "",
    };

    static readonly string[] Pictures = ["image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp"];

    /// <summary>A kept group as the editor shows it: its picture, else its text.</summary>
    static string Placeholder(OpenXmlPart part, List<OpenXmlElement> group, int k)
    {
        if (group.SelectMany(e => e.Descendants<A.Blip>()).FirstOrDefault()?.Embed?.Value is { } id
            && part.TryGetPartById(id, out var target) && target is ImagePart image && Pictures.Contains(image.ContentType))
        {
            using var stream = image.GetStream();
            using var bytes = new MemoryStream();
            stream.CopyTo(bytes);
            var extent = group.SelectMany(e => e.Descendants<DW.Extent>()).FirstOrDefault();
            var size = extent?.Cx?.Value is { } cx && extent.Cy?.Value is { } cy ? $" width=\"{cx / 9525}\" height=\"{cy / 9525}\"" : "";
            return $"<img data-keep=\"{k}\"{size} src=\"data:{image.ContentType};base64,{Convert.ToBase64String(bytes.ToArray())}\">";
        }
        return $"<span data-keep=\"{k}\">{InlineHtml.Esc(string.Concat(group.SelectMany(e => e.Descendants<W.Text>()).Select(t => t.Text)))}</span>";
    }

    /// <summary>Replaces a header (or footer) with the given html; empty removes it. Its paragraphs are rewritten in place, so their
    /// styles and borders stay, and tables and what the html shows as placeholders are kept.</summary>
    public static void SetHeaderFooter(DocxDocument doc, bool header, bool first, string html)
    {
        var section = Section(doc);
        var main = doc.Main;
        var part = Part(doc, section, header, first);
        if (html.Trim().Length == 0)
        {
            if (Reference(section, header, first) is not { } reference) return;
            var id = reference.Id?.Value;
            reference.Remove();
            if (part is not null && !main.Document!.Descendants<W.HeaderFooterReferenceType>().Any(r => r.Id?.Value == id)) main.DeletePart(part);
            return;
        }
        if (part is null)
        {
            part = header ? main.AddNewPart<HeaderPart>() : main.AddNewPart<FooterPart>();
            var type = first ? W.HeaderFooterValues.First : W.HeaderFooterValues.Default;
            AddReference(section, header ? new W.HeaderReference { Type = type, Id = main.GetIdOfPart(part) } : new W.FooterReference { Type = type, Id = main.GetIdOfPart(part) });
        }
        OpenXmlCompositeElement container = part is HeaderPart hp ? hp.Header ??= new W.Header() : ((FooterPart)part).Footer ??= new W.Footer();
        var kept = container.Elements<W.Paragraph>().SelectMany(p => Items(doc, p)).Where(x => x.Keep is not null).Select(x => x.Keep!).ToList();
        var used = new HashSet<int>();
        var slots = container.Elements<W.Paragraph>().ToList();
        var lines = Lines(html);
        W.Paragraph? last = null;
        for (var i = 0; i < lines.Count; i++)
        {
            var p = i < slots.Count ? slots[i] : new W.Paragraph();
            if (i >= slots.Count && last is null) container.Append(p);
            else if (i >= slots.Count) last!.InsertAfterSelf(p);
            foreach (var child in p.ChildElements.Where(c => c is not W.ParagraphProperties).ToList()) child.Remove();
            if (lines[i].Align is { } align) (p.ParagraphProperties ??= new W.ParagraphProperties()).Justification = DocxParagraph.JustificationOf(align);
            foreach (var e in Content(doc, lines[i].Specs, kept, used)) p.Append(e);
            last = p;
        }
        foreach (var extra in slots.Skip(lines.Count)) extra.Remove();
        if (container.LastChild is not W.Paragraph) container.Append(new W.Paragraph());
    }

    /// <summary>The html's paragraphs: p (or div) blocks with their text-align, left when unstated; without blocks, lines that leave
    /// the paragraph's alignment alone. A br starts another paragraph; placeholders become numbered marks in the text.</summary>
    static List<(string? Align, List<RunSpec> Specs)> Lines(string html)
    {
        html = KeepPattern().Replace(html, m => Keep + (m.Groups[1].Success ? m.Groups[1] : m.Groups[2]).Value + KeepEnd);
        var blocks = BlockPattern().Matches(html);
        List<(string? Align, string Html)> parts = blocks.Count == 0 ? [(null, html)]
            : [.. blocks.Select(m => ((string?)(AlignPattern().Match(m.Groups[2].Value) is { Success: true } a ? a.Groups[1].Value.ToLowerInvariant() : "left"), m.Groups[3].Value))];
        var lines = new List<(string? Align, List<RunSpec> Specs)>();
        foreach (var (align, inner) in parts)
        {
            var line = new List<RunSpec>();
            lines.Add((align, line));
            foreach (var spec in InlineHtml.Parse(inner).Select(s => s.Flattened()))
            {
                var pieces = spec.Text.Split('\n');
                for (var i = 0; i < pieces.Length; i++)
                {
                    if (i > 0) lines.Add((align, line = []));
                    if (pieces[i].Length > 0) line.Add(spec with { Text = pieces[i] });
                }
            }
        }
        return lines;
    }

    /// <summary>A line's content: text as runs, {page} / {pages} as fields formatted like their text, placeholders as the elements
    /// they stand for (a copy when one is used twice).</summary>
    static IEnumerable<OpenXmlElement> Content(DocxDocument doc, List<RunSpec> specs, List<List<OpenXmlElement>> kept, HashSet<int> used)
    {
        foreach (var spec in specs)
            foreach (var piece in TokenPattern().Split(spec.Text))
            {
                if (piece.Length == 0) continue;
                if (piece is "{page}" or "{pages}")
                    yield return new W.SimpleField(DocxRuns.MakeRuns(doc, [spec with { Text = "1" }], null)) { Instruction = piece == "{page}" ? " PAGE " : " NUMPAGES " };
                else if (piece[0] == Keep)
                {
                    if (!int.TryParse(piece.AsSpan(1, piece.Length - 2), NumberStyles.None, CultureInfo.InvariantCulture, out var k) || k >= kept.Count) continue;
                    var fresh = used.Add(k);
                    foreach (var e in kept[k]) yield return fresh ? e : e.CloneNode(true);
                }
                else foreach (var run in DocxRuns.MakeRuns(doc, [spec with { Text = piece }], null)) yield return run;
            }
    }

    [GeneratedRegex(@"<img\b[^>]*?\bdata-keep=""(\d+)""[^>]*>|<span\b[^>]*?\bdata-keep=""(\d+)""[^>]*>.*?</span\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex KeepPattern();

    [GeneratedRegex(@"<(p|div)\b([^>]*)>(.*?)</\1\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex BlockPattern();

    [GeneratedRegex(@"text-align\s*:\s*([a-z]+)", RegexOptions.IgnoreCase)]
    private static partial Regex AlignPattern();

    [GeneratedRegex("(\\{pages?\\}|\uE000\\d+\uE001)")]
    private static partial Regex TokenPattern();
}
