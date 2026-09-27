using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Writer.Core;
using Writer.Formats.Cite;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>A citation in the text as Word makes one (References › Insert Citation): its citation content control (w:sdt with
/// w:citation) around a CITATION field, whose result is the citation as the document's style writes it. It sits at a character offset
/// of its paragraph's text, as a note's mark does; its own text is not the paragraph's, so editing the words around it leaves it whole.</summary>
sealed class DocxCitation(DocxDocument doc, W.Paragraph p, W.SdtRun sdt) : Node
{
    public override string Kind => "citation";
    public override object Anchor => sdt;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var field = CiteField.Of(sdt);
        var props = new Dictionary<string, string>
        {
            ["id"] = DocxCitations.IdOf(sdt),
            ["sources"] = string.Join(";", field.Sources.Select(x => x.Tag)),
            ["text"] = DocxCitations.ResultText(sdt),
            ["html"] = DocxCitations.ResultHtml(sdt),
            ["at"] = DocxFootnotes.OffsetOf(p, sdt).ToString(CultureInfo.InvariantCulture),
        };
        if (field.Sources.Any(x => x.Pages.Length > 0)) props["pages"] = string.Join(";", field.Sources.Select(x => x.Pages));
        if (field.NoAuthor) props["noAuthor"] = "true";
        if (field.NoYear) props["noYear"] = "true";
        if (field.Prefix.Length > 0) props["prefix"] = field.Prefix;
        if (field.Suffix.Length > 0) props["suffix"] = field.Suffix;
        return props;
    }

    public override string GetRaw() => sdt.OuterXml;

    public override void SetProp(string name, string value)
    {
        if (name == "at")
        {
            var (previous, next) = (sdt.PreviousSibling(), sdt.NextSibling());
            sdt.Remove();
            DocxRuns.Rejoin(previous, next);
            DocxFootnotes.Place(p, sdt, int.Parse(value, CultureInfo.InvariantCulture));
            return;
        }
        CiteField.Write(sdt, CiteField.Of(sdt).With(name, value, doc));
        DocxCitations.Refresh(doc);
    }

    public override void Remove()
    {
        var (previous, next) = (sdt.PreviousSibling(), sdt.NextSibling());
        sdt.Remove();
        DocxRuns.Rejoin(previous, next);
        DocxCitations.Refresh(doc);
    }
}

/// <summary>A works-cited list as Word makes one (References › Bibliography): its Bibliographies content control, a title, and a
/// bibliography control around a BIBLIOGRAPHY field whose result is an entry per source of the document, in its style and order.</summary>
sealed class DocxBibliography(DocxDocument doc, W.SdtBlock sdt) : Node
{
    const string Gallery = "Bibliographies";

    public override string Kind => "bibliography";
    public override object Anchor => sdt;

    public static bool Is(W.SdtBlock block) =>
        block.SdtProperties?.Descendants<W.DocPartGallery>().Any(g => g.Val?.Value == Gallery) == true
        || block.SdtProperties?.GetFirstChild<W.SdtContentBibliography>() is not null;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string>();
        if (TitleOf(sdt) is { Length: > 0 } title) props["title"] = title;
        props["text"] = string.Join("\n", Entries(sdt).Select(DocxRuns.ParagraphText).Where(t => t.Length > 0));
        props["html"] = string.Join("", Entries(sdt).Select(e => "<p>" + DocxCitations.ParagraphHtml(e) + "</p>"));
        return props;
    }

    public override string GetRaw() => sdt.OuterXml;

    public override void SetProp(string name, string value)
    {
        if (name == "title") Fill(doc, sdt, value);
        else Fill(doc, sdt, TitleOf(sdt) ?? ""); // anything else: drawn again from the sources
    }

    public override void Remove() => DocxBlocks.Detach(sdt);
    public override void MoveTo(Node newParent, int? index) => DocxBlocks.Move(newParent, sdt, index);
    public override void SetRaw(string raw) => DocxBlocks.ReplaceRaw(doc, sdt, raw);

    public static (OpenXmlElement Element, string[] Consumed) New(DocxDocument doc, IReadOnlyDictionary<string, string> props)
    {
        var sdt = new W.SdtBlock(
            new W.SdtProperties(new W.SdtId { Val = DocxCitations.NewId() }, new W.SdtContentDocPartObject(new W.DocPartGallery { Val = Gallery }, new W.DocPartUnique())),
            new W.SdtContentBlock());
        Fill(doc, sdt, props.GetValueOrDefault("title"));
        return (sdt, ["title"]);
    }

    /// <summary>The title paragraphs above the list: whatever the Bibliographies control holds before the bibliography control.</summary>
    public static string? TitleOf(W.SdtBlock sdt)
    {
        var content = sdt.SdtContentBlock?.ChildElements.ToList() ?? [];
        var at = content.FindIndex(e => e is W.SdtBlock b && b.SdtProperties?.GetFirstChild<W.SdtContentBibliography>() is not null);
        var title = string.Join("\n", content.Take(at < 0 ? 0 : at).OfType<W.Paragraph>().Select(DocxRuns.ParagraphText).Where(t => t.Length > 0));
        return title.Length > 0 ? title : null;
    }

    static IEnumerable<W.Paragraph> Entries(W.SdtBlock sdt)
    {
        var inner = sdt.Descendants<W.SdtBlock>().FirstOrDefault(b => b.SdtProperties?.GetFirstChild<W.SdtContentBibliography>() is not null);
        return inner?.SdtContentBlock?.Elements<W.Paragraph>() ?? (sdt.SdtProperties?.GetFirstChild<W.SdtContentBibliography>() is not null ? sdt.SdtContentBlock?.Elements<W.Paragraph>() ?? [] : []);
    }

    /// <summary>The list drawn again: its title (the style's own when null: Works Cited, References, Bibliography) on a new page,
    /// centred, as the styles set it, and an entry per source in the style's order, each with a hanging indent of half an inch.</summary>
    public static void Fill(DocxDocument doc, W.SdtBlock sdt, string? title)
    {
        var style = DocxSources.Style(doc) ?? CiteStyle.Mla;
        var sources = DocxSources.Read(doc);
        var format = new Citations(style, sources);
        title ??= Citations.ListTitle(style);
        var content = sdt.SdtContentBlock ??= new W.SdtContentBlock();
        content.RemoveAllChildren();
        var bare = sdt.SdtProperties?.GetFirstChild<W.SdtContentBibliography>() is not null; // a bibliography control on its own: its entries only
        var dbl = style is CiteStyle.Mla or CiteStyle.Apa;
        W.SpacingBetweenLines Spacing(bool entry) => dbl ? new W.SpacingBetweenLines { Before = "0", After = "0", Line = "480", LineRule = W.LineSpacingRuleValues.Auto }
            : new W.SpacingBetweenLines { Before = "0", After = entry ? "240" : "480", Line = "240", LineRule = W.LineSpacingRuleValues.Auto };
        if (title.Length > 0 && !bare)
        {
            var run = new W.Run(DocxRuns.TextElements(title));
            if (style == CiteStyle.Apa) run.PrependChild(new W.RunProperties(new W.Bold()));
            content.Append(new W.Paragraph(new W.ParagraphProperties(new W.KeepNext(), new W.PageBreakBefore(), Spacing(false), new W.Justification { Val = W.JustificationValues.Center },
                new W.OutlineLevel { Val = 0 }), run)); // an outline level: in the navigation pane and the contents, drawn as the text around it
        }
        var entries = format.Sort(sources).Select(s => format.Entry(s)).ToList();
        if (entries.Count == 0) entries.Add([new Span("There are no sources in the current document.")]);
        var styleId = doc.Styles.TryResolveStyle("Bibliography", "paragraph");
        var paragraphs = entries.Select(spans =>
        {
            var pp = new W.ParagraphProperties();
            if (styleId is not null) pp.ParagraphStyleId = new W.ParagraphStyleId { Val = styleId };
            pp.SpacingBetweenLines = Spacing(true);
            pp.Indentation = new W.Indentation { Left = "720", Hanging = "720" };
            var para = new W.Paragraph(pp);
            foreach (var r in DocxCitations.Runs(spans, null)) para.Append(r);
            return para;
        }).ToList();
        paragraphs[0].InsertAfter(DocxCitations.FieldRun(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }, null), paragraphs[0].ParagraphProperties);
        paragraphs[0].InsertAfter(DocxCitations.FieldRun(new W.FieldCode(" BIBLIOGRAPHY ") { Space = SpaceProcessingModeValues.Preserve }, null), paragraphs[0].ParagraphProperties);
        paragraphs[0].InsertAfter(DocxCitations.FieldRun(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }, null), paragraphs[0].ParagraphProperties);
        paragraphs[^1].Append(DocxCitations.FieldRun(new W.FieldChar { FieldCharType = W.FieldCharValues.End }, null));
        if (bare) content.Append(paragraphs);
        else content.Append(new W.SdtBlock(new W.SdtProperties(new W.SdtId { Val = DocxCitations.NewId() }, new W.SdtContentBibliography()), new W.SdtContentBlock(paragraphs)));
    }
}

/// <summary>A CITATION field's code: the sources it cites (the first by its tag, more after \m), each with its pages (\p), and \n
/// (no author), \y (no year), \f (text before) and \s (text after).</summary>
sealed record CiteField(List<(string Tag, string Pages)> Sources, bool NoAuthor = false, bool NoYear = false, string Prefix = "", string Suffix = "")
{
    static readonly Regex Token = new(@"""[^""]*""|\S+");

    public static CiteField Of(W.SdtRun sdt) => Parse(DocxCitations.Code(sdt));

    public static CiteField Parse(string code)
    {
        var tokens = Token.Matches(code).Select(m => m.Value.Trim('"')).ToList();
        var f = new CiteField([]);
        var (noAuthor, noYear, prefix, suffix) = (false, false, "", "");
        for (var i = 0; i < tokens.Count; i++)
        {
            var t = tokens[i];
            string Next() => i + 1 < tokens.Count ? tokens[++i] : "";
            if (i == 0 && t.Equals("CITATION", StringComparison.OrdinalIgnoreCase)) continue;
            switch (t)
            {
                case "\\m": f.Sources.Add((Next(), "")); break;
                case "\\p" when f.Sources.Count > 0: f.Sources[^1] = (f.Sources[^1].Tag, Next()); break;
                case "\\n": noAuthor = true; break;
                case "\\y": noYear = true; break;
                case "\\f": prefix = Next(); break;
                case "\\s": suffix = Next(); break;
                case "\\l" or "\\v" or "\\p": Next(); break;
                case "\\t": break;
                default: if (!t.StartsWith('\\') && f.Sources.Count == 0) f.Sources.Add((t, "")); break;
            }
        }
        return f with { NoAuthor = noAuthor, NoYear = noYear, Prefix = prefix, Suffix = suffix };
    }

    public string Code()
    {
        static string Q(string v) => v.Any(char.IsWhiteSpace) || v.Length == 0 ? "\"" + v + "\"" : v;
        var sb = new StringBuilder("CITATION");
        for (var i = 0; i < Sources.Count; i++)
        {
            sb.Append(i == 0 ? " " : " \\m ").Append(Sources[i].Tag);
            if (Sources[i].Pages.Length > 0) sb.Append(" \\p ").Append(Q(Sources[i].Pages));
        }
        if (NoAuthor) sb.Append(" \\n");
        if (NoYear) sb.Append(" \\y");
        if (Prefix.Length > 0) sb.Append(" \\f ").Append(Q(Prefix));
        if (Suffix.Length > 0) sb.Append(" \\s ").Append(Q(Suffix));
        return sb.Append(" \\l 1033").ToString();
    }

    /// <summary>The field with a property of the citation node written: sources (tags, ; between), pages (for each source, ; between),
    /// noAuthor, noYear, prefix, suffix.</summary>
    public CiteField With(string name, string value, DocxDocument doc)
    {
        List<string> Split(string v) => v.Split(';').Select(x => x.Trim()).ToList();
        switch (name)
        {
            case "sources":
                var tags = Split(value).Where(t => t.Length > 0).ToList();
                if (tags.Count == 0) throw new WriterException(ErrorCode.Validation, "A citation cites a source", "Give sources=<tag>, the tag of one of the document's sources.");
                var known = DocxSources.Read(doc).Select(s => s.Tag).ToList();
                var unknown = tags.FirstOrDefault(t => !known.Contains(t, StringComparer.OrdinalIgnoreCase));
                if (unknown is not null) throw new WriterException(ErrorCode.PathNotFound, $"No source '{unknown}' in the document", known.Count > 0 ? "Its sources: " + string.Join(", ", known) + "." : "Add it first: set / --prop source={…}.");
                var pagesOf = Sources.ToDictionary(x => x.Tag, x => x.Pages, StringComparer.OrdinalIgnoreCase);
                return this with { Sources = tags.Select(t => (known.First(k => string.Equals(k, t, StringComparison.OrdinalIgnoreCase)), pagesOf.GetValueOrDefault(t) ?? "")).ToList() };
            case "pages":
                var pages = Split(value);
                return this with { Sources = Sources.Select((x, i) => (x.Tag, i < pages.Count && pages[i] != "none" ? pages[i] : "")).ToList() };
            case "noAuthor": return this with { NoAuthor = value == "true" };
            case "noYear": return this with { NoYear = value == "true" };
            case "prefix": return this with { Prefix = value == "none" ? "" : value };
            case "suffix": return this with { Suffix = value == "none" ? "" : value };
            default: throw new WriterException(ErrorCode.Validation, $"A citation has no '{name}'", "Its props: sources, pages, noAuthor, noYear, prefix, suffix, at.");
        }
    }

    public static void Write(W.SdtRun sdt, CiteField f)
    {
        if (DocxCitations.IsNoteCitation(sdt))
        {
            (sdt.SdtProperties!.GetFirstChild<W.Tag>() ?? sdt.SdtProperties.AppendChild(new W.Tag())).Val = DocxCitations.NoteTag + f.Code();
            return;
        }
        var code = sdt.Descendants<W.FieldCode>().FirstOrDefault();
        if (code is not null) { code.Text = " " + f.Code() + " "; code.Space = SpaceProcessingModeValues.Preserve; }
        else DocxCitations.Rebuild(sdt, f, []);
    }
}

/// <summary>Citations, notes that cite and works-cited lists: made, read, and drawn again in the document's style whenever what they
/// show may have changed (a source, the style, a citation added or taken away).</summary>
static class DocxCitations
{
    /// <summary>The tag of a note's citation in Chicago's notes: a content control Word keeps as it is (its CITATION field would draw
    /// Word's own author-date citation there), naming the sources as a CITATION field would.</summary>
    public const string NoteTag = "WriterCite ";

    public static bool IsCitation(W.SdtRun sdt) => sdt.SdtProperties?.GetFirstChild<W.SdtContentCitation>() is not null || IsNoteCitation(sdt);

    public static bool IsNoteCitation(W.SdtRun sdt) => sdt.SdtProperties?.GetFirstChild<W.Tag>()?.Val?.Value?.StartsWith(NoteTag, StringComparison.Ordinal) == true;

    /// <summary>The citations in the paragraph, in reading order.</summary>
    public static IEnumerable<Node> In(DocxDocument doc, W.Paragraph p) =>
        p.Descendants<W.SdtRun>().Where(s => s.SdtProperties?.GetFirstChild<W.SdtContentCitation>() is not null).Select(s => (Node)new DocxCitation(doc, p, s));

    public static string IdOf(W.SdtRun sdt) => sdt.SdtProperties?.GetFirstChild<W.SdtId>()?.Val?.Value.ToString(CultureInfo.InvariantCulture) ?? "";

    /// <summary>The code of a citation's field (a note's citation keeps it in its tag).</summary>
    public static string Code(W.SdtRun sdt) => IsNoteCitation(sdt)
        ? sdt.SdtProperties!.GetFirstChild<W.Tag>()!.Val!.Value![NoteTag.Length..]
        : string.Concat(sdt.Descendants<W.FieldCode>().Select(c => c.Text)).Trim() is { Length: > 0 } code ? code
            : sdt.Descendants<W.SimpleField>().FirstOrDefault()?.Instruction?.Value?.Trim() ?? "";

    static readonly Random Ids = new();
    public static int NewId() { lock (Ids) return Ids.Next(1, int.MaxValue); }

    /// <summary>Adds a citation of `sources` (tags, ; between) at a character offset of the paragraph (its end when none), in the look
    /// of the text there.</summary>
    public static Node Add(DocxDocument doc, W.Paragraph p, IReadOnlyDictionary<string, string> props)
    {
        var field = new CiteField([]);
        field = field.With("sources", props.GetValueOrDefault("sources") ?? props.GetValueOrDefault("source") ?? "", doc);
        foreach (var name in new[] { "pages", "noAuthor", "noYear", "prefix", "suffix" })
            if (props.TryGetValue(name, out var v)) field = field.With(name, v, doc);
        var at = props.TryGetValue("at", out var a) ? int.Parse(a, CultureInfo.InvariantCulture) : int.MaxValue;
        var sdt = new W.SdtRun(new W.SdtProperties(new W.SdtId { Val = NewId() }, new W.SdtContentCitation()), new W.SdtContentRun());
        Rebuild(sdt, field, [], LookAt(p, at));
        DocxFootnotes.Place(p, sdt, at);
        Refresh(doc);
        return new DocxCitation(doc, p, sdt);
    }

    /// <summary>The run properties of the text a citation goes into (its font and size), without what would dress the citation itself.</summary>
    static W.RunProperties? LookAt(W.Paragraph p, int at)
    {
        var pos = 0;
        W.Run? last = null;
        foreach (var (r, _) in DocxRuns.Walk(p))
        {
            last = r;
            pos += DocxRuns.RunText(r).Length;
            if (pos >= at) break;
        }
        if (last?.RunProperties?.CloneNode(true) is not W.RunProperties rp) return null;
        foreach (var e in rp.ChildElements.Where(e => e is W.Italic or W.ItalicComplexScript or W.Bold or W.BoldComplexScript or W.Underline or W.Strike or W.VerticalTextAlignment or W.RunStyle).ToList()) e.Remove();
        return rp.HasChildren ? rp : null;
    }

    // ----- a note that cites (Chicago) -----

    /// <summary>The citation of a note, when it has one.</summary>
    public static W.SdtRun? NoteCitation(OpenXmlElement note) => note.Descendants<W.SdtRun>().FirstOrDefault(IsNoteCitation);

    /// <summary>Makes the note cite `sources` (with its pages), in front of the note's own text; none takes the citation out.</summary>
    public static void SetNoteCitation(DocxDocument doc, OpenXmlElement note, string? sources, string? pages)
    {
        var existing = NoteCitation(note);
        if (sources == "none") { existing?.Remove(); Refresh(doc); return; }
        var field = existing is not null ? CiteField.Of(existing) : new CiteField([]);
        if (sources is not null) field = field.With("sources", sources, doc);
        if (pages is not null) field = field.With("pages", pages, doc);
        if (field.Sources.Count == 0) throw new WriterException(ErrorCode.Validation, "A note's citation cites a source", "Give cite=<tag>.");
        if (existing is null)
        {
            existing = new W.SdtRun(new W.SdtProperties(new W.SdtId { Val = NewId() }, new W.Tag { Val = NoteTag + field.Code() }), new W.SdtContentRun());
            PlaceNoteCitation(note.Elements<W.Paragraph>().FirstOrDefault() ?? note.AppendChild(new W.Paragraph()), existing);
        }
        else CiteField.Write(existing, field);
        Refresh(doc);
    }

    /// <summary>Puts a note's citation in its first paragraph: after the note's own mark and a space, before the note's own text.</summary>
    public static void PlaceNoteCitation(W.Paragraph p, W.SdtRun cite)
    {
        static W.Run Space() => new(new W.Text(" ") { Space = SpaceProcessingModeValues.Preserve });
        var mark = p.Elements<W.Run>().FirstOrDefault(r => r.Descendants<W.FootnoteReferenceMark>().Any() || r.Descendants<W.EndnoteReferenceMark>().Any());
        if (mark is not null) { mark.InsertAfterSelf(cite); cite.InsertBeforeSelf(Space()); }
        else if (p.ParagraphProperties is { } pp) pp.InsertAfterSelf(cite);
        else p.PrependChild(cite);
        if (cite.NextSibling() is W.Run after && DocxRuns.RunText(after) is { Length: > 0 } t && !char.IsWhiteSpace(t[0])) cite.InsertAfterSelf(Space());
    }

    // ----- drawing them again -----

    /// <summary>Draws every citation, note citation and works-cited list again from the sources in the document's style: citations in
    /// the order they are read (a Chicago note cites in full the first time, short after), each list sorted as the style sorts.</summary>
    public static void Refresh(DocxDocument doc)
    {
        var body = doc.Main.Document?.Body;
        if (body is null) return;
        var sources = DocxSources.Read(doc);
        var style = DocxSources.Style(doc) ?? CiteStyle.Mla;
        var format = new Citations(style, sources);
        var inText = style == CiteStyle.Chicago ? new Citations(CiteStyle.ChicagoDate, sources) : format; // Chicago's notes style: a citation in the text is author-date
        var byTag = sources.GroupBy(s => s.Tag, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        var cited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        List<Cited> Cites(CiteField f) => f.Sources.Select(x => byTag.TryGetValue(x.Tag, out var s) ? new Cited(s, x.Pages) : new Cited(new Source { Tag = x.Tag, Title = x.Tag }, x.Pages)).ToList();
        void Draw(OpenXmlElement container)
        {
            foreach (var e in container.Descendants().ToList())
            {
                if (e is W.SdtRun sdt && IsCitation(sdt))
                {
                    var f = CiteField.Of(sdt);
                    var cites = Cites(f);
                    if (IsNoteCitation(sdt))
                    {
                        SetResult(sdt, f, Affix(f, new Citations(CiteStyle.Chicago, sources).Note(cites, cited)));
                        foreach (var c in cites) cited.Add(c.Source.Tag); // a source's first note is its full one; a citation in the text is no note
                    }
                    else SetResult(sdt, f, Affix(f, inText.InText(cites, f.NoAuthor, f.NoYear)));
                }
                else if (e is W.FootnoteReference fr && fr.Id?.Value is { } fid && DocxFootnotes.Note(doc, false, fid.ToString(CultureInfo.InvariantCulture)) is { } footnote) Draw(footnote);
                else if (e is W.EndnoteReference er && er.Id?.Value is { } eid && DocxFootnotes.Note(doc, true, eid.ToString(CultureInfo.InvariantCulture)) is { } endnote) Draw(endnote);
            }
        }
        Draw(body);
        foreach (var list in body.Descendants<W.SdtBlock>().Where(DocxBibliography.Is).Where(b => !b.Ancestors<W.SdtBlock>().Any(DocxBibliography.Is)).ToList())
        {
            var title = DocxBibliography.TitleOf(list) ?? "";
            // a list still under another style's own title takes this style's
            if (Enum.GetValues<CiteStyle>().Select(Citations.ListTitle).Contains(title)) title = Citations.ListTitle(style);
            DocxBibliography.Fill(doc, list, title);
        }
    }

    /// <summary>A citation with the text a field puts before and after it (\f, \s), inside its parentheses.</summary>
    static List<Span> Affix(CiteField f, List<Span> spans)
    {
        if (f.Prefix.Length == 0 && f.Suffix.Length == 0) return spans;
        var list = spans.ToList();
        var open = list.Count > 0 && list[0].Text.StartsWith('(');
        if (f.Prefix.Length > 0)
        {
            if (open) list[0] = list[0] with { Text = "(" + f.Prefix + " " + list[0].Text[1..] };
            else list.Insert(0, new Span(f.Prefix + " "));
        }
        if (f.Suffix.Length > 0)
        {
            var last = list[^1];
            if (last.Text.EndsWith(')')) list[^1] = last with { Text = last.Text[..^1] + ", " + f.Suffix + ")" };
            else list.Add(new Span(" " + f.Suffix));
        }
        return list;
    }

    /// <summary>Writes a citation's result: the runs between its field's separator and end (a note's citation: all its runs), in the look
    /// of the result before it.</summary>
    static void SetResult(W.SdtRun sdt, CiteField f, List<Span> spans)
    {
        var content = sdt.SdtContentRun ??= new W.SdtContentRun();
        var runs = content.Elements<W.Run>().ToList();
        var look = runs.FirstOrDefault(r => DocxRuns.HasText(r))?.RunProperties ?? runs.FirstOrDefault()?.RunProperties;
        var clean = look?.CloneNode(true) as W.RunProperties;
        if (clean is not null) foreach (var e in clean.ChildElements.Where(e => e is W.Italic or W.ItalicComplexScript).ToList()) e.Remove();
        if (IsNoteCitation(sdt))
        {
            content.RemoveAllChildren();
            foreach (var r in Runs(spans, clean)) content.Append(r);
            return;
        }
        var separate = runs.FindIndex(r => r.GetFirstChild<W.FieldChar>()?.FieldCharType?.Value == W.FieldCharValues.Separate);
        var end = runs.FindIndex(r => r.GetFirstChild<W.FieldChar>()?.FieldCharType?.Value == W.FieldCharValues.End);
        if (separate < 0 || end < separate) { Rebuild(sdt, f, spans, clean); return; }
        foreach (var r in runs.Skip(separate + 1).Take(end - separate - 1)) r.Remove();
        var at = runs[separate];
        foreach (var r in Runs(spans, clean)) { at.InsertAfterSelf(r); at = r; }
    }

    /// <summary>A citation's content anew: its field's begin, code, separator, result and end, in `look`.</summary>
    public static void Rebuild(W.SdtRun sdt, CiteField f, List<Span> spans, W.RunProperties? look = null)
    {
        var content = sdt.SdtContentRun ??= new W.SdtContentRun();
        content.RemoveAllChildren();
        content.Append(FieldRun(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }, look));
        content.Append(FieldRun(new W.FieldCode(" " + f.Code() + " ") { Space = SpaceProcessingModeValues.Preserve }, look));
        content.Append(FieldRun(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }, look));
        foreach (var r in Runs(spans, look)) content.Append(r);
        content.Append(FieldRun(new W.FieldChar { FieldCharType = W.FieldCharValues.End }, look));
    }

    public static W.Run FieldRun(OpenXmlElement e, W.RunProperties? look)
    {
        var run = new W.Run();
        if (look?.CloneNode(true) is W.RunProperties rp) run.Append(rp);
        run.Append(e);
        return run;
    }

    /// <summary>Formatted text as runs in `look`, italic where it is.</summary>
    public static IEnumerable<W.Run> Runs(IEnumerable<Span> spans, W.RunProperties? look)
    {
        foreach (var s in spans.Where(s => s.Text.Length > 0))
        {
            var rp = look?.CloneNode(true) as W.RunProperties ?? new W.RunProperties();
            rp.NoProof = new W.NoProof();
            if (s.Italic) rp.Italic = new W.Italic();
            var run = new W.Run(rp);
            foreach (var t in DocxRuns.TextElements(s.Text)) run.Append(t);
            foreach (var t in run.Elements<W.Text>().Where(t => t.Text.Length > 0 && (char.IsWhiteSpace(t.Text[0]) || char.IsWhiteSpace(t.Text[^1])))) t.Space = SpaceProcessingModeValues.Preserve;
            yield return run;
        }
    }

    /// <summary>A citation's text as it shows: its field's result (a note's citation: all of it).</summary>
    public static string ResultText(W.SdtRun sdt) => string.Concat(ResultRuns(sdt).Select(DocxRuns.RunText));

    public static string ResultHtml(W.SdtRun sdt) => Common.InlineHtml.Render(ResultRuns(sdt).Select(r => new RunSpec(DocxRuns.RunText(r), Italic: DocxRun.On(r.RunProperties?.Italic))));

    public static string ParagraphHtml(W.Paragraph p) => Common.InlineHtml.Render(DocxRuns.Walk(p).Select(x => new RunSpec(DocxRuns.RunText(x.Run), Italic: DocxRun.On(x.Run.RunProperties?.Italic), Bold: DocxRun.On(x.Run.RunProperties?.Bold))));

    static IEnumerable<W.Run> ResultRuns(W.SdtRun sdt)
    {
        var runs = sdt.SdtContentRun?.Descendants<W.Run>().ToList() ?? [];
        if (IsNoteCitation(sdt)) return runs;
        var separate = runs.FindIndex(r => r.GetFirstChild<W.FieldChar>()?.FieldCharType?.Value == W.FieldCharValues.Separate);
        var end = runs.FindIndex(r => r.GetFirstChild<W.FieldChar>()?.FieldCharType?.Value == W.FieldCharValues.End);
        return separate < 0 ? runs.Where(DocxRuns.HasText) : runs.Skip(separate + 1).Take((end < 0 ? runs.Count : end) - separate - 1);
    }

    /// <summary>Whether a source is cited anywhere in the document: in a citation or a note's.</summary>
    public static bool Cites(DocxDocument doc, string tag)
    {
        IEnumerable<OpenXmlElement> Parts()
        {
            if (doc.Main.Document?.Body is { } body) yield return body;
            if (doc.Main.FootnotesPart?.Footnotes is { } fn) yield return fn;
            if (doc.Main.EndnotesPart?.Endnotes is { } en) yield return en;
        }
        return Parts().SelectMany(p => p.Descendants<W.SdtRun>()).Where(IsCitation).Any(s => CiteField.Of(s).Sources.Any(x => string.Equals(x.Tag, tag, StringComparison.OrdinalIgnoreCase)));
    }
}
