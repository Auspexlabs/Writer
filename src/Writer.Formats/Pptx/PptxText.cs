using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
using Writer.Formats.Docx;
using A = DocumentFormat.OpenXml.Drawing;
using Source = Writer.Formats.Docx.DocxReplace.Source;

namespace Writer.Formats.Pptx;

/// <summary>DrawingML text: reading and writing paragraphs and runs inside shapes and table cells.</summary>
static class PptxText
{
    public static IEnumerable<A.Paragraph> Paragraphs(OpenXmlCompositeElement? body) => body?.Elements<A.Paragraph>() ?? [];

    public static string ParagraphText(A.Paragraph p)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var child in p.ChildElements)
        {
            switch (child)
            {
                case A.Run r: sb.Append(r.Text?.Text); break;
                case A.Field f: sb.Append(f.Text?.Text); break;
                case A.Break: sb.Append('\n'); break;
            }
        }
        return sb.ToString();
    }

    public static string BodyText(OpenXmlCompositeElement? body) => string.Join("\n", Paragraphs(body).Select(ParagraphText));

    /// <summary>Writes new text into a text body in place, as DocxReplace does for a Word paragraph. Line breaks in the runs start
    /// new paragraphs. Old and new paragraphs are matched (longest common subsequence) with the formatting the source can express:
    /// a paragraph that is the same is left alone; a changed one keeps its pPr and endParaRPr and has only the words that differ
    /// rewritten (see Rewrite); a new one takes its neighbour's pPr and endParaRPr, and the run properties of the run it continues.
    /// bodyPr and lstStyle are never touched.</summary>
    public static void SetBody(OpenXmlCompositeElement body, SlidePart slide, IEnumerable<RunSpec> specs, Source source = Source.Text)
    {
        var key = KeyFor(source);
        var lines = SplitLines(specs);
        var olds = body.Elements<A.Paragraph>().ToList();
        var ids = new Dictionary<string, int>();
        int Id(IEnumerable<RunSpec> runs)
        {
            var k = ParagraphKey(runs, key);
            return ids.TryGetValue(k, out var id) ? id : ids[k] = ids.Count;
        }
        var x = olds.Select(p => Id(Pieces(p, slide, source == Source.Text).Select(piece => piece.Spec))).ToArray();
        var y = lines.Select(Id).ToArray();
        foreach (var (a, b, c, d) in DocxReplace.Spans(x, y))
        {
            var paired = Math.Min(b - a, d - c);
            for (var i = 0; i < paired; i++) Rewrite(olds[a + i], slide, lines[c + i], source);
            for (var i = a + paired; i < b; i++) olds[i].Remove();
            var previous = paired > 0 ? olds[a + paired - 1] : a > 0 ? olds[a - 1] : null;
            var next = b < olds.Count ? olds[b] : null;
            for (var i = c + paired; i < d; i++)
            {
                var template = previous ?? next;
                var p = new A.Paragraph();
                if (template?.ParagraphProperties is { } pPr) p.Append(pPr.CloneNode(true));
                if (template?.GetFirstChild<A.EndParagraphRunProperties>() is { } end) p.Append(end.CloneNode(true));
                var runs = template?.Elements<A.Run>();
                var continued = previous is not null ? runs?.LastOrDefault() : runs?.FirstOrDefault();
                if (previous is not null) previous.InsertAfterSelf(p);
                else if (next is not null) next.InsertBeforeSelf(p);
                else body.Append(p);
                Rewrite(p, slide, lines[i], source, continued?.RunProperties);
                previous = p;
            }
        }
    }

    /// <summary>Writes new runs into one paragraph in place (see Rewrite); line breaks become a:br.</summary>
    public static void SetParagraph(A.Paragraph p, SlidePart slide, IEnumerable<RunSpec> specs, Source source = Source.Text) =>
        Rewrite(p, slide, specs.Select(s => s with { Text = s.Text.ReplaceLineEndings("\n") }).ToList(), source);

    /// <summary>A stretch of a paragraph's text: a run, or with plain text also a field's text or a line break (\n).</summary>
    sealed record Piece(OpenXmlElement Element, int Start, int End, RunSpec Spec);

    /// <summary>The paragraph's text as the source sees it. Plain text shows fields and line breaks; html and markdown do not (a
    /// shape's html has neither), so there they keep their place.</summary>
    static List<Piece> Pieces(A.Paragraph p, SlidePart slide, bool textual)
    {
        var pieces = new List<Piece>();
        var pos = 0;
        foreach (var child in p.ChildElements)
        {
            var text = child switch
            {
                A.Run r => r.Text?.Text ?? "",
                A.Field f when textual => f.Text?.Text ?? "",
                A.Break when textual => "\n",
                _ => null,
            };
            if (text is null) continue;
            pieces.Add(new Piece(child, pos, pos + text.Length, SpecOf(PropertiesOf(child), slide, text)));
            pos += text.Length;
        }
        return pieces;
    }

    static A.RunProperties? PropertiesOf(OpenXmlElement? e) => e switch
    {
        A.Run r => r.RunProperties,
        A.Field f => f.RunProperties,
        A.Break b => b.RunProperties,
        _ => null,
    };

    static RunSpec SpecOf(A.RunProperties? rPr, SlidePart slide, string text)
    {
        var props = new Dictionary<string, string> { ["text"] = text };
        ReadRun(rPr, slide, props);
        return RunSpec.FromProps(props);
    }

    /// <summary>What the source can say about a run's formatting (DocxReplace's keys): plain text nothing; markdown bold, italic,
    /// strike, code and link; html everything a slide run keeps, a link as html shows it.</summary>
    static Func<RunSpec, RunSpec?> KeyFor(Source source) => source switch
    {
        Source.Text => _ => null,
        Source.Markdown => s =>
        {
            s = s.Flattened();
            return new RunSpec("", s.Bold, s.Italic, s.Strike, s.Code, s.Link);
        },
        _ => s => s.Flattened() with { Text = "", Style = null, Outline = false, Shadow = false, Link = s.Link is null ? null : InlineHtml.SafeUrl(s.Link) },
    };

    /// <summary>A paragraph's text with its keys, runs that say the same run together.</summary>
    static string ParagraphKey(IEnumerable<RunSpec> runs, Func<RunSpec, RunSpec?> key)
    {
        var sb = new StringBuilder();
        RunSpec? last = null;
        var first = true;
        foreach (var run in runs)
        {
            if (run.Text.Length == 0) continue;
            var k = key(run);
            if (first || !Equals(k, last)) sb.Append('\u0001').Append(k).Append('\u0002');
            sb.Append(run.Text);
            (last, first) = (k, false);
        }
        return sb.ToString();
    }

    /// <summary>Rewrites only the words that differ (matched word by word, a word also ending where its formatting does); the
    /// text in between keeps its runs untouched. New text takes the run properties of the nearest original run with the same
    /// formatting, else of the text it replaces or continues, and changes only what the source states differently.</summary>
    static void Rewrite(A.Paragraph p, SlidePart slide, List<RunSpec> input, Source source, A.RunProperties? fallback = null)
    {
        var key = KeyFor(source);
        var specs = input.Where(s => s.Text.Length > 0).ToList();
        var pieces = Pieces(p, slide, source == Source.Text);
        var hunks = DocxReplace.Hunks(string.Concat(pieces.Select(x => x.Spec.Text)), DocxReplace.Keys(pieces.Select(x => x.Spec), key),
            string.Concat(specs.Select(s => s.Text)), DocxReplace.Keys(specs, key));
        // from the end backwards, so the offsets of the stretches still to do stay valid
        for (var i = hunks.Count - 1; i >= 0; i--)
        {
            var (a, b, c, d) = hunks[i];
            Apply(p, slide, a, b, DocxReplace.Take(specs, c, d), source, fallback);
        }
    }

    /// <summary>Rewrites old characters [a, b) as the given runs.</summary>
    static void Apply(A.Paragraph p, SlidePart slide, int a, int b, List<RunSpec> middle, Source source, A.RunProperties? fallback)
    {
        var key = KeyFor(source);
        var textual = source == Source.Text;
        Cut(Pieces(p, slide, textual), b);
        Cut(Pieces(p, slide, textual), a); // afresh: an insertion (a = b) must not cut the same run twice
        var pieces = Pieces(p, slide, textual);
        var region = pieces.Where(x => x.Start >= a && x.End <= b && x.End > x.Start).ToList();
        var before = pieces.LastOrDefault(x => x.End == a && x.End > x.Start);
        var after = pieces.FirstOrDefault(x => x.Start == b && x.End > x.Start);
        var near = region.Concat(new[] { before, after }.OfType<Piece>()).ToList();
        var made = new List<OpenXmlElement>();
        foreach (var (spec, from) in Sources(region, near, middle, key))
        {
            var rPr = from is not null ? (A.RunProperties?)PropertiesOf(from.Element)?.CloneNode(true) ?? new A.RunProperties()
                : (A.RunProperties?)fallback?.CloneNode(true) ?? AsRunProperties(p.GetFirstChild<A.EndParagraphRunProperties>()) ?? new A.RunProperties { Language = "en-US" };
            Restate(rPr, slide, key(from?.Spec ?? SpecOf(rPr, slide, "")), key(spec));
            var lines = spec.Text.Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                if (i > 0) made.Add(new A.Break { RunProperties = (A.RunProperties)rPr.CloneNode(true) });
                if (lines[i].Length > 0) made.Add(rPr.HasAttributes || rPr.HasChildren ? new A.Run((A.RunProperties)rPr.CloneNode(true), new A.Text(lines[i])) : new A.Run(new A.Text(lines[i])));
            }
        }
        // where the replaced text began (a break or field inside it, which html does not show, stays after the new text), else
        // after the text before, else before the text after
        var cursor = before?.Element;
        foreach (var e in made)
        {
            if (region.Count > 0) region[0].Element.InsertBeforeSelf(e);
            else if (cursor is not null) cursor = cursor.InsertAfterSelf(e);
            else if (after is not null) after.Element.InsertBeforeSelf(e);
            else if (p.GetFirstChild<A.EndParagraphRunProperties>() is { } end) end.InsertBeforeSelf(e);
            else p.Append(e);
        }
        foreach (var x in region) x.Element.Remove();
        Merge([before?.Element, .. made, after?.Element]);
        // a paragraph left without text keeps the look of its last run as its end mark, for what is typed there next
        if (made.Count == 0 && !p.Elements<A.Run>().Any() && p.GetFirstChild<A.EndParagraphRunProperties>() is null && PropertiesOf(region.LastOrDefault()?.Element) is { } last)
            p.Append(Copy<A.EndParagraphRunProperties>(last));
    }

    /// <summary>The run each stretch of new text takes its properties from. Over several old runs it goes character by character
    /// (a kept character from its own run, a new one from the run of the character it replaces or follows), so a word across runs
    /// that differ only in what the source cannot state (language, kerning) keeps each run's; otherwise the nearest run that looks
    /// the same, else the one replaced or continued.</summary>
    static IEnumerable<(RunSpec Spec, Piece? From)> Sources(List<Piece> region, List<Piece> near, List<RunSpec> middle, Func<RunSpec, RunSpec?> key)
    {
        if (region.Count < 2)
        {
            foreach (var spec in middle) yield return (spec, near.FirstOrDefault(x => Equals(key(x.Spec), key(spec))) ?? near.FirstOrDefault());
            yield break;
        }
        var owner = region.SelectMany((x, i) => Enumerable.Repeat(i, x.End - x.Start)).ToArray();
        var newText = string.Concat(middle.Select(s => s.Text));
        var from = new int[newText.Length];
        var (oi, nj) = (0, 0);
        var spans = DocxReplace.Spans(string.Concat(region.Select(x => x.Spec.Text)).Select(c => (int)c).ToArray(), newText.Select(c => (int)c).ToArray());
        foreach (var (a, b, c, d) in spans.Append((owner.Length, owner.Length, newText.Length, newText.Length)))
        {
            for (; nj < c; oi++, nj++) from[nj] = owner[oi];
            for (; nj < d; nj++) from[nj] = owner[b > a ? a + Math.Min(nj - c, b - a - 1) : Math.Max(a - 1, 0)];
            oi = b;
        }
        var pos = 0;
        foreach (var spec in middle)
        {
            for (var i = 0; i < spec.Text.Length;)
            {
                var j = i + 1;
                while (j < spec.Text.Length && (from[pos + j] == from[pos + i] || char.IsLowSurrogate(spec.Text[j]))) j++;
                yield return (spec with { Text = spec.Text[i..j] }, region[from[pos + i]]);
                i = j;
            }
            pos += spec.Text.Length;
        }
    }

    /// <summary>Cuts a run in two at a character offset of the paragraph's text; the second half follows with the same properties.</summary>
    static void Cut(List<Piece> pieces, int offset)
    {
        if (pieces.FirstOrDefault(x => x.Start < offset && offset < x.End) is not { Element: A.Run { Text: { } text } run } piece) return;
        var tail = (A.Run)run.CloneNode(true);
        tail.Text!.Text = text.Text[(offset - piece.Start)..];
        text.Text = text.Text[..(offset - piece.Start)];
        run.InsertAfterSelf(tail);
    }

    /// <summary>Joins side-by-side runs with the same properties among new text and its neighbours, so text retyped and set back
    /// ends as the runs it started as.</summary>
    static void Merge(IEnumerable<OpenXmlElement?> sequence)
    {
        var runs = sequence.OfType<A.Run>().ToList();
        for (var i = runs.Count - 1; i > 0; i--)
        {
            var (left, right) = (runs[i - 1], runs[i]);
            if (left.Parent is null || left.NextSibling() != right || left.Text is null || right.Text is null
                || (left.RunProperties?.OuterXml ?? "") != (right.RunProperties?.OuterXml ?? "")) continue;
            left.Text.Text += right.Text.Text;
            right.Remove();
        }
    }

    /// <summary>End-of-paragraph properties as a run's: what text typed into an empty paragraph looks like.</summary>
    static A.RunProperties? AsRunProperties(A.EndParagraphRunProperties? end) => end is null ? null : Copy<A.RunProperties>(end);

    /// <summary>The same character properties under another name (a run's, an end-of-paragraph mark's).</summary>
    static T Copy<T>(A.TextCharacterPropertiesType from) where T : A.TextCharacterPropertiesType, new()
    {
        var to = new T();
        to.SetAttributes(from.GetAttributes());
        foreach (var child in from.ChildElements) to.AppendChild(child.CloneNode(true));
        return to;
    }

    /// <summary>Changes the run properties where the new text's formatting (as the source states it) differs from the run they
    /// were taken from, and only there: a language, kerning, a theme colour or a custom baseline the source cannot see stays.</summary>
    static void Restate(A.RunProperties rPr, SlidePart slide, RunSpec? from, RunSpec? to)
    {
        if (from is null || to is null || from == to) return;
        if (from.Bold != to.Bold) rPr.Bold = to.Bold;
        if (from.Italic != to.Italic) rPr.Italic = to.Italic;
        if (from.Underline != to.Underline || from.UnderlineStyle != to.UnderlineStyle) rPr.Underline = to.Underline ? UnderlineValue(to.UnderlineStyle) : A.TextUnderlineValues.None;
        if (from.Strike != to.Strike) rPr.Strike = to.Strike ? A.TextStrikeValues.SingleStrike : A.TextStrikeValues.NoStrike;
        // text with no font, colour or size of its own shows the box's (the slide editor draws those for all its text): the run
        // keeps its own rather than falling back to the layout's
        static string? FontOf(RunSpec s) => s.Font ?? (s.Code ? "Consolas" : null);
        if (FontOf(to) is { } font && font != FontOf(from)) SetFont(rPr, font);
        if (to.Color is not null && to.Color != from.Color) SetColor(rPr, to.Color);
        if (to.Size is not null && to.Size != from.Size) rPr.FontSize = Hundredths(to.Size);
        if (from.Highlight != to.Highlight) SetHighlight(rPr, to.Highlight ?? "none");
        if (from.Link != to.Link) SetLink(rPr, slide, to.Link ?? "");
        if (from.VertAlign != to.VertAlign) rPr.Baseline = BaselineOf(to.VertAlign);
        if (from.Spacing != to.Spacing) rPr.Spacing = SpacingOf(to.Spacing);
        if (from.Caps != to.Caps) rPr.Capital = CapsOf(to.Caps);
    }

    /// <summary>A run's formatting as the run node shows it.</summary>
    public static void ReadRun(A.RunProperties? rPr, SlidePart slide, Dictionary<string, string> props)
    {
        if (rPr is null) return;
        if (rPr.Bold?.Value == true) props["bold"] = "true";
        if (rPr.Italic?.Value == true) props["italic"] = "true";
        if (rPr.Underline?.InnerText is { } u && u != "none")
        {
            props["underline"] = "true";
            if (UnderlineStyleOf(u) is { } line) props["underlineStyle"] = line;
        }
        if (rPr.Strike?.InnerText is "sngStrike" or "dblStrike") props["strike"] = "true";
        if (Color(rPr) is { } color) props["color"] = color;
        if (rPr.FontSize?.Value is { } size) props["size"] = Points(size);
        if (rPr.GetFirstChild<A.LatinFont>()?.Typeface?.Value is { } font && !font.StartsWith('+')) props["font"] = font;
        if (Highlight(rPr) is { } highlight) props["highlight"] = highlight;
        if (Link(rPr, slide) is { } link) props["link"] = link;
        if (rPr.Baseline?.Value is { } baseline && baseline != 0) props["vertAlign"] = baseline > 0 ? "superscript" : "subscript";
        if (rPr.Spacing?.Value is { } spacing && spacing != 0) props["spacing"] = Points(spacing) + "pt";
        if (rPr.Capital?.InnerText is "all" or "small") props["caps"] = rPr.Capital.InnerText;
    }

    /// <summary>DrawingML's eighteen underlines as the line styles html draws (null: a single line).</summary>
    static string? UnderlineStyleOf(string u) => u == "dbl" ? "double" : u.StartsWith("wavy") ? "wavy" : u.StartsWith("dotted") ? "dotted"
        : u.StartsWith("dash") || u.StartsWith("dot") ? "dashed" : null;

    public static A.TextUnderlineValues UnderlineValue(string? style) => style switch
    {
        "double" => A.TextUnderlineValues.Double,
        "dotted" => A.TextUnderlineValues.Dotted,
        "dashed" => A.TextUnderlineValues.Dash,
        "wavy" => A.TextUnderlineValues.Wavy,
        _ => A.TextUnderlineValues.Single,
    };

    /// <summary>PowerPoint's own superscript and subscript offsets.</summary>
    public static Int32Value? BaselineOf(string? vertAlign) => vertAlign switch
    {
        "superscript" => 30000,
        "subscript" => -25000,
        _ => null,
    };

    public static Int32Value? SpacingOf(string? points) => points is null or "none" ? null : Hundredths(points.Trim().TrimEnd('t', 'p'));

    public static A.TextCapsValues CapsOf(string? caps) => caps switch
    {
        "all" => A.TextCapsValues.All,
        "small" => A.TextCapsValues.Small,
        _ => A.TextCapsValues.None,
    };

    static List<List<RunSpec>> SplitLines(IEnumerable<RunSpec> specs)
    {
        var lines = new List<List<RunSpec>> { new() };
        foreach (var spec in specs)
        {
            var parts = spec.Text.ReplaceLineEndings("\n").Split('\n');
            for (var i = 0; i < parts.Length; i++)
            {
                if (i > 0) lines.Add([]);
                if (parts[i].Length > 0) lines[^1].Add(spec with { Text = parts[i] });
            }
        }
        return lines;
    }

    public static string? Highlight(A.RunProperties? rPr) => rPr?.GetFirstChild<A.Highlight>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();

    public static void SetHighlight(A.RunProperties rPr, string value)
    {
        rPr.RemoveAllChildren<A.Highlight>();
        if (value == "none") return;
        InsertBeforeAny(rPr, new A.Highlight(new A.RgbColorModelHex { Val = value }), e => e is A.UnderlineFollowsText or A.Underline or A.UnderlineFillText
            or A.UnderlineFill or A.LatinFont or A.EastAsianFont or A.ComplexScriptFont or A.SymbolFont or A.HyperlinkOnClick or A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList);
    }

    public static string? Color(A.RunProperties? rPr) => rPr?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();

    public static void SetColor(OpenXmlCompositeElement properties, string value)
    {
        foreach (var fill in properties.ChildElements.Where(IsFill).ToList()) fill.Remove();
        if (value == "none") return;
        var fillElement = new A.SolidFill(new A.RgbColorModelHex { Val = value });
        if (properties.GetFirstChild<A.Outline>() is { } line && properties is A.TextCharacterPropertiesType) properties.InsertAfter(fillElement, line);
        else if (properties is A.TextCharacterPropertiesType or A.Outline) properties.PrependChild(fillElement); // a line's fill comes before its dash and ends
        else InsertBeforeAny(properties, fillElement, e => e is A.Outline or A.EffectList or A.EffectDag or A.Scene3DType or A.Shape3DType or A.ExtensionList);
    }

    static bool IsFill(OpenXmlElement e) => e is A.SolidFill or A.NoFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill;

    public static void SetFont(A.RunProperties rPr, string typeface)
    {
        foreach (var f in rPr.ChildElements.Where(e => e is A.LatinFont or A.EastAsianFont).ToList()) f.Remove();
        var latin = new A.LatinFont { Typeface = typeface };
        InsertBeforeAny(rPr, latin, e => e is A.ComplexScriptFont or A.SymbolFont or A.HyperlinkOnClick or A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList);
        rPr.InsertAfter(new A.EastAsianFont { Typeface = typeface }, latin);
    }

    public static void SetLink(A.RunProperties rPr, SlidePart slide, string target)
    {
        rPr.RemoveAllChildren<A.HyperlinkOnClick>();
        if (target.Length == 0) return;
        Uri uri;
        try { uri = new Uri(target, UriKind.RelativeOrAbsolute); }
        catch (UriFormatException) { throw new WriterException(ErrorCode.Validation, $"'{target}' is not a valid link", "Use a URL like https://example.com."); }
        var relationship = slide.AddHyperlinkRelationship(uri, true);
        InsertBeforeAny(rPr, new A.HyperlinkOnClick { Id = relationship.Id }, e => e is A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList);
    }

    public static string? Link(A.RunProperties? rPr, SlidePart slide)
    {
        var id = rPr?.GetFirstChild<A.HyperlinkOnClick>()?.Id?.Value;
        return id is null ? null : slide.HyperlinkRelationships.FirstOrDefault(r => r.Id == id)?.Uri.OriginalString;
    }

    /// <summary>Inserts before the first child matching the predicate, or appends.</summary>
    public static void InsertBeforeAny(OpenXmlCompositeElement parent, OpenXmlElement element, Func<OpenXmlElement, bool> after)
    {
        var anchor = parent.ChildElements.FirstOrDefault(after);
        if (anchor is null) parent.Append(element);
        else parent.InsertBefore(element, anchor);
    }

    public static string? AlignOf(A.ParagraphProperties? pPr) => pPr?.Alignment?.InnerText switch
    {
        "l" => "left",
        "ctr" => "center",
        "r" => "right",
        "just" or "justLow" => "justify",
        "dist" => "distribute",
        _ => null,
    };

    public static A.TextAlignmentTypeValues AlignValue(string align) => align switch
    {
        "center" => A.TextAlignmentTypeValues.Center,
        "right" => A.TextAlignmentTypeValues.Right,
        "justify" => A.TextAlignmentTypeValues.Justified,
        "distribute" => A.TextAlignmentTypeValues.Distributed,
        _ => A.TextAlignmentTypeValues.Left,
    };

    public static string? ListOf(A.ParagraphProperties? pPr)
    {
        if (pPr is null) return null;
        if (pPr.GetFirstChild<A.CharacterBullet>() is not null || pPr.GetFirstChild<A.PictureBullet>() is not null) return "bullet";
        if (pPr.GetFirstChild<A.AutoNumberedBullet>() is not null) return "number";
        return null;
    }

    public static void SetList(A.ParagraphProperties pPr, string kind)
    {
        foreach (var e in pPr.ChildElements.Where(e => e is A.BulletFont or A.CharacterBullet or A.AutoNumberedBullet or A.NoBullet or A.PictureBullet or A.BulletFontText).ToList()) e.Remove();
        var level = pPr.Level?.Value ?? 0;
        if (kind == "none")
        {
            pPr.LeftMargin = 0;
            pPr.Indent = 0;
            InsertBeforeAny(pPr, new A.NoBullet(), e => e is A.TabStopList or A.DefaultRunProperties or A.ExtensionList);
            return;
        }
        pPr.LeftMargin = 342900 + 457200 * level;
        pPr.Indent = -342900;
        if (kind == "bullet")
        {
            var font = new A.BulletFont { Typeface = "Arial" };
            InsertBeforeAny(pPr, font, e => e is A.TabStopList or A.DefaultRunProperties or A.ExtensionList);
            pPr.InsertAfter(new A.CharacterBullet { Char = "•" }, font);
        }
        else InsertBeforeAny(pPr, new A.AutoNumberedBullet { Type = A.TextAutoNumberSchemeValues.ArabicPeriod }, e => e is A.TabStopList or A.DefaultRunProperties or A.ExtensionList);
    }

    public static string Points(int? hundredths) => ((hundredths ?? 0) / 100.0).ToString("0.##", CultureInfo.InvariantCulture);

    public static int Hundredths(string points) => (int)Math.Round(double.Parse(points, CultureInfo.InvariantCulture) * 100);

    // ---------- the box-wide text settings the slide editor keeps per shape ----------

    /// <summary>What a text body sets for all its text: spacing on the first paragraph, character spacing and WordArt effects on
    /// the first run (or the end-of-paragraph properties), columns, autofit and direction on the body.</summary>
    public static void ReadBox(OpenXmlCompositeElement? body, Dictionary<string, string> props)
    {
        if (body is null) return;
        var bodyPr = body.GetFirstChild<A.BodyProperties>();
        if (bodyPr?.ColumnCount?.Value is { } cols && cols > 1) props["columns"] = cols.ToString(CultureInfo.InvariantCulture);
        if (bodyPr?.Vertical?.InnerText is { } vert && vert != "horz") props["direction"] = vert;
        if (bodyPr?.GetFirstChild<A.NormalAutoFit>() is { } fit) props["autofit"] = fit.FontScale?.Value is { } scale && scale != 100000 ? "shrink:" + (scale / 1000).ToString(CultureInfo.InvariantCulture) : "shrink";
        else if (bodyPr?.GetFirstChild<A.ShapeAutoFit>() is not null) props["autofit"] = "resize";
        var first = Paragraphs(body).FirstOrDefault();
        var pPr = first?.ParagraphProperties;
        if (pPr?.LineSpacing?.SpacingPercent?.Val?.Value is { } pct) props["lineSpacing"] = (pct / 100000.0).ToString("0.##", CultureInfo.InvariantCulture);
        if (pPr?.SpaceBefore?.SpacingPoints?.Val?.Value is { } before) props["spaceBefore"] = Points(before) + "pt";
        if (pPr?.SpaceAfter?.SpacingPoints?.Val?.Value is { } after) props["spaceAfter"] = Points(after) + "pt";
        var rPr = (OpenXmlCompositeElement?)first?.Elements<A.Run>().FirstOrDefault()?.RunProperties ?? first?.GetFirstChild<A.EndParagraphRunProperties>();
        if (rPr is A.TextCharacterPropertiesType t)
        {
            if (t.Spacing?.Value is { } spc && spc != 0) props["charSpacing"] = Points(spc);
            if (t.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value is { } outline) props["textOutline"] = outline.ToUpperInvariant();
            if (t.GetFirstChild<A.EffectList>()?.GetFirstChild<A.OuterShadow>() is not null) props["textShadow"] = "true";
            if (t.GetFirstChild<A.GradientFill>() is { } grad && PptxOutline.GradientOf(grad) is { } g) props["textGradient"] = g;
        }
    }

    /// <summary>True when name is one of the box-wide settings; it is then written to every paragraph and run.</summary>
    public static bool SetBox(OpenXmlCompositeElement body, string name, string value)
    {
        var bodyPr = body.GetFirstChild<A.BodyProperties>() ?? body.PrependChild(new A.BodyProperties());
        double Pt(string v) => double.TryParse(v.Trim().TrimEnd('t', 'p'), NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : throw new WriterException(ErrorCode.Validation, $"{name}: '{value}' is not a size in points", $"Example: {name}=6pt");
        switch (name)
        {
            case "columns":
                var n = int.Parse(value, CultureInfo.InvariantCulture);
                bodyPr.ColumnCount = n > 1 ? n : null;
                bodyPr.ColumnSpacing = n > 1 ? (bodyPr.ColumnSpacing?.Value ?? 457200) : null;
                return true;
            case "direction":
                bodyPr.Vertical = value == "horz" ? null : new A.TextVerticalValues(value);
                return true;
            case "autofit":
                foreach (var old in bodyPr.ChildElements.Where(e => e is A.NoAutoFit or A.NormalAutoFit or A.ShapeAutoFit).ToList()) old.Remove();
                var kind = value.Split(':')[0];
                OpenXmlElement fit = kind switch
                {
                    "none" => new A.NoAutoFit(),
                    "shrink" => new A.NormalAutoFit { FontScale = value.Contains(':') && int.TryParse(value.Split(':')[1], out var pc) && pc is > 0 and < 100 ? pc * 1000 : null },
                    "resize" => new A.ShapeAutoFit(),
                    _ => throw new WriterException(ErrorCode.Validation, $"autofit: '{value}' is not valid", "Use none, shrink, shrink:<percent> or resize."),
                };
                InsertBeforeAny(bodyPr, fit, e => e is A.Scene3DType or A.Shape3DType or A.FlatText or A.ExtensionList);
                return true;
            case "lineSpacing":
                var mult = double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture);
                foreach (var p in Paragraphs(body)) (p.ParagraphProperties ??= new A.ParagraphProperties()).LineSpacing = mult > 0 && Math.Abs(mult - 1) > 0.001 ? new A.LineSpacing(new A.SpacingPercent { Val = (int)Math.Round(mult * 100000) }) : null;
                return true;
            case "spaceBefore":
                var before = Pt(value);
                foreach (var p in Paragraphs(body)) (p.ParagraphProperties ??= new A.ParagraphProperties()).SpaceBefore = before > 0 ? new A.SpaceBefore(new A.SpacingPoints { Val = (int)Math.Round(before * 100) }) : null;
                return true;
            case "spaceAfter":
                var after = Pt(value);
                foreach (var p in Paragraphs(body)) (p.ParagraphProperties ??= new A.ParagraphProperties()).SpaceAfter = after > 0 ? new A.SpaceAfter(new A.SpacingPoints { Val = (int)Math.Round(after * 100) }) : null;
                return true;
            case "charSpacing" or "textOutline" or "textShadow" or "textGradient":
                foreach (var rPr in RunProperties(body))
                    switch (name)
                    {
                        case "charSpacing":
                            var spc = Pt(value);
                            rPr.Spacing = spc == 0 ? null : (int)Math.Round(spc * 100);
                            break;
                        case "textOutline":
                            rPr.RemoveAllChildren<A.Outline>();
                            if (value != "none") rPr.PrependChild(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = value })) { Width = 9525 });
                            break;
                        case "textShadow":
                            rPr.RemoveAllChildren<A.EffectList>();
                            if (value == "true") InsertBeforeAny(rPr, PptxOutline.Shadow(), e => e is A.Highlight or A.UnderlineFollowsText or A.Underline or A.UnderlineFillText or A.UnderlineFill or A.LatinFont or A.EastAsianFont or A.ComplexScriptFont or A.SymbolFont or A.HyperlinkOnClick or A.HyperlinkOnMouseOver or A.RightToLeft or A.ExtensionList);
                            break;
                        default:
                            var first = rPr.GetFirstChild<A.GradientFill>()?.GradientStopList?.GetFirstChild<A.GradientStop>()?.RgbColorModelHex?.Val?.Value;
                            if (value.Length == 0 || value == "none") { if (first is not null) SetColor(rPr, first); break; } // the letters keep the gradient's first colour
                            foreach (var f in rPr.ChildElements.Where(IsFill).ToList()) f.Remove();
                            var grad = PptxOutline.Gradient(value, "textGradient");
                            if (rPr.GetFirstChild<A.Outline>() is { } ln) rPr.InsertAfter(grad, ln); else rPr.PrependChild(grad);
                            break;
                    }
                return true;
        }
        return false;
    }

    /// <summary>Every run's properties and each paragraph's end-of-paragraph properties (made where missing), so a setting
    /// reaches the text there is and the text typed next.</summary>
    static IEnumerable<A.TextCharacterPropertiesType> RunProperties(OpenXmlCompositeElement body)
    {
        foreach (var p in Paragraphs(body).ToList())
        {
            foreach (var run in p.Elements<A.Run>()) yield return run.RunProperties ??= new A.RunProperties { Language = "en-US" };
            yield return p.GetFirstChild<A.EndParagraphRunProperties>() ?? p.AppendChild(new A.EndParagraphRunProperties { Language = "en-US" });
        }
    }
}
