using System.Text;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>Writes new text into a paragraph in place. Old and new text are matched word by word (longest common subsequence), together
/// with the formatting the source can express, and only the stretches that differ are rewritten, each where it was. Everything that is not plain
/// text (fields, equations, pictures, footnote references, bookmarks, comment anchors, content controls, pending deletions) keeps its
/// place, and runs outside the change keep every property. With change tracking on, the rewritten words become a deletion followed by
/// an insertion; withdrawing text that is itself a pending insertion simply removes it, as in Word.</summary>
static class DocxReplace
{
    public enum Source { Text, Markdown, Html }

    sealed record Piece(W.Run Run, W.Hyperlink? Link, int Start, int End, RunSpec Spec);

    public static void Paragraph(DocxDocument doc, W.Paragraph p, IEnumerable<RunSpec> input, Source source)
    {
        var specs = input.Where(s => s.Text.Length > 0).ToList();
        var tracked = DocxRevisions.Tracking(doc) && specs.All(s => s.Change is null);
        // Html (the editor's format) always states the revisions; markdown only when it has ins / del tags; plain text never.
        var revisions = !tracked && (source == Source.Html || specs.Any(s => s.Change is not null));
        RunSpec? Key(RunSpec s)
        {
            var key = source switch
            {
                Source.Text => null,
                Source.Markdown => new RunSpec("", s.Bold, s.Italic, s.Strike, s.Code, s.Link),
                _ => s with { Text = "", Change = null, Author = null, Date = null },
            };
            // Who and when are not compared: the editor's own <ins> / <del> carry neither, and must not rewrite stored revisions.
            return revisions && key is not null ? key with { Change = s.Change } : key;
        }

        var pieces = Pieces(doc, p, revisions);
        var oldText = string.Concat(pieces.Select(x => x.Spec.Text));
        var newText = string.Concat(specs.Select(s => s.Text));
        var hunks = Hunks(oldText, Keys(pieces.Select(x => x.Spec), Key), newText, Keys(specs, Key));
        var author = DocxRevisions.Author(doc);
        var now = DocxRevisions.Now();
        var before = new HashSet<OpenXmlElement>(p.Descendants(), ReferenceEqualityComparer.Instance);
        // From the end backwards, so the offsets of the stretches still to do stay valid.
        for (var i = hunks.Count - 1; i >= 0; i--)
        {
            var (a, b, c, d) = hunks[i];
            Apply(doc, p, a, b, Take(specs, c, d), source, Key, tracked, revisions, author, now);
        }
        if (hunks.Count > 0) Tidy(p, before);
    }

    /// <summary>Puts back together what the edit cut apart, so a paragraph does not end up in more pieces with every save: a run, link or
    /// revision mark the edit made or split joins the one before it when the two look the same, and a run it left empty goes. Things
    /// that sat side by side before the edit stay as they were.</summary>
    static void Tidy(OpenXmlElement container, HashSet<OpenXmlElement> before)
    {
        foreach (var e in container.ChildElements.ToList())
        {
            var made = !before.Contains(e);
            if (made && e is W.Run && !e.ChildElements.Any(c => c is not W.RunProperties)) { e.Remove(); continue; }
            if (e.PreviousSibling() is not { } previous || !made && before.Contains(previous) || !Same(previous, e)) continue;
            if (previous is W.Run run) Join(run, (W.Run)e);
            else
            {
                foreach (var c in e.ChildElements.ToList()) { c.Remove(); previous.Append(c); }
                e.Remove();
            }
            if (!made) before.Add(previous);
        }
        foreach (var e in container.ChildElements.Where(e => e is W.Hyperlink or W.RunTrackChangeType or W.SdtRun or W.SdtContentRun or W.SimpleField or W.CustomXmlRun))
            Tidy(e, before);
    }

    static bool Same(OpenXmlElement a, OpenXmlElement b) => (a, b) switch
    {
        (W.Run x, W.Run y) => OnlyText(x) && OnlyText(y) && Look(x) == Look(y),
        (W.Hyperlink x, W.Hyperlink y) => x.Id?.Value == y.Id?.Value && x.Anchor?.Value == y.Anchor?.Value,
        (W.RunTrackChangeType x, W.RunTrackChangeType y) => x.GetType() == y.GetType() && x.Author?.Value == y.Author?.Value && x.Date?.InnerText == y.Date?.InnerText,
        _ => false,
    };

    static string Look(W.Run run) => run.RunProperties is { HasChildren: true } rp ? rp.OuterXml : "";

    static bool OnlyText(W.Run run) => run.ChildElements.All(c => c is W.RunProperties || DocxRuns.IsTextElement(c));

    /// <summary>Moves the second run's text into the first, the two w:t at the seam becoming one.</summary>
    static void Join(W.Run first, W.Run second)
    {
        foreach (var c in second.ChildElements.Where(c => c is not W.RunProperties).ToList())
        {
            c.Remove();
            if (first.LastChild is W.TextType t && c is W.TextType u && t.GetType() == u.GetType())
            {
                t.Text += u.Text;
                Preserve(t);
            }
            else first.Append(c);
        }
        second.Remove();
    }

    /// <summary>xml:space="preserve" where the text needs it (spaces at either end); an existing one stays.</summary>
    static void Preserve(W.TextType t)
    {
        if (t.Text.Length > 0 && (char.IsWhiteSpace(t.Text[0]) || char.IsWhiteSpace(t.Text[^1]))) t.Space = SpaceProcessingModeValues.Preserve;
    }

    /// <summary>Rewrites old characters [a, b) as the given runs. New text that looks like the text it replaces or continues (the same
    /// formatting key, the same link) takes that run's properties exactly, so the two can join again.</summary>
    static void Apply(DocxDocument doc, W.Paragraph p, int a, int b, List<RunSpec> middle, Source source, Func<RunSpec, RunSpec?> key, bool tracked, bool revisions, string author, string now)
    {
        var pieces = Pieces(doc, p, revisions);
        Cut(pieces, b);
        Cut(pieces, a);
        pieces = Pieces(doc, p, revisions);
        var region = pieces.Where(x => x.Start >= a && x.End <= b && x.End > x.Start).ToList();
        var left = pieces.LastOrDefault(x => x.End == a && x.End > x.Start);
        var right = pieces.FirstOrDefault(x => x.Start >= b && x.End > x.Start);
        var adjacent = region.FirstOrDefault() ?? left ?? right;
        var looks = region.Append(left).Append(right).OfType<Piece>().Select(x => (Key: key(x.Spec), x.Spec.Link, x.Run.RunProperties, Whole: OnlyText(x.Run))).ToList();
        var baseProperties = Base(adjacent, source);
        var placeholder = Placeholder(doc, p, pieces, region, a);
        var emptied = new List<OpenXmlElement>();
        W.DeletedRun? deletion = null;
        foreach (var x in region)
        {
            if (!tracked || DocxRuns.Revision(x.Run) is W.InsertedRun) Strip(x.Run, emptied);
            else foreach (var run in TextOnly(x.Run)) deletion = Delete(doc, run, deletion, author, now);
        }
        if (tracked) middle = middle.Select(s => s with { Change = "inserted", Author = author, Date = now }).ToList();
        foreach (var spec in middle)
        {
            var like = looks.FindIndex(x => x.Link == spec.Link && Equals(x.Key, key(spec)));
            if (like < 0)
            {
                placeholder.InsertBeforeSelf(DocxRuns.MakeRuns(doc, [spec], baseProperties)[0]);
                continue;
            }
            var rp = looks[like].RunProperties?.CloneNode(true) as W.RunProperties;
            // A format change stays with text typed into its run (the two join again); an insertion, or a run that stays apart, starts without it.
            if (tracked || !looks[like].Whole) rp?.RunPropertiesChange?.Remove();
            placeholder.InsertBeforeSelf(DocxRuns.MakeRun(doc, spec, rp));
        }
        placeholder.Remove();
        foreach (var e in emptied) RemoveIfEmpty(e);
    }

    static List<Piece> Pieces(DocxDocument doc, W.Paragraph p, bool deleted)
    {
        var list = new List<Piece>();
        var pos = 0;
        foreach (var (run, link) in DocxRuns.Walk(p, deleted: deleted))
        {
            var spec = RunSpec.FromProps(new DocxRun(doc, run, link).GetProps());
            list.Add(new Piece(run, link, pos, pos + spec.Text.Length, spec));
            pos += spec.Text.Length;
        }
        return list;
    }

    internal static RunSpec?[] Keys(IEnumerable<RunSpec> specs, Func<RunSpec, RunSpec?> key) =>
        specs.SelectMany(s => Enumerable.Repeat(key(s), s.Text.Length)).ToArray();

    /// <summary>The stretches that differ, as old characters [A, B) becoming new characters [C, D), in order. Tokens are words
    /// (letters and digits), CJK characters and every other character on its own, compared with their formatting keys; a word
    /// also ends where its formatting changes (H2O with a subscript 2 is three tokens).</summary>
    internal static List<(int A, int B, int C, int D)> Hunks(string oldText, RunSpec?[] oldKeys, string newText, RunSpec?[] newKeys)
    {
        var ids = new Dictionary<string, int>();
        var keyIds = new Dictionary<RunSpec, int>();
        int Id(string text, RunSpec?[] keys, (int Start, int End) t)
        {
            var sb = new StringBuilder(text, t.Start, t.End - t.Start, 64);
            for (var i = t.Start; i < t.End; i++)
                sb.Append('\u0001').Append(keys[i] is { } k ? keyIds.TryGetValue(k, out var n) ? n : keyIds[k] = keyIds.Count : -1);
            var s = sb.ToString();
            return ids.TryGetValue(s, out var id) ? id : ids[s] = ids.Count;
        }
        var oldTokens = Tokens(oldText, oldKeys);
        var newTokens = Tokens(newText, newKeys);
        var x = oldTokens.Select(t => Id(oldText, oldKeys, t)).ToArray();
        var y = newTokens.Select(t => Id(newText, newKeys, t)).ToArray();
        int At(List<(int Start, int End)> tokens, int i, int length) => i < tokens.Count ? tokens[i].Start : length;
        return Spans(x, y).Select(h => (At(oldTokens, h.A, oldText.Length), At(oldTokens, h.B, oldText.Length),
            At(newTokens, h.C, newText.Length), At(newTokens, h.D, newText.Length))).ToList();
    }

    /// <summary>Where two sequences differ, as x[A, B) becoming y[C, D), in order (longest common subsequence).</summary>
    internal static List<(int A, int B, int C, int D)> Spans(int[] x, int[] y)
    {
        var start = 0;
        while (start < x.Length && start < y.Length && x[start] == y[start]) start++;
        var end = 0;
        while (end < x.Length - start && end < y.Length - start && x[^(end + 1)] == y[^(end + 1)]) end++;
        int n = x.Length - start - end, m = y.Length - start - end;
        var spans = new List<(int A, int B, int C, int D)>();
        if ((long)n * m > 4_000_000) spans.Add((start, start + n, start, start + m)); // ponytail: quadratic match; one stretch beyond ~2000 x 2000 changed tokens
        else if (n > 0 || m > 0)
        {
            var lcs = new int[n + 1, m + 1];
            for (var i = n - 1; i >= 0; i--)
                for (var j = m - 1; j >= 0; j--)
                    lcs[i, j] = x[start + i] == y[start + j] ? lcs[i + 1, j + 1] + 1 : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);
            int oi = 0, nj = 0;
            while (oi < n || nj < m)
            {
                if (oi < n && nj < m && x[start + oi] == y[start + nj]) { oi++; nj++; continue; }
                var (i0, j0) = (oi, nj);
                while ((oi < n || nj < m) && !(oi < n && nj < m && x[start + oi] == y[start + nj]))
                {
                    if (nj < m && (oi == n || lcs[oi, nj + 1] >= lcs[oi + 1, nj])) nj++;
                    else oi++;
                }
                spans.Add((start + i0, start + oi, start + j0, start + nj));
            }
        }
        return spans;
    }

    /// <summary>Words of letters and digits (cut where the keys, when given, change); any other character (a CJK character, a
    /// surrogate pair) on its own.</summary>
    static List<(int Start, int End)> Tokens(string text, RunSpec?[]? keys = null)
    {
        var tokens = new List<(int, int)>();
        for (var i = 0; i < text.Length;)
        {
            var j = i + 1;
            if (IsWordChar(text[i])) while (j < text.Length && IsWordChar(text[j]) && (keys is null || Equals(keys[j], keys[i]))) j++;
            else if (char.IsHighSurrogate(text[i]) && j < text.Length && char.IsLowSurrogate(text[j])) j++;
            tokens.Add((i, j));
            i = j;
        }
        return tokens;
    }

    static bool IsWordChar(char c) => char.IsLetterOrDigit(c) && c < '\u2E80';

    /// <summary>The specs covering characters [from, to), cut at the ends.</summary>
    internal static List<RunSpec> Take(List<RunSpec> specs, int from, int to)
    {
        var result = new List<RunSpec>();
        var pos = 0;
        foreach (var s in specs)
        {
            var start = Math.Max(pos, from);
            var end = Math.Min(pos + s.Text.Length, to);
            if (end > start) result.Add(s with { Text = s.Text[(start - pos)..(end - pos)] });
            pos += s.Text.Length;
        }
        return result;
    }

    static void Cut(List<Piece> pieces, int offset)
    {
        if (pieces.FirstOrDefault(x => x.Start < offset && offset < x.End) is { } piece) SplitRun(piece.Run, offset - piece.Start);
    }

    static int Length(OpenXmlElement e) => e switch
    {
        W.Text or W.DeletedText => ((W.TextType)e).Text.Length,
        W.TabChar or W.Break or W.CarriageReturn or W.NoBreakHyphen => 1,
        _ => 0,
    };

    /// <summary>Cuts a run in two at a character offset of its text; the second half follows it with the same properties. An offset at
    /// either end of the text cuts nothing (no empty run is left behind).</summary>
    public static void SplitRun(W.Run run, int offset)
    {
        if (offset <= 0 || offset >= run.ChildElements.Sum(Length)) return;
        var tail = (W.Run)run.CloneNode(false);
        if (run.RunProperties is { } rp) tail.AppendChild(rp.CloneNode(true));
        var pos = 0;
        foreach (var child in run.ChildElements.Where(c => c is not W.RunProperties).ToList())
        {
            var length = Length(child);
            if (pos >= offset)
            {
                child.Remove();
                tail.AppendChild(child);
            }
            else if (pos + length > offset && child is W.Text or W.DeletedText)
            {
                var text = (W.TextType)child;
                var rest = (W.TextType)text.CloneNode(true);
                rest.Text = text.Text[(offset - pos)..];
                text.Text = text.Text[..(offset - pos)];
                Preserve(text);
                Preserve(rest);
                tail.AppendChild(rest);
            }
            pos += length;
        }
        run.InsertAfterSelf(tail);
    }

    /// <summary>The properties new runs start from: those of the text they replace or continue, less a tracked format change, a
    /// link's character style, and whatever the source states itself (it decides bold, colour and the rest).</summary>
    static W.RunProperties? Base(Piece? piece, Source source)
    {
        if (piece?.Run.RunProperties?.CloneNode(true) is not W.RunProperties rp) return null;
        rp.RunPropertiesChange?.Remove();
        if (piece.Link is not null) rp.RunStyle?.Remove();
        Func<OpenXmlElement, bool> stated = source switch
        {
            Source.Html => e => e is W.Bold or W.BoldComplexScript or W.Italic or W.ItalicComplexScript or W.Strike or W.Underline
                or W.Color or W.FontSize or W.FontSizeComplexScript or W.Shading or W.Highlight or W.RunStyle
                or W.VerticalTextAlignment or W.Spacing or W.Outline or W.Shadow,
            Source.Markdown => e => e is W.Bold or W.BoldComplexScript or W.Italic or W.ItalicComplexScript or W.Strike,
            _ => _ => false,
        };
        foreach (var e in rp.ChildElements.Where(stated).ToList()) e.Remove();
        return rp.HasChildren ? rp : null;
    }

    static readonly Func<OpenXmlElement, bool> IsProperties = e => e is W.SdtProperties or W.SdtEndCharProperties or W.CustomXmlProperties or W.FieldData;

    /// <summary>Marks where the new runs go: after the replaced text, else after the text before it (past the end of a field whose result
    /// that is), else before the text after it. The mark leaves links and revision marks, splitting them when it sits inside, so new runs
    /// carry their own; it leaves a field, content control or move (which the text cannot state) at its edge unless the replaced text
    /// was inside it.</summary>
    static W.Run Placeholder(DocxDocument doc, W.Paragraph p, List<Piece> pieces, List<Piece> region, int a)
    {
        var placeholder = new W.Run();
        OpenXmlElement anchor;
        var after = true;
        if (region.Count > 0) anchor = region[^1].Run;
        else if (pieces.LastOrDefault(x => x.End == a && x.End > x.Start) is { } before)
        {
            anchor = before.Run;
            while (anchor.NextSibling() is W.Run next && next.GetFirstChild<W.FieldChar>()?.FieldCharType?.Value == W.FieldCharValues.End) anchor = next;
        }
        else if (pieces.FirstOrDefault(x => x.End > x.Start) is { } first)
        {
            anchor = first.Run;
            after = false;
        }
        else
        {
            if (p.ParagraphProperties is { } pp) pp.InsertAfterSelf(placeholder);
            else p.PrependChild(placeholder);
            return placeholder;
        }
        while (anchor.Parent is { } parent and not W.Paragraph)
        {
            var transparent = parent is W.Hyperlink or W.InsertedRun or W.DeletedRun;
            if (!transparent && region.Any(x => x.Run.Ancestors().Contains(parent))) break;
            var edge = after ? anchor.NextSibling() is null : ReferenceEquals(parent.ChildElements.FirstOrDefault(e => !IsProperties(e)), anchor);
            if (!edge)
            {
                if (!transparent) break;
                SplitWrapper(doc, parent, anchor, after);
                after = true;
            }
            anchor = parent;
        }
        if (after) anchor.InsertAfterSelf(placeholder);
        else anchor.InsertBeforeSelf(placeholder);
        return placeholder;
    }

    /// <summary>Splits a link or revision mark at a child: the children after it (from it, when not after) move to a copy that follows.</summary>
    static void SplitWrapper(DocxDocument doc, OpenXmlElement wrapper, OpenXmlElement at, bool after)
    {
        var tail = wrapper.CloneNode(false);
        if (tail is W.RunTrackChangeType change) change.Id = doc.NextId();
        foreach (var e in wrapper.ChildElements.SkipWhile(e => !ReferenceEquals(e, at)).Skip(after ? 1 : 0).ToList())
        {
            e.Remove();
            tail.AppendChild(e);
        }
        wrapper.InsertAfterSelf(tail);
    }

    /// <summary>Removes the run's text, keeping anything else it holds; a run left with nothing goes.</summary>
    static void Strip(W.Run run, List<OpenXmlElement> emptied)
    {
        foreach (var e in run.ChildElements.Where(DocxRuns.IsTextElement).ToList()) e.Remove();
        if (run.ChildElements.Any(e => e is not W.RunProperties)) return;
        emptied.Add(run.Parent!);
        run.Remove();
    }

    static void RemoveIfEmpty(OpenXmlElement? e)
    {
        while (e is W.Hyperlink or W.InsertedRun or W.DeletedRun && e.Parent is not null && !e.HasChildren)
        {
            var parent = e.Parent;
            e.Remove();
            e = parent;
        }
    }

    /// <summary>A run holding text and something else (a picture, a field character) split into runs of one kind each; returns those with text.</summary>
    static List<W.Run> TextOnly(W.Run run)
    {
        var content = run.ChildElements.Where(e => e is not W.RunProperties).ToList();
        if (content.All(DocxRuns.IsTextElement)) return [run];
        var result = new List<W.Run>();
        OpenXmlElement cursor = run;
        W.Run? group = null;
        var groupIsText = false;
        foreach (var child in content)
        {
            var isText = DocxRuns.IsTextElement(child);
            if (group is null || isText != groupIsText)
            {
                group = (W.Run)run.CloneNode(false);
                if (run.RunProperties is { } rp) group.AppendChild(rp.CloneNode(true));
                cursor = cursor.InsertAfterSelf(group);
                groupIsText = isText;
                if (isText) result.Add(group);
            }
            child.Remove();
            group.AppendChild(child);
        }
        run.Remove();
        return result;
    }

    /// <summary>Turns a text run into a tracked deletion, joining the deletion just made when it directly precedes the run.</summary>
    static W.DeletedRun Delete(DocxDocument doc, W.Run run, W.DeletedRun? previous, string author, string now)
    {
        foreach (var t in run.Elements<W.Text>().ToList())
            run.ReplaceChild(new W.DeletedText(t.Text) { Space = SpaceProcessingModeValues.Preserve }, t);
        if (previous is not null && ReferenceEquals(run.PreviousSibling(), previous))
        {
            run.Remove();
            previous.AppendChild(run);
            return previous;
        }
        var deletion = new W.DeletedRun { Author = author, Date = new DateTimeValue { InnerText = now }, Id = doc.NextId() };
        run.InsertBeforeSelf(deletion);
        run.Remove();
        deletion.AppendChild(run);
        return deletion;
    }
}
