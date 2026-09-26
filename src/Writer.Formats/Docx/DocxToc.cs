using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>A table of contents: Word's "Table of Contents" content control around a TOC field, or a bare TOC field
/// spanning one or more paragraphs. One leaf block in the tree; set regenerates the entries from the current headings
/// (as a content control) and asks Word to update fields on open.</summary>
sealed class DocxToc(DocxDocument doc, List<OpenXmlElement> elements) : Node
{
    const string Gallery = "Table of Contents";
    const string Placeholder = "No table of contents entries found.";
    static readonly Regex LevelsPattern = new(@"\\o\s+""?\d-(\d)""?");
    static readonly Regex Whitespace = new(@"\s+"); // an entry is one line, as in Word: breaks and tabs in a heading become spaces
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    List<OpenXmlElement> _elements = elements;

    public override string Kind => "toc";
    public override object Anchor => _elements[0];

    W.SdtBlock? Sdt => _elements[0] as W.SdtBlock;

    IEnumerable<W.Paragraph> Paragraphs => Sdt is { } sdt ? sdt.SdtContentBlock?.Elements<W.Paragraph>() ?? [] : _elements.OfType<W.Paragraph>();

    /// <summary>Word's TOC content control, or a control holding nothing but a TOC field (and a title before it).
    /// A control that merely contains a TOC among other content stays transparent, so its other blocks remain editable.</summary>
    public static bool IsToc(W.SdtBlock sdt)
    {
        if (sdt.SdtProperties?.Descendants<W.DocPartGallery>().Any(g => g.Val?.Value == Gallery) == true) return true;
        var content = sdt.SdtContentBlock?.ChildElements.ToList() ?? [];
        var at = content.FindIndex(e => e is W.Paragraph p && HasTocField(p));
        return at is 0 or 1 && content.All(e => e is W.Paragraph) && ReferenceEquals(BareField((W.Paragraph)content[at])[^1], content[^1]);
    }

    /// <summary>A "TOC Heading" paragraph, the title Word puts right above a table of contents.</summary>
    public static bool IsTitle(DocxDocument doc, W.Paragraph p) =>
        p.ParagraphProperties?.ParagraphStyleId?.Val?.Value is { } id
        && (id == "TOCHeading" || string.Equals(doc.Styles.Find(id)?.StyleName?.Val?.Value, "TOC Heading", StringComparison.OrdinalIgnoreCase));

    /// <summary>The paragraph where a bare TOC field begins, when <paramref name="p"/> is that paragraph or the title right above it.</summary>
    public static W.Paragraph? FieldAt(DocxDocument doc, W.Paragraph p) =>
        HasTocField(p) ? p : IsTitle(doc, p) && p.NextSibling() is W.Paragraph next && HasTocField(next) ? next : null;

    /// <summary>The paragraphs of a bare TOC field, with the title paragraph right above it when there is one.</summary>
    public static List<OpenXmlElement> Group(DocxDocument doc, W.Paragraph field)
    {
        var group = BareField(field);
        if (field.PreviousSibling() is W.Paragraph title && IsTitle(doc, title)) group.Insert(0, title);
        return group;
    }

    /// <summary>True for a paragraph in which a TOC field begins.</summary>
    public static bool HasTocField(W.Paragraph p) =>
        p.Descendants<W.FieldCode>().Any(c => IsTocCode(c.Text))
        && p.Descendants<W.FieldChar>().Any(f => f.FieldCharType?.Value == W.FieldCharValues.Begin);

    static bool IsTocCode(string? code) =>
        code?.TrimStart() is { } s && s.StartsWith("TOC", StringComparison.OrdinalIgnoreCase) && (s.Length == 3 || char.IsWhiteSpace(s[3]));

    /// <summary>A bare TOC field's paragraphs: from the one holding the field begin to the one holding its end.
    /// A field that never ends among the following paragraphs is only its first paragraph, so it cannot swallow the document.</summary>
    public static List<OpenXmlElement> BareField(W.Paragraph first)
    {
        var group = new List<OpenXmlElement>();
        var depth = 0;
        for (OpenXmlElement? e = first; e is W.Paragraph p; e = e.NextSibling())
        {
            group.Add(p);
            foreach (var fc in p.Descendants<W.FieldChar>())
                depth += fc.FieldCharType?.Value == W.FieldCharValues.Begin ? 1 : fc.FieldCharType?.Value == W.FieldCharValues.End ? -1 : 0;
            if (depth <= 0) return group;
        }
        return [first];
    }

    public static (OpenXmlElement Element, string[] Consumed) New(DocxDocument doc, IReadOnlyDictionary<string, string> props)
    {
        var sdt = NewSdt();
        Fill(doc, sdt, props.TryGetValue("levels", out var l) ? int.Parse(l, Inv) : 3, props.GetValueOrDefault("title") ?? "", props.GetValueOrDefault("style") ?? "classic");
        return (sdt, ["levels", "title", "style"]);
    }

    static W.SdtBlock NewSdt() => new(
        new W.SdtProperties(new W.SdtContentDocPartObject(new W.DocPartGallery { Val = Gallery }, new W.DocPartUnique())),
        new W.SdtContentBlock());

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var paragraphs = Paragraphs.ToList();
        var field = paragraphs.FirstOrDefault(HasTocField);
        var at = field is null ? 0 : paragraphs.IndexOf(field);
        var props = new Dictionary<string, string> { ["levels"] = LevelsOf(field).ToString(Inv), ["style"] = StyleOf(field, paragraphs.Skip(at)) };
        var title = string.Join("\n", paragraphs.Take(at).Select(DocxRuns.ParagraphText).Where(t => t.Length > 0));
        if (title.Length > 0) props["title"] = title;
        props["text"] = string.Join("\n", paragraphs.Skip(at).Select(p => DocxRuns.ParagraphText(p).TrimEnd('\t')).Where(t => t.Length > 0));
        return props;
    }

    /// <summary>plain when the field leaves page numbers out (\n), simple when its entries' tab has no leader, else classic (dots).</summary>
    static string StyleOf(W.Paragraph? field, IEnumerable<W.Paragraph> entries)
    {
        var code = field?.Descendants<W.FieldCode>().Select(c => c.Text).FirstOrDefault(IsTocCode) ?? "";
        if (Regex.IsMatch(code, @"\\n(\s|$)")) return "plain";
        var leader = entries.Select(p => p.ParagraphProperties?.Tabs?.Elements<W.TabStop>().FirstOrDefault(t => t.Val?.Value == W.TabStopValues.Right)).FirstOrDefault(t => t is not null)?.Leader?.Value;
        return leader == W.TabStopLeaderCharValues.None ? "simple" : "classic";
    }

    static int LevelsOf(W.Paragraph? field)
    {
        var code = field?.Descendants<W.FieldCode>().Select(c => c.Text).FirstOrDefault(IsTocCode);
        return code is not null && LevelsPattern.Match(code) is { Success: true } m ? int.Parse(m.Groups[1].Value, Inv) : 3;
    }

    public override string GetRaw() => string.Concat(_elements.Select(e => e.OuterXml));

    /// <summary>Regenerates the block; a bare field becomes a content control at the same place.</summary>
    public override void SetProp(string name, string value)
    {
        var props = GetProps();
        var levels = name == "levels" ? int.Parse(value, Inv) : int.Parse(props["levels"], Inv);
        var title = name == "title" ? value : props.GetValueOrDefault("title") ?? "";
        var style = name == "style" ? value : props.GetValueOrDefault("style") ?? "classic";
        var sdt = Sdt;
        if (sdt is null)
        {
            sdt = NewSdt();
            _elements[0].Parent!.InsertBefore(sdt, _elements[0]);
            foreach (var e in _elements) e.Remove();
            _elements = [sdt];
        }
        Fill(doc, sdt, levels, title, style);
    }

    /// <summary>The entries of the headings up to `levels`, under the title; style says how an entry ends (see StyleOf).</summary>
    static void Fill(DocxDocument doc, W.SdtBlock sdt, int levels, string title, string style)
    {
        var content = sdt.SdtContentBlock ??= new W.SdtContentBlock();
        content.RemoveAllChildren();
        if (title.Length > 0)
            content.Append(new W.Paragraph(
                new W.ParagraphProperties(new W.ParagraphStyleId { Val = doc.Styles.ResolveStyle("TOCHeading", "paragraph") }),
                new W.Run(DocxRuns.TextElements(title))));
        var pos = DocxTable.ContentWidthTwips(doc);
        var first = true;
        foreach (var (text, level, anchor) in Headings(doc, sdt, levels))
        {
            var p = new W.Paragraph(new W.ParagraphProperties(
                new W.ParagraphStyleId { Val = doc.Styles.ResolveStyle($"TOC{level}", "paragraph") },
                new W.Tabs(new W.TabStop { Val = W.TabStopValues.Right, Leader = style == "simple" ? W.TabStopLeaderCharValues.None : W.TabStopLeaderCharValues.Dot, Position = pos })));
            if (first)
            {
                p.Append(FieldChar(W.FieldCharValues.Begin), Code($" TOC \\o \"1-{levels}\" \\h \\z \\u {(style == "plain" ? "\\n " : "")}"), FieldChar(W.FieldCharValues.Separate));
                first = false;
            }
            p.Append(anchor is null ? new W.Run(DocxRuns.TextElements(text)) : Entry(text, anchor, style != "plain"));
            content.Append(p);
        }
        content.Append(new W.Paragraph(FieldChar(W.FieldCharValues.End)));
        UpdateFieldsOnOpen(doc);
    }

    static W.Run FieldChar(W.FieldCharValues type) => new(new W.FieldChar { FieldCharType = type });

    static W.Run Code(string instruction) => new(new W.FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve });

    /// <summary>heading text, then (with page) the tab to the number and a PAGEREF field Word fills with it; the whole entry links to the heading.</summary>
    static W.Hyperlink Entry(string text, string anchor, bool page) => page
        ? new(new W.Run(DocxRuns.TextElements(text)), new W.Run(new W.TabChar()),
            FieldChar(W.FieldCharValues.Begin), Code($" PAGEREF {anchor} \\h "), FieldChar(W.FieldCharValues.Separate), FieldChar(W.FieldCharValues.End)) { Anchor = anchor, History = true }
        : new(new W.Run(DocxRuns.TextElements(text))) { Anchor = anchor, History = true };

    /// <summary>Headings of the wanted levels with a _Toc bookmark on each (added when missing); the placeholder when there are none.</summary>
    static IEnumerable<(string Text, int Level, string? Anchor)> Headings(DocxDocument doc, W.SdtBlock self, int levels)
    {
        var document = doc.Main.Document!;
        var found = document.Body!.Descendants<W.Paragraph>()
            .Where(p => !p.Ancestors().Contains(self))
            .Select(p => (Paragraph: p, Level: doc.Styles.HeadingLevel(p), Text: Whitespace.Replace(DocxRuns.ParagraphText(p), " ").Trim()))
            .Where(x => x.Level >= 1 && x.Level <= levels && x.Text.Length > 0)
            .ToList();
        if (found.Count == 0)
        {
            yield return (Placeholder, 1, null);
            yield break;
        }
        var starts = document.Descendants<W.BookmarkStart>().ToList();
        var names = starts.Select(b => b.Name?.Value).OfType<string>().ToHashSet();
        var nextId = starts.Select(b => int.TryParse(b.Id?.Value, NumberStyles.Integer, Inv, out var i) ? i : 0).DefaultIfEmpty(0).Max() + 1;
        foreach (var (p, level, text) in found)
        {
            var name = p.Elements<W.BookmarkStart>().Select(b => b.Name?.Value).FirstOrDefault(n => n?.StartsWith("_Toc", StringComparison.Ordinal) == true);
            if (name is null)
            {
                var n = nextId;
                while (names.Contains($"_Toc{n}")) n++;
                name = $"_Toc{n}";
                names.Add(name);
                var id = nextId.ToString(Inv);
                nextId++;
                var start = new W.BookmarkStart { Id = id, Name = name };
                if (p.ParagraphProperties is { } pp) p.InsertAfter(start, pp);
                else p.PrependChild(start);
                p.Append(new W.BookmarkEnd { Id = id });
            }
            yield return (text, level, name);
        }
    }

    static void UpdateFieldsOnOpen(DocxDocument doc)
    {
        var part = doc.Main.DocumentSettingsPart ?? doc.Main.AddNewPart<DocumentSettingsPart>();
        var settings = part.Settings ??= new W.Settings();
        settings.GetFirstChild<W.UpdateFieldsOnOpen>()?.Remove();
        settings.AddChild(new W.UpdateFieldsOnOpen { Val = true });
    }

    public override void Remove()
    {
        foreach (var e in _elements) DocxBlocks.Detach(e);
    }

    public override void MoveTo(Node newParent, int? index)
    {
        DocxBlocks.Move(newParent, _elements[0], index);
        for (var i = 1; i < _elements.Count; i++)
        {
            _elements[i].Remove();
            _elements[i - 1].InsertAfterSelf(_elements[i]);
        }
    }

    public override void SetRaw(string raw)
    {
        foreach (var e in _elements.Skip(1)) e.Remove();
        DocxBlocks.ReplaceRaw(doc, _elements[0], raw);
    }
}
