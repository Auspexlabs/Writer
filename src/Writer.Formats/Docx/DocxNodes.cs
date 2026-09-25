using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

sealed class DocxRoot(DocxDocument doc) : Node
{
    public override string Kind => "document";
    public override string Format => "docx";
    public override object Anchor => doc.Package;

    protected override IEnumerable<Node> ProjectChildren() => [new DocxBody(doc, doc.Main.Document!.Body!)];

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["format"] = "docx" };
        if (doc.Package.PackageProperties.Title is { Length: > 0 } title) props["title"] = title;
        DocxSection.Read(doc, props);
        if (DocxRevisions.Tracking(doc)) props["track"] = "true";
        if (doc.Main.DocumentSettingsPart?.Settings?.GetFirstChild<W.AutoHyphenation>() is { } hyphenation && DocxRun.On(hyphenation)) props["hyphenation"] = "true";
        if (DocxRevisions.DefaultAuthor(doc) is { } author) props["author"] = author;
        props["revisions"] = DocxRevisions.Count(doc).ToString(CultureInfo.InvariantCulture);
        props["comments"] = DocxComments.Count(doc).ToString(CultureInfo.InvariantCulture);
        props["styles"] = DocxStyleGallery.Read(doc);
        return props;
    }

    /// <summary>The document's own look (its default paragraph style over docDefaults, theme fonts resolved) and its last section's page as lengths.</summary>
    public override IReadOnlyDictionary<string, string>? GetComputed(IReadOnlyDictionary<string, string> props)
    {
        var computed = DocxLook.Of(doc, null);
        if (doc.Main.Document!.Body!.GetFirstChild<W.SectionProperties>() is { } section) DocxSection.ReadGeometry(section, computed);
        return computed;
    }

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "title": doc.Package.PackageProperties.Title = value.Length > 0 ? value : null; break;
            case "style": DocxStyleGallery.Define(doc, value); break;
            case "page": DocxSection.SetPage(doc, value); break;
            case "orientation": DocxSection.SetOrientation(doc, value); break;
            case "margin": DocxSection.SetMargin(doc, value); break;
            case "columns": DocxSection.SetColumns(doc, value); break;
            case "header": DocxSection.SetHeaderFooter(doc, header: true, first: false, value); break;
            case "footer": DocxSection.SetHeaderFooter(doc, header: false, first: false, value); break;
            case "firstHeader": DocxSection.SetHeaderFooter(doc, header: true, first: true, value); break;
            case "firstFooter": DocxSection.SetHeaderFooter(doc, header: false, first: true, value); break;
            case "titlePg": DocxSection.SetTitlePage(doc, value == "true"); break;
            case "lineNumbers": DocxSection.SetLineNumbers(doc, value == "true"); break;
            case "hyphenation":
                var settings = (doc.Main.DocumentSettingsPart ?? doc.Main.AddNewPart<DocumentSettingsPart>()).Settings ??= new W.Settings();
                settings.RemoveAllChildren<W.AutoHyphenation>();
                if (value == "true") settings.AddChild(new W.AutoHyphenation());
                break;
            case "track": DocxRevisions.SetTracking(doc, value == "true"); break;
            case "author": DocxRevisions.SetDefaultAuthor(doc, value); break;
            case "accept": DocxRevisions.Resolve(doc.Main.Document!.Body!, accept: true); break;
            case "reject": DocxRevisions.Resolve(doc.Main.Document!.Body!, accept: false); break;
        }
    }
}

sealed class DocxBody(DocxDocument doc, W.Body body) : Node, IDocxContainer
{
    public override string Kind => "body";
    public override object Anchor => body;
    public OpenXmlElement Container => body;
    protected override IEnumerable<Node> ProjectChildren() => DocxBlocks.Project(doc, body);
    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string>();
    public override string GetRaw() => body.OuterXml;
    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index) => DocxBlocks.Add(doc, this, body, kind, props, index);
    public override void SetRaw(string raw) => DocxBlocks.ReplaceRaw(doc, body, raw);
}

/// <summary>A paragraph, projected as heading, code or paragraph depending on its style.</summary>
sealed class DocxParagraph(DocxDocument doc, W.Paragraph p) : Node, IDocxContainer
{
    public override object Anchor => p;
    public OpenXmlElement Container => p;
    public override string Kind => doc.Styles.HeadingLevel(p) > 0 ? "heading" : doc.Styles.IsCode(p) ? "code" : "paragraph";

    protected override IEnumerable<Node> ProjectChildren() =>
        Kind == "code" ? [] : DocxRuns.Walk(p, deleted: true).Select(x => (Node)new DocxRun(doc, x.Run, x.Link))
            .Concat(DocxImage.In(p).Select(d => (Node)new DocxImage(doc, d))).Concat(DocxComments.In(doc, p)).Concat(DocxFootnotes.In(doc, p)).Concat(DocxEquation.In(p)).Concat(DocxShape.In(p));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = DocxRuns.ParagraphText(p) };
        if (Kind != "code") props["html"] = Exporter.HtmlOf(this);
        var level = doc.Styles.HeadingLevel(p);
        if (level > 0) props["level"] = level.ToString(CultureInfo.InvariantCulture);
        else if (!doc.Styles.IsCode(p))
        {
            if (p.ParagraphProperties?.ParagraphStyleId?.Val?.Value is { } style) props["style"] = style;
            var (list, listLevel) = doc.Styles.ListInfo(p);
            if (list is not null)
            {
                props["list"] = list;
                props["level"] = listLevel.ToString(CultureInfo.InvariantCulture);
                // numbering that starts again here: the item before it counts in another list of the same kind
                if (list != "bullet" && p.PreviousSibling<W.Paragraph>() is { } previous && doc.Styles.ListInfo(previous).List == list && !doc.Styles.SameList(previous, p))
                    props["restart"] = "true";
            }
        }
        if (AlignOf(p.ParagraphProperties) is { } align) props["align"] = align;
        if (Kind != "code" && p.ParagraphProperties?.PageBreakBefore is { } pageBreak) props["pageBreakBefore"] = DocxRun.On(pageBreak) ? "true" : "false";
        if (Kind != "code" && FillOf(p.ParagraphProperties?.Shading) is { } fill) props["fill"] = fill;
        if (Kind != "code") DocxParaFormat.Read(p.ParagraphProperties, props);
        if (DocxMarks.Bookmark(p) is { } bookmark) props["bookmark"] = bookmark;
        if (Kind != "code" && DocxMarks.Caption(p) is { } caption) props["caption"] = caption;
        if (DocxMarks.DropCap(p) is { } dropCap) props["dropCap"] = dropCap;
        if (DocxSection.SectionBreakOf(p) is { } sectionBreak) { props["sectionBreak"] = sectionBreak; DocxSection.ReadPage(p.ParagraphProperties!.SectionProperties!, props); }
        if (p.ParagraphId?.Value is { } id) props["id"] = id;
        return props;
    }

    /// <summary>What the paragraph's style gives where the paragraph itself is silent: a page break before it, its shading, and the
    /// spacing, indents, font, size and colour it has beyond the document's own look (the root's computed); for a paragraph that ends a
    /// section, that section's page as lengths.</summary>
    public override IReadOnlyDictionary<string, string>? GetComputed(IReadOnlyDictionary<string, string> props)
    {
        if (Kind == "code") return null;
        var computed = new Dictionary<string, string>();
        if (!props.ContainsKey("pageBreakBefore") && DocxRun.On(doc.Styles.Inherited(p, s => s.PageBreakBefore))) computed["pageBreakBefore"] = "true";
        if (p.ParagraphProperties?.Shading is null && FillOf(doc.Styles.Inherited(p, s => s.Shading)) is { } fill) computed["fill"] = fill;
        foreach (var (key, value) in DocxLook.Differences(doc, p.ParagraphProperties?.ParagraphStyleId?.Val?.Value, props)) computed[key] = value;
        if (p.ParagraphProperties?.SectionProperties is { } section) DocxSection.ReadGeometry(section, computed);
        return computed.Count > 0 ? computed : null;
    }

    static string? FillOf(W.Shading? shading) =>
        shading?.Fill?.Value is { } fill && !fill.Equals("auto", StringComparison.OrdinalIgnoreCase) ? fill.ToUpperInvariant() : null;

    internal static string? AlignOf(W.ParagraphProperties? pp) => AlignOf(pp?.Justification);

    internal static string? AlignOf(W.Justification? jc) => jc?.Val?.InnerText switch
    {
        "left" or "start" => "left",
        "center" => "center",
        "right" or "end" => "right",
        "both" => "justify",
        "distribute" => "distribute",
        _ => null,
    };

    internal static W.Justification JustificationOf(string align) => new()
    {
        Val = align switch
        {
            "center" => W.JustificationValues.Center,
            "right" => W.JustificationValues.Right,
            "justify" => W.JustificationValues.Both,
            "distribute" => W.JustificationValues.Distribute,
            _ => W.JustificationValues.Left,
        },
    };

    public override string GetRaw() => p.OuterXml;

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "text": DocxReplace.Paragraph(doc, p, [new RunSpec(value)], DocxReplace.Source.Text); break;
            case "md": DocxReplace.Paragraph(doc, p, InlineMarkdown.Parse(value), DocxReplace.Source.Markdown); break;
            case "html": DocxReplace.Paragraph(doc, p, InlineHtml.Parse(value, pageBreaks: true), DocxReplace.Source.Html); break;
            case "accept": DocxRevisions.Resolve(p, accept: true); break;
            case "reject": DocxRevisions.Resolve(p, accept: false); break;
            case "style":
                Properties().ParagraphStyleId = new W.ParagraphStyleId { Val = doc.Styles.ResolveStyle(value, "paragraph") };
                break;
            case "level" when Kind == "heading":
                Properties().ParagraphStyleId = new W.ParagraphStyleId { Val = doc.Styles.HeadingStyleId(int.Parse(value, CultureInfo.InvariantCulture)) };
                break;
            case "level":
                var numPr = p.ParagraphProperties?.NumberingProperties
                    ?? throw new WriterException(ErrorCode.Validation, "level applies to list paragraphs and headings",
                        "Set list=bullet or list=number first.");
                numPr.NumberingLevelReference = new W.NumberingLevelReference { Val = int.Parse(value, CultureInfo.InvariantCulture) };
                break;
            case "list":
                SetList(value);
                break;
            case "restart":
                SetRestart(value == "true");
                break;
            case "align":
                Properties().Justification = JustificationOf(value);
                break;
            case "pageBreakBefore": // off says so only where the style would turn it on
                Properties().PageBreakBefore = value == "true" ? new W.PageBreakBefore()
                    : DocxRun.On(doc.Styles.Inherited(p, s => s.PageBreakBefore)) ? new W.PageBreakBefore { Val = false } : null;
                break;
            case "fill": // none says so only where the style would shade it
                Properties().Shading = value != "none" ? new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = value }
                    : doc.Styles.Inherited(p, s => s.Shading) is not null ? new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = "auto" } : null;
                break;
            case "lineSpacing" or "spaceBefore" or "spaceAfter" or "indentLeft" or "indentRight" or "indentFirst" or "border" or "keepNext" or "keepLines" or "tabs":
                DocxParaFormat.Write(Properties(), name, value);
                if (!p.ParagraphProperties!.HasChildren) p.ParagraphProperties.Remove();
                break;
            case "sectionBreak": DocxSection.SetSectionBreak(doc, p, value); break;
            case "bookmark": DocxMarks.SetBookmark(doc, p, value); break;
            case "caption": DocxMarks.SetCaption(doc, p, value); break;
            case "dropCap": DocxMarks.SetDropCap(p, value); break;
            case "page" or "orientation" or "margin" or "columns":
                var section = p.ParagraphProperties?.SectionProperties
                    ?? throw new WriterException(ErrorCode.Validation, $"{name} applies to a paragraph that ends a section", "Set sectionBreak=nextPage or continuous first; the document's own page setup is on /.");
                if (name == "page") DocxSection.SetPage(section, value); else if (name == "orientation") DocxSection.SetOrientation(section, value); else if (name == "margin") DocxSection.SetMargin(section, value); else DocxSection.SetColumns(section, value);
                break;
        }
    }

    W.ParagraphProperties Properties() => p.ParagraphProperties ??= new W.ParagraphProperties();

    void SetList(string kind)
    {
        var pp = Properties();
        if (kind == "none")
        {
            pp.NumberingProperties?.Remove();
            if (pp.ParagraphStyleId?.Val?.Value == "ListParagraph") pp.ParagraphStyleId.Remove();
            return;
        }
        var level = pp.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
        if (doc.Styles.ListInfo(p).List == kind) return; // already in a list of this kind: it keeps counting where it is
        int? continueFrom = null;
        if (kind != "bullet" && p.PreviousSibling<W.Paragraph>() is { } previous && doc.Styles.ListInfo(previous).List == kind)
            continueFrom = doc.Styles.NumIdOf(previous);
        var numId = doc.Styles.ListNumId(kind, continueFrom);
        pp.NumberingProperties = new W.NumberingProperties(
            new W.NumberingLevelReference { Val = level },
            new W.NumberingId { Val = numId });
        if (pp.ParagraphStyleId is null || pp.ParagraphStyleId.Val?.Value == "Normal")
            pp.ParagraphStyleId = new W.ParagraphStyleId { Val = doc.Styles.ResolveStyle("ListParagraph", "paragraph") };
    }

    /// <summary>重新开始编号 / 继续编号: this item and the ones after it in its list count in a new list from 1, or join the list of
    /// the item before them.</summary>
    void SetRestart(bool on)
    {
        var old = doc.Styles.NumIdOf(p) ?? throw new WriterException(ErrorCode.Validation, "restart applies to list paragraphs", "Set list=number first.");
        int numId;
        if (on) numId = doc.Styles.RestartedNumId(p);
        else if (p.PreviousSibling<W.Paragraph>() is { } previous && doc.Styles.ListInfo(previous).List == doc.Styles.ListInfo(p).List) numId = doc.Styles.NumIdOf(previous)!.Value;
        else return;
        for (var q = p; q is not null && doc.Styles.NumIdOf(q) == old; q = q.NextSibling<W.Paragraph>())
            Renumber(q, numId);
    }

    static void Renumber(W.Paragraph q, int numId)
    {
        var pp = q.ParagraphProperties ??= new W.ParagraphProperties();
        var numPr = pp.NumberingProperties ??= new W.NumberingProperties(new W.NumberingLevelReference { Val = 0 });
        numPr.NumberingId = new W.NumberingId { Val = numId };
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        if (kind == "comment") return DocxComments.Add(doc, p, props);
        if (kind == "footnote") return DocxFootnotes.Add(doc, p, props);
        if (kind == "equation") return DocxEquation.Add(p, props);
        if (kind == "shape") return DocxShape.Add(doc, p, props);
        if (kind == "image") return DocxImage.AddTo(doc, this, p, props, index);
        if (!props.ContainsKey("text") && !props.ContainsKey("md"))
            throw new WriterException(ErrorCode.Validation, "A run needs text", "Add --prop text=\"...\" or --prop md=\"...\".");
        var run = new W.Run();
        DocxBlocks.InsertAt(this, p, run, index);
        var node = new DocxRun(doc, run, null);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    /// <summary>Removes the paragraph and, as Word does, the comments and notes anchored in it.</summary>
    public override void Remove()
    {
        foreach (var comment in DocxComments.In(doc, p).ToList()) comment.Remove();
        foreach (var note in DocxFootnotes.In(doc, p).ToList()) note.Remove();
        DocxBlocks.Detach(p);
    }

    public override void MoveTo(Node newParent, int? index) => DocxBlocks.Move(newParent, p, index);
    public override void SetRaw(string raw) => DocxBlocks.ReplaceRaw(doc, p, raw);
}

/// <summary>A paragraph whose only content is a page break.</summary>
sealed class DocxPageBreak(DocxDocument doc, W.Paragraph p) : Node
{
    public override string Kind => "pagebreak";
    public override object Anchor => p;

    public static bool IsPageBreak(W.Paragraph p)
    {
        var content = p.Elements<W.Run>().SelectMany(r => r.ChildElements).Where(c => c is not (W.RunProperties or W.LastRenderedPageBreak)).ToList();
        return content.Count > 0
            && content.All(c => c is W.Break b && b.Type?.Value == W.BreakValues.Page)
            && p.ChildElements.All(c => c is W.ParagraphProperties or W.Run);
    }

    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string>();
    public override string GetRaw() => p.OuterXml;
    public override void Remove() => DocxBlocks.Detach(p);
    public override void MoveTo(Node newParent, int? index) => DocxBlocks.Move(newParent, p, index);
    public override void SetRaw(string raw) => DocxBlocks.ReplaceRaw(doc, p, raw);
}

sealed class DocxRun(DocxDocument doc, W.Run run, W.Hyperlink? link) : Node
{
    public override string Kind => "run";
    public override object Anchor => run;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = DocxRuns.RunText(run) };
        var rp = run.RunProperties;
        if (On(rp?.Bold)) props["bold"] = "true";
        if (On(rp?.Italic)) props["italic"] = "true";
        if (rp?.Underline is { } u && u.Val?.InnerText != "none") props["underline"] = "true";
        if (On(rp?.Strike)) props["strike"] = "true";
        if (rp?.Color?.Val?.Value is { } color && !color.Equals("auto", StringComparison.OrdinalIgnoreCase)) props["color"] = color.ToUpperInvariant();
        if (rp?.FontSize?.Val?.Value is { } half && double.TryParse(half, NumberStyles.Float, CultureInfo.InvariantCulture, out var halfPoints))
            props["size"] = (halfPoints / 2).ToString("0.##", CultureInfo.InvariantCulture);
        if (rp?.RunFonts?.Ascii?.Value is { } font) props["font"] = font;
        if (rp?.Shading?.Fill?.Value is { } fill && !fill.Equals("auto", StringComparison.OrdinalIgnoreCase)) props["highlight"] = fill.ToUpperInvariant();
        else if (rp?.Highlight?.Val?.InnerText is { } highlight && DocxRuns.HighlightHex(highlight) is { } named) props["highlight"] = named;
        if (rp?.VerticalTextAlignment?.Val?.Value is { } vertical && vertical != W.VerticalPositionValues.Baseline) props["vertAlign"] = vertical == W.VerticalPositionValues.Superscript ? "superscript" : "subscript";
        if (rp?.Spacing?.Val?.Value is { } spacing && spacing != 0) props["spacing"] = (spacing / 20.0).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
        if (On(rp?.Outline)) props["outline"] = "true";
        if (On(rp?.Shadow)) props["shadow"] = "true";
        var currentLink = link ?? run.Parent as W.Hyperlink;
        if (currentLink is not null && LinkTarget(doc, currentLink) is { } target) props["link"] = target;
        if (rp?.RunStyle?.Val?.Value is { Length: > 0 } characterStyle && !characterStyle.Equals("Hyperlink", StringComparison.OrdinalIgnoreCase)) props["style"] = characterStyle;
        if (DocxRuns.Revision(run) is W.RunTrackChangeType change)
        {
            props["change"] = change is W.DeletedRun ? "deleted" : "inserted";
            if (change.Author?.Value is { Length: > 0 } author) props["author"] = author;
            if (change.Date?.InnerText is { Length: > 0 } date) props["date"] = date;
        }
        return props;
    }

    /// <summary>The fonts the run's own w:rFonts name through the theme (asciiTheme, eastAsiaTheme), and its East Asian font where it is not its font.</summary>
    public override IReadOnlyDictionary<string, string>? GetComputed(IReadOnlyDictionary<string, string> props)
    {
        if (run.RunProperties?.RunFonts is not { } fonts) return null;
        var computed = new Dictionary<string, string>();
        if (!props.ContainsKey("font") && DocxLook.Font(doc, fonts.AsciiTheme?.InnerText, null) is { } latin) computed["font"] = latin;
        if (DocxLook.Font(doc, fonts.EastAsiaTheme?.InnerText, fonts.EastAsia?.Value) is { } eastAsia && eastAsia != (props.GetValueOrDefault("font") ?? computed.GetValueOrDefault("font"))) computed["fontEa"] = eastAsia;
        return computed.Count > 0 ? computed : null;
    }

    internal static bool On(W.OnOffType? t) => t is not null && (t.Val is null || t.Val.Value);

    internal static string? LinkTarget(DocxDocument doc, W.Hyperlink hyperlink)
    {
        if (hyperlink.Id?.Value is { } id)
            return doc.Main.HyperlinkRelationships.FirstOrDefault(r => r.Id == id)?.Uri.OriginalString;
        return hyperlink.Anchor?.Value is { } anchor ? "#" + anchor : null;
    }

    public override string GetRaw() => run.OuterXml;

    public override void SetProp(string name, string value)
    {
        if (name == "text")
        {
            DocxRuns.SetRunText(run, value);
            return;
        }
        if (name == "md")
        {
            var specs = InlineMarkdown.Parse(value);
            DocxRuns.SetRunText(run, string.Concat(specs.Select(s => s.Text)));
            return;
        }
        if (name == "link")
        {
            SetLink(value);
            return;
        }
        var rp = run.RunProperties ??= new W.RunProperties();
        var on = value == "true";
        switch (name)
        {
            case "bold": rp.Bold = on ? new W.Bold() : new W.Bold { Val = false }; break;
            case "italic": rp.Italic = on ? new W.Italic() : new W.Italic { Val = false }; break;
            case "strike": rp.Strike = on ? new W.Strike() : new W.Strike { Val = false }; break;
            case "underline": rp.Underline = new W.Underline { Val = on ? W.UnderlineValues.Single : W.UnderlineValues.None }; break;
            case "color": rp.Color = value == "none" ? null : new W.Color { Val = value }; break;
            case "size":
                var half = ((int)Math.Round(double.Parse(value, CultureInfo.InvariantCulture) * 2)).ToString(CultureInfo.InvariantCulture);
                rp.FontSize = new W.FontSize { Val = half };
                rp.FontSizeComplexScript = new W.FontSizeComplexScript { Val = half };
                break;
            case "font": rp.RunFonts = new W.RunFonts { Ascii = value, HighAnsi = value, EastAsia = value, ComplexScript = value }; break;
            case "highlight": DocxRuns.SetHighlight(rp, value == "none" ? null : value); break;
            case "vertAlign": rp.VerticalTextAlignment = value is "baseline" or "none" ? null : new W.VerticalTextAlignment { Val = value == "superscript" ? W.VerticalPositionValues.Superscript : W.VerticalPositionValues.Subscript }; break;
            case "spacing": rp.Spacing = value == "none" ? null : new W.Spacing { Val = (int)Math.Round(double.Parse(value.TrimEnd('p', 't', ' '), CultureInfo.InvariantCulture) * 20) }; break;
            case "outline": rp.Outline = on ? new W.Outline() : null; break;
            case "shadow": rp.Shadow = on ? new W.Shadow() : null; break;
            case "style": rp.RunStyle = value is "none" or "" ? null : new W.RunStyle { Val = doc.Styles.ResolveStyle(value, "character") }; break;
        }
        if (!rp.HasChildren) rp.Remove();
    }

    void SetLink(string target)
    {
        var current = run.Parent as W.Hyperlink;
        if (target.Length == 0)
        {
            if (current is null) return;
            run.Remove();
            current.Parent!.InsertBefore(run, current);
            if (!current.HasChildren) current.Remove();
            run.RunProperties?.RunStyle?.Remove();
            return;
        }
        if (current is not null)
        {
            if (target.StartsWith('#'))
            {
                current.Id = null;
                current.Anchor = target[1..];
            }
            else
            {
                var replacement = DocxRuns.MakeHyperlink(doc, target, new W.Run());
                current.Id = replacement.Id;
                current.Anchor = null;
            }
        }
        else
        {
            var parent = run.Parent!;
            var next = run.NextSibling();
            run.Remove();
            var hyperlink = DocxRuns.MakeHyperlink(doc, target, run);
            if (next is null) parent.Append(hyperlink);
            else parent.InsertBefore(hyperlink, next);
        }
        var rp = run.RunProperties ??= new W.RunProperties();
        rp.RunStyle = new W.RunStyle { Val = doc.Styles.ResolveStyle("Hyperlink", "character") };
    }

    public override void Remove()
    {
        var parent = run.Parent;
        run.Remove();
        if (parent is W.Hyperlink h && !h.Elements<W.Run>().Any()) h.Remove();
    }

    public override void MoveTo(Node newParent, int? index) => DocxBlocks.Move(newParent, run, index);
    public override void SetRaw(string raw) => DocxBlocks.ReplaceRaw(doc, run, raw);
}
