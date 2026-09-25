using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>Style and numbering lookups for one document, plus on-demand creation of the built-in ones.</summary>
sealed class DocxStyles(WordprocessingDocument package)
{
    static readonly HashSet<string> CodeStyleIds = new(StringComparer.OrdinalIgnoreCase) { "Code", "SourceCode", "HTMLPreformatted", "CodeBlock" };

    Dictionary<string, W.Style>? _byId;

    MainDocumentPart Main => package.MainDocumentPart!;

    public W.Style? Find(string? id)
    {
        if (id is null) return null;
        _byId ??= Main.StyleDefinitionsPart?.Styles?.Elements<W.Style>()
            .Where(s => s.StyleId?.Value is not null)
            .GroupBy(s => s.StyleId!.Value!)
            .ToDictionary(g => g.Key, g => g.First()) ?? new Dictionary<string, W.Style>();
        return _byId.GetValueOrDefault(id);
    }

    public void Invalidate() => _byId = null;

    /// <summary>1-9 for headings, 0 otherwise. Checks the style name ("heading 2"), the style id ("Heading2") and outline levels.</summary>
    public int HeadingLevel(W.Paragraph p)
    {
        var pp = p.ParagraphProperties;
        var id = pp?.ParagraphStyleId?.Val?.Value;
        var style = Find(id);
        var name = style?.StyleName?.Val?.Value;
        if (name is not null && name.StartsWith("heading ", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(name.AsSpan(8), out var n) && n is >= 1 and <= 9) return n;
        if (id is not null && id.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(id.AsSpan(7), out var m) && m is >= 1 and <= 9) return m;
        var outline = pp?.OutlineLevel?.Val?.Value ?? style?.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
        return outline is >= 0 and < 9 ? outline.Value + 1 : 0;
    }

    /// <summary>A paragraph property as the paragraph's style gives it (the default paragraph style when it names none), following
    /// basedOn: the first style that sets it wins. Null when none does.</summary>
    public T? Inherited<T>(W.Paragraph p, Func<W.StyleParagraphProperties, T?> pick) where T : class
    {
        var style = Find(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value) ?? DefaultParagraphStyle();
        for (var depth = 0; style is not null && depth < 16; depth++, style = Find(style.BasedOn?.Val?.Value))
            if (style.StyleParagraphProperties is { } spp && pick(spp) is { } found) return found;
        return null;
    }

    /// <summary>The style of paragraphs that name none (Normal, 正文).</summary>
    public W.Style? DefaultParagraphStyle() =>
        Main.StyleDefinitionsPart?.Styles?.Elements<W.Style>().FirstOrDefault(s => s.Type?.Value == W.StyleValues.Paragraph && s.Default?.Value == true);

    public bool IsCode(W.Paragraph p) =>
        p.ParagraphProperties?.ParagraphStyleId?.Val?.Value is { } id && CodeStyleIds.Contains(id);

    /// <summary>("bullet" | "number" | "outline" | "chinese", level) for list paragraphs, (null, 0) otherwise. The kind is the
    /// list's: bullets; 1. a. i.; 1. 1.1 1.1.1 (outline, every level numbered with its parents' numbers); 一、（一）1. (chinese).</summary>
    public (string? List, int Level) ListInfo(W.Paragraph p)
    {
        var numPr = NumberingOf(p);
        var numId = numPr?.NumberingId?.Val?.Value;
        if (numId is null or 0) return (null, 0);
        var level = numPr!.NumberingLevelReference?.Val?.Value ?? 0;
        var abstractNum = AbstractNumOf(numId.Value);
        var levels = abstractNum?.Elements<W.Level>().ToList() ?? [];
        var format = levels.FirstOrDefault(l => l.LevelIndex?.Value == level)?.NumberingFormat?.Val?.InnerText;
        if (format == "bullet") return ("bullet", level);
        var first = levels.FirstOrDefault(l => l.LevelIndex?.Value == 0)?.NumberingFormat?.Val?.InnerText ?? "";
        if (first.StartsWith("chinese", StringComparison.Ordinal) || first.StartsWith("ideograph", StringComparison.Ordinal) || first.StartsWith("japaneseCounting", StringComparison.Ordinal)) return ("chinese", level);
        var second = levels.FirstOrDefault(l => l.LevelIndex?.Value == 1)?.LevelText?.Val?.Value ?? "";
        return (second.Contains("%1.%2", StringComparison.Ordinal) ? "outline" : "number", level);
    }

    W.NumberingProperties? NumberingOf(W.Paragraph p) => p.ParagraphProperties?.NumberingProperties
        ?? Find(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value)?.StyleParagraphProperties?.NumberingProperties;

    /// <summary>The numId a list paragraph uses, null for other paragraphs.</summary>
    public int? NumIdOf(W.Paragraph p) => NumberingOf(p)?.NumberingId?.Val?.Value is { } id and not 0 ? id : null;

    W.AbstractNum? AbstractNumOf(int numId)
    {
        var numbering = Main.NumberingDefinitionsPart?.Numbering;
        var abstractId = numbering?.Elements<W.NumberingInstance>().FirstOrDefault(x => x.NumberID?.Value == numId)?.AbstractNumId?.Val?.Value;
        return numbering?.Elements<W.AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId);
    }

    /// <summary>True when the two list paragraphs count in the same list (one numbering instance), so the second continues the first.</summary>
    public bool SameList(W.Paragraph a, W.Paragraph b) => NumIdOf(a) is { } x && NumIdOf(b) == x;

    /// <summary>A fresh instance of the paragraph's own list that starts again at 1: what Word's 重新开始编号 makes.</summary>
    public int RestartedNumId(W.Paragraph p)
    {
        var numbering = NumberingRoot();
        var abstractId = AbstractNumOf(NumIdOf(p) ?? 0)?.AbstractNumberId?.Value
            ?? throw new WriterException(ErrorCode.Validation, "restart applies to list paragraphs", "Set list=number first.");
        var numId = (numbering.Elements<W.NumberingInstance>().Select(n => n.NumberID?.Value).Max() ?? 0) + 1;
        numbering.Append(new W.NumberingInstance(new W.AbstractNumId { Val = abstractId }, new W.LevelOverride(new W.StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 }) { NumberID = numId });
        return numId;
    }

    W.Styles StylesRoot()
    {
        var part = Main.StyleDefinitionsPart ?? Main.AddNewPart<StyleDefinitionsPart>();
        return part.Styles ??= new W.Styles();
    }

    /// <summary>The id of a style given its id or name, adding the built-in definition when the document lacks it.</summary>
    public string ResolveStyle(string idOrName, string type)
    {
        var styles = StylesRoot();
        if (Find(idOrName) is { } byId)
            return TypeOf(byId) == type
                ? idOrName
                : throw new WriterException(ErrorCode.Validation, $"'{idOrName}' is a {TypeOf(byId)} style, not a {type} style", $"Pick a {type} style.");
        if (ByName(idOrName, type) is { } id) return id;
        if (Builtin(idOrName, type) is { } builtin) return builtin;
        var available = styles.Elements<W.Style>().Where(s => TypeOf(s) == type).Select(s => s.StyleId?.Value).Where(s => s is not null).Take(40);
        throw new WriterException(ErrorCode.Validation, $"No {type} style '{idOrName}'",
            $"In this document: {string.Join(", ", available)}. Built-in: {string.Join(", ", DocxTemplate.StyleIds)}.");
    }

    /// <summary>ResolveStyle, or null for a style neither the document nor the template knows (a stale id in edited html).</summary>
    public string? TryResolveStyle(string idOrName, string type)
    {
        try { return ResolveStyle(idOrName, type); }
        catch (WriterException) { return null; }
    }

    string? ByName(string? name, string type) => StylesRoot().Elements<W.Style>()
        .FirstOrDefault(s => TypeOf(s) == type && string.Equals(s.StyleName?.Val?.Value, name, StringComparison.OrdinalIgnoreCase))?.StyleId?.Value;

    /// <summary>A built-in style's id in this document: the document's own style of the same name when it has one (files saved by
    /// a localized Word use ids like "a3" for Table Grid), otherwise the template definition, added with the style it is based on.</summary>
    string? Builtin(string id, string type)
    {
        if (DocxTemplate.StyleXml(id) is not { } xml) return null;
        var style = new W.Style(xml);
        if (TypeOf(style) != type)
            throw new WriterException(ErrorCode.Validation, $"'{id}' is a {TypeOf(style)} style, not a {type} style", $"Pick a {type} style.");
        if (ByName(style.StyleName?.Val?.Value, type) is { } existing) return existing;
        if (style.BasedOn?.Val?.Value is { } basedOn && Find(basedOn) is null)
            style.BasedOn.Val = Builtin(basedOn, type) ?? basedOn;
        if (style.NextParagraphStyle?.Val?.Value is { } next && Find(next) is null && ByName(next, type) is { } nextId)
            style.NextParagraphStyle.Val = nextId;
        StylesRoot().Append(style);
        Invalidate();
        return id;
    }

    static string TypeOf(W.Style s) => s.Type?.InnerText ?? "paragraph";

    /// <summary>The style id used for headings of this level, preferring the document's own definition.</summary>
    public string HeadingStyleId(int level)
    {
        var existing = StylesRoot().Elements<W.Style>()
            .FirstOrDefault(s => string.Equals(s.StyleName?.Val?.Value, $"heading {level}", StringComparison.OrdinalIgnoreCase));
        return existing?.StyleId?.Value ?? ResolveStyle($"Heading{level}", "paragraph");
    }

    W.Numbering NumberingRoot()
    {
        var part = Main.NumberingDefinitionsPart ?? Main.AddNewPart<NumberingDefinitionsPart>();
        return part.Numbering ??= new W.Numbering();
    }

    /// <summary>A numId for a bullet or numbered list. Bullets share one instance; numbered lists continue
    /// <paramref name="continueNumId"/> when given and otherwise start fresh at 1.</summary>
    public int ListNumId(string kind, int? continueNumId)
    {
        var numbering = NumberingRoot();
        if (continueNumId is { } c && numbering.Elements<W.NumberingInstance>().Any(n => n.NumberID?.Value == c)) return c;
        var name = $"writer-{kind}";
        var abstractNum = numbering.Elements<W.AbstractNum>().FirstOrDefault(a => a.AbstractNumDefinitionName?.Val?.Value == name);
        if (abstractNum is null)
        {
            var absId = (numbering.Elements<W.AbstractNum>().Select(a => a.AbstractNumberId?.Value).Max() ?? -1) + 1;
            abstractNum = new W.AbstractNum(AbstractNumXml(kind, absId, name));
            if (numbering.Elements<W.NumberingInstance>().FirstOrDefault() is { } firstNum) numbering.InsertBefore(abstractNum, firstNum);
            else numbering.Append(abstractNum);
        }
        var abstractId = abstractNum.AbstractNumberId!.Value;
        if (kind == "bullet" && numbering.Elements<W.NumberingInstance>().FirstOrDefault(n => n.AbstractNumId?.Val?.Value == abstractId) is { } shared)
            return shared.NumberID!.Value;
        var numId = (numbering.Elements<W.NumberingInstance>().Select(n => n.NumberID?.Value).Max() ?? 0) + 1;
        var num = new W.NumberingInstance(new W.AbstractNumId { Val = abstractId }) { NumberID = numId };
        if (kind != "bullet") num.Append(new W.LevelOverride(new W.StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 });
        numbering.Append(num);
        return numId;
    }

    /// <summary>The nine levels of a list kind, as Word's galleries define them: bullets • ◦ ▪; numbers 1. a. i.; outline 1. 1.1 1.1.1
    /// (multilevel, every level with its parents' numbers); chinese 一、（一）1. (the pattern repeating every three levels).</summary>
    static string AbstractNumXml(string kind, int id, string name)
    {
        var sb = new StringBuilder();
        sb.Append($"<w:abstractNum xmlns:w=\"{DocxTemplate.Ns}\" w:abstractNumId=\"{id}\"><w:multiLevelType w:val=\"{(kind == "outline" ? "multilevel" : "hybridMultilevel")}\"/><w:name w:val=\"{name}\"/>");
        for (var level = 0; level < 9; level++)
        {
            var left = 720 * (level + 1);
            var (format, text) = kind switch
            {
                "bullet" => ("bullet", (level % 3) switch { 0 => "•", 1 => "◦", _ => "▪" }),
                "outline" => ("decimal", string.Join('.', Enumerable.Range(1, level + 1).Select(n => $"%{n}")) + (level == 0 ? "." : "")),
                "chinese" => (level % 3) switch { 0 => ("chineseCountingThousand", $"%{level + 1}、"), 1 => ("chineseCountingThousand", $"（%{level + 1}）"), _ => ("decimal", $"%{level + 1}.") },
                _ => ((level % 3) switch { 0 => "decimal", 1 => "lowerLetter", _ => "lowerRoman" }, $"%{level + 1}."),
            };
            var hanging = kind == "outline" && level > 0 ? 360 + 180 * level : 360;
            sb.Append($"<w:lvl w:ilvl=\"{level}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{format}\"/><w:lvlText w:val=\"{text}\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"{left}\" w:hanging=\"{hanging}\"/></w:pPr></w:lvl>");
        }
        sb.Append("</w:abstractNum>");
        return sb.ToString();
    }
}
