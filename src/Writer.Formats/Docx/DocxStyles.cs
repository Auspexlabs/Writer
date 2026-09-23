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
        var style = Find(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
            ?? Main.StyleDefinitionsPart?.Styles?.Elements<W.Style>().FirstOrDefault(s => s.Type?.Value == W.StyleValues.Paragraph && s.Default?.Value == true);
        for (var depth = 0; style is not null && depth < 16; depth++, style = Find(style.BasedOn?.Val?.Value))
            if (style.StyleParagraphProperties is { } spp && pick(spp) is { } found) return found;
        return null;
    }

    public bool IsCode(W.Paragraph p) =>
        p.ParagraphProperties?.ParagraphStyleId?.Val?.Value is { } id && CodeStyleIds.Contains(id);

    /// <summary>("bullet" | "number", level) for list paragraphs, (null, 0) otherwise.</summary>
    public (string? List, int Level) ListInfo(W.Paragraph p)
    {
        var numPr = p.ParagraphProperties?.NumberingProperties
            ?? Find(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value)?.StyleParagraphProperties?.NumberingProperties;
        var numId = numPr?.NumberingId?.Val?.Value;
        if (numId is null or 0) return (null, 0);
        var level = numPr!.NumberingLevelReference?.Val?.Value ?? 0;
        var numbering = Main.NumberingDefinitionsPart?.Numbering;
        var abstractId = numbering?.Elements<W.NumberingInstance>().FirstOrDefault(x => x.NumberID?.Value == numId)?.AbstractNumId?.Val?.Value;
        var format = numbering?.Elements<W.AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == abstractId)?
            .Elements<W.Level>().FirstOrDefault(l => l.LevelIndex?.Value == level)?.NumberingFormat?.Val?.InnerText;
        return (format == "bullet" ? "bullet" : "number", level);
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
        if (kind == "number") num.Append(new W.LevelOverride(new W.StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 });
        numbering.Append(num);
        return numId;
    }

    static string AbstractNumXml(string kind, int id, string name)
    {
        var sb = new StringBuilder();
        sb.Append($"<w:abstractNum xmlns:w=\"{DocxTemplate.Ns}\" w:abstractNumId=\"{id}\"><w:multiLevelType w:val=\"hybridMultilevel\"/><w:name w:val=\"{name}\"/>");
        for (var level = 0; level < 9; level++)
        {
            var left = 720 * (level + 1);
            var (format, text) = kind == "bullet"
                ? ("bullet", (level % 3) switch { 0 => "•", 1 => "◦", _ => "▪" })
                : ((level % 3) switch { 0 => "decimal", 1 => "lowerLetter", _ => "lowerRoman" }, $"%{level + 1}.");
            sb.Append($"<w:lvl w:ilvl=\"{level}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{format}\"/><w:lvlText w:val=\"{text}\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"{left}\" w:hanging=\"360\"/></w:pPr></w:lvl>");
        }
        sb.Append("</w:abstractNum>");
        return sb.ToString();
    }
}
