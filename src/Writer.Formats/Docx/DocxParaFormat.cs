using System.Globalization;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>Word's 段落 dialog on a paragraph's own properties: line spacing, space before and after, indents (lengths or
/// characters, as CJK Word counts them), borders, keep with next / keep lines together, and tab stops. Values are read and
/// written in the same words: 1.5 or 18pt or min 18pt; 6pt; 2ch or 0.74cm, negative for a hanging indent; top bottom left right
/// or box; left 2cm, center 8cm, right 15cm, decimal 10cm.</summary>
static class DocxParaFormat
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    const long EmuPerTwip = 635;
    public static readonly string[] Names = ["lineSpacing", "spaceBefore", "spaceAfter", "indentLeft", "indentRight", "indentFirst", "border", "keepNext", "keepLines", "tabs"];

    /// <summary>Reads the settings of a paragraph's pPr, or of a style's (the same children under another parent).</summary>
    public static void Read(OpenXmlElement? pp, Dictionary<string, string> props)
    {
        if (pp is null) return;
        if (pp.GetFirstChild<W.SpacingBetweenLines>() is { } sp)
        {
            if (LineSpacingOf(sp) is { } line) props["lineSpacing"] = line;
            if (sp.Before?.Value is { } before) props["spaceBefore"] = Pt(before);
            if (sp.After?.Value is { } after) props["spaceAfter"] = Pt(after);
        }
        if (pp.GetFirstChild<W.Indentation>() is { } ind)
        {
            if (Length(ind.LeftChars?.Value, ind.Left?.Value ?? ind.Start?.Value) is { } left) props["indentLeft"] = left;
            if (Length(ind.RightChars?.Value, ind.Right?.Value ?? ind.End?.Value) is { } right) props["indentRight"] = right;
            if (Length(ind.HangingChars?.Value, ind.Hanging?.Value) is { } hanging) props["indentFirst"] = "-" + hanging;
            else if (Length(ind.FirstLineChars?.Value, ind.FirstLine?.Value) is { } first) props["indentFirst"] = first;
        }
        if (pp.GetFirstChild<W.ParagraphBorders>() is { } borders && Sides(borders) is { } sides) props["border"] = sides;
        if (DocxRun.On(pp.GetFirstChild<W.KeepNext>())) props["keepNext"] = "true";
        if (DocxRun.On(pp.GetFirstChild<W.KeepLines>())) props["keepLines"] = "true";
        if (pp.GetFirstChild<W.Tabs>() is { } tabs)
        {
            var stops = tabs.Elements<W.TabStop>().Where(t => t.Val?.Value is { } v && v != W.TabStopValues.Clear && v != W.TabStopValues.Number && t.Position?.Value is not null)
                .Select(t => $"{t.Val!.InnerText} {Cm(t.Position!.Value)}").ToList();
            if (stops.Count > 0) props["tabs"] = string.Join(", ", stops);
        }
    }

    public static void Write(W.ParagraphProperties pp, string name, string value)
    {
        var off = value is "none" or "";
        switch (name)
        {
            case "lineSpacing":
                var spacing = pp.SpacingBetweenLines ??= new W.SpacingBetweenLines();
                if (off) { spacing.Line = null; spacing.LineRule = null; }
                else if (value.StartsWith("min ", StringComparison.OrdinalIgnoreCase)) { spacing.Line = Twips(value[4..]).ToString(Inv); spacing.LineRule = W.LineSpacingRuleValues.AtLeast; }
                else if (double.TryParse(value, NumberStyles.Float, Inv, out var multiple)) { spacing.Line = ((int)Math.Round(multiple * 240)).ToString(Inv); spacing.LineRule = W.LineSpacingRuleValues.Auto; }
                else { spacing.Line = Twips(value).ToString(Inv); spacing.LineRule = W.LineSpacingRuleValues.Exact; }
                Trim(spacing);
                break;
            case "spaceBefore" or "spaceAfter":
                var sb = pp.SpacingBetweenLines ??= new W.SpacingBetweenLines();
                StringValue? twips = off ? null : Twips(value).ToString(Inv);
                if (name == "spaceBefore") { sb.Before = twips; sb.BeforeLines = null; sb.BeforeAutoSpacing = null; } else { sb.After = twips; sb.AfterLines = null; sb.AfterAutoSpacing = null; }
                Trim(sb);
                break;
            case "indentLeft" or "indentRight" or "indentFirst":
                var ind = pp.Indentation ??= new W.Indentation();
                var (chars, length) = off ? (null, null) : Measure(value.TrimStart('-'));
                var hanging = name == "indentFirst" && value.StartsWith('-');
                if (name == "indentLeft") { ind.LeftChars = chars; ind.Left = length; ind.Start = null; ind.StartCharacters = null; }
                else if (name == "indentRight") { ind.RightChars = chars; ind.Right = length; ind.End = null; ind.EndCharacters = null; }
                else { ind.FirstLineChars = hanging ? null : chars; ind.FirstLine = hanging ? null : length; ind.HangingChars = hanging ? chars : null; ind.Hanging = hanging ? length : null; }
                if (!ind.HasAttributes) ind.Remove();
                break;
            case "border":
                pp.ParagraphBorders?.Remove();
                if (off) break;
                var box = value.Trim().Equals("box", StringComparison.OrdinalIgnoreCase);
                var wanted = box ? new HashSet<string> { "top", "bottom", "left", "right" } : value.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries).Select(s => s.ToLowerInvariant()).ToHashSet();
                if (wanted.Except(["top", "bottom", "left", "right"]).Any())
                    throw new WriterException(ErrorCode.Validation, $"Unknown border side in '{value}'", "Use top, bottom, left, right (space-separated), box or none.");
                var pb = new W.ParagraphBorders();
                if (wanted.Contains("top")) pb.TopBorder = Line<W.TopBorder>();
                if (wanted.Contains("left")) pb.LeftBorder = Line<W.LeftBorder>();
                if (wanted.Contains("bottom")) pb.BottomBorder = Line<W.BottomBorder>();
                if (wanted.Contains("right")) pb.RightBorder = Line<W.RightBorder>();
                pp.ParagraphBorders = pb;
                break;
            case "keepNext": pp.KeepNext = value == "true" ? new W.KeepNext() : null; break;
            case "keepLines": pp.KeepLines = value == "true" ? new W.KeepLines() : null; break;
            case "tabs":
                pp.Tabs?.Remove();
                if (off) break;
                var list = new W.Tabs();
                foreach (var stop in value.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries))
                {
                    var parts = stop.Trim().Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);
                    var kind = parts.Length == 2 ? parts[0].ToLowerInvariant() : "left";
                    var val = kind switch
                    {
                        "left" => W.TabStopValues.Left, "center" => W.TabStopValues.Center, "right" => W.TabStopValues.Right, "decimal" => W.TabStopValues.Decimal, "bar" => W.TabStopValues.Bar,
                        _ => throw new WriterException(ErrorCode.Validation, $"Unknown tab stop '{stop.Trim()}'", "Write each stop as kind and position: left 2cm, center 8cm, right 15cm, decimal 10cm."),
                    };
                    list.Append(new W.TabStop { Val = val, Position = (int)Twips(parts[^1]) });
                }
                pp.Tabs = list;
                break;
        }
    }

    static T Line<T>() where T : W.BorderType, new() => new() { Val = W.BorderValues.Single, Size = 4U, Space = 1U, Color = "auto" };

    static void Trim(W.SpacingBetweenLines spacing) { if (!spacing.HasAttributes) spacing.Remove(); }

    static string? LineSpacingOf(W.SpacingBetweenLines sp)
    {
        if (sp.Line?.Value is not { } text || !int.TryParse(text, NumberStyles.Integer, Inv, out var line)) return null;
        var rule = sp.LineRule?.Value ?? W.LineSpacingRuleValues.Auto;
        if (rule == W.LineSpacingRuleValues.Exact) return Pt(line);
        if (rule == W.LineSpacingRuleValues.AtLeast) return "min " + Pt(line);
        return (line / 240.0).ToString("0.##", Inv);
    }

    static string? Sides(W.ParagraphBorders borders)
    {
        var sides = new List<string>();
        if (Drawn(borders.TopBorder)) sides.Add("top");
        if (Drawn(borders.BottomBorder)) sides.Add("bottom");
        if (Drawn(borders.LeftBorder)) sides.Add("left");
        if (Drawn(borders.RightBorder)) sides.Add("right");
        return sides.Count == 0 ? null : sides.Count == 4 ? "box" : string.Join(' ', sides);
    }

    static bool Drawn(W.BorderType? border) => border?.Val?.Value is { } v && v != W.BorderValues.Nil && v != W.BorderValues.None;

    static string Pt(string twips) => int.TryParse(twips, NumberStyles.Integer, Inv, out var n) ? Pt(n) : twips;
    static string Pt(long twips) => (twips / 20.0).ToString("0.##", Inv) + "pt";
    static string Cm(long twips) => (twips * EmuPerTwip / 360000.0).ToString("0.##", Inv) + "cm";

    /// <summary>A character count as Nch (hundredths in the file), else a twip length as cm.</summary>
    static string? Length(int? hundredthsOfChars, string? twips)
    {
        if (hundredthsOfChars is { } chars && chars != 0) return (chars / 100.0).ToString("0.##", Inv) + "ch";
        return twips is not null && long.TryParse(twips, NumberStyles.AllowLeadingSign, Inv, out var t) && t != 0 ? Cm(t) : null;
    }

    /// <summary>Points or a length as twips: 6pt, 6 (points), 0.5cm, 12px.</summary>
    static long Twips(string value)
    {
        var v = value.Trim();
        if (double.TryParse(v, NumberStyles.Float, Inv, out var points)) return (long)Math.Round(points * 20);
        return (long)Math.Round(Units.ParseLength(v) / (double)EmuPerTwip);
    }

    /// <summary>An indent as characters (2ch, 2em → 200 hundredths) or as a twip length.</summary>
    static (int? Chars, string? Twips) Measure(string value)
    {
        var v = value.Trim();
        if (v.EndsWith("ch", StringComparison.OrdinalIgnoreCase) || v.EndsWith("em", StringComparison.OrdinalIgnoreCase))
        {
            if (!double.TryParse(v[..^2], NumberStyles.Float, Inv, out var chars)) throw new WriterException(ErrorCode.Validation, $"Cannot read '{value}'", "Use characters like 2ch or a length like 0.74cm.");
            return ((int)Math.Round(chars * 100), null);
        }
        return (null, Twips(v).ToString(Inv));
    }
}
