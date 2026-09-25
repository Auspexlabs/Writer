using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

/// <summary>Sheet-level rules, each read as JSON and written from JSON replacing the whole set, as merges are: conditional
/// formatting, data validation, the autofilter's criteria, hidden rows and columns, and the tab colour.</summary>
static class XlsxRules
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    const string X14Cf = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}", X14Dv = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";

    // ---- conditional formatting: [{range, type, operator, value, value2, text, rank, percent, bottom, above, stdDev, period, color, colors, iconSet, fill, bold…}] ----
    public static string? Cf(Worksheet ws, XlsxStyles styles)
    {
        var rules = ws.Elements<ConditionalFormatting>()
            .SelectMany(cf => cf.Elements<ConditionalFormattingRule>().Select(r => (Range: cf.SequenceOfReferences?.InnerText ?? "", Rule: r)))
            .OrderBy(x => x.Rule.Priority?.Value ?? int.MaxValue).ToList();
        if (rules.Count == 0) return null;
        return NodeJson.Compact(w =>
        {
            w.WriteStartArray();
            foreach (var (range, r) in rules)
            {
                w.WriteStartObject();
                w.WriteString("range", range);
                var type = r.Type?.InnerText ?? "expression";
                w.WriteString("type", type);
                if (r.Operator?.InnerText is { } op && type == "cellIs") w.WriteString("operator", op);
                var formulas = r.Elements<Formula>().Select(f => f.Text).ToList();
                if (r.Text?.Value is { } text) w.WriteString("text", text);
                else if (type is "cellIs" or "expression")
                {
                    if (formulas.Count > 0) w.WriteString("value", formulas[0]);
                    if (formulas.Count > 1) w.WriteString("value2", formulas[1]);
                }
                if (r.Rank?.Value is { } rank) w.WriteNumber("rank", rank);
                if (r.Percent?.Value == true) w.WriteBoolean("percent", true);
                if (r.Bottom?.Value == true) w.WriteBoolean("bottom", true);
                if (r.AboveAverage?.Value == false) w.WriteBoolean("above", false);
                if (r.StdDev?.Value is { } sd) w.WriteNumber("stdDev", sd);
                if (r.TimePeriod?.InnerText is { } period) w.WriteString("period", period);
                if (r.GetFirstChild<DataBar>() is { } bar && styles.Colors.Resolve(bar.GetFirstChild<Color>()) is { } barColor) w.WriteString("color", barColor);
                if (r.GetFirstChild<ColorScale>() is { } scale)
                {
                    w.WritePropertyName("colors");
                    w.WriteStartArray();
                    foreach (var c in scale.Elements<Color>()) if (styles.Colors.Resolve(c) is { } rgb) w.WriteStringValue(rgb);
                    w.WriteEndArray();
                }
                if (r.GetFirstChild<IconSet>() is { } icons)
                {
                    w.WriteString("iconSet", icons.IconSetValue?.InnerText ?? "3TrafficLights1");
                    if (icons.Reverse?.Value == true) w.WriteBoolean("reverse", true);
                }
                if (r.FormatId?.Value is { } dxf)
                    foreach (var (k, v) in styles.Dxf(dxf)) { if (v == "true") w.WriteBoolean(k, true); else w.WriteString(k, v); }
                w.WriteEndObject();
            }
            w.WriteEndArray();
        });
    }

    static readonly string[] CfTypes = ["cellIs", "containsText", "notContainsText", "beginsWith", "endsWith", "duplicateValues", "uniqueValues", "containsBlanks", "notContainsBlanks",
        "containsErrors", "notContainsErrors", "top10", "aboveAverage", "dataBar", "colorScale", "iconSet", "timePeriod", "expression"];
    static readonly string[] CfOps = ["equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual", "between", "notBetween"];

    static readonly Dictionary<string, string> Periods = new()
    {
        ["today"] = "FLOOR({0},1)=TODAY()", ["yesterday"] = "FLOOR({0},1)=TODAY()-1", ["tomorrow"] = "FLOOR({0},1)=TODAY()+1",
        ["last7Days"] = "AND(TODAY()-FLOOR({0},1)<=6,FLOOR({0},1)<=TODAY())",
        ["thisWeek"] = "AND(TODAY()-ROUNDDOWN({0},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({0},0)-TODAY()<=7-WEEKDAY(TODAY()))",
        ["lastWeek"] = "AND(TODAY()-ROUNDDOWN({0},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN({0},0)<(WEEKDAY(TODAY())+7))",
        ["nextWeek"] = "AND(ROUNDDOWN({0},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({0},0)-TODAY()<(15-WEEKDAY(TODAY())))",
        ["thisMonth"] = "AND(MONTH({0})=MONTH(TODAY()),YEAR({0})=YEAR(TODAY()))",
        ["lastMonth"] = "AND(MONTH({0})=MONTH(EDATE(TODAY(),0-1)),YEAR({0})=YEAR(EDATE(TODAY(),0-1)))",
        ["nextMonth"] = "AND(MONTH({0})=MONTH(EDATE(TODAY(),0+1)),YEAR({0})=YEAR(EDATE(TODAY(),0+1)))",
    };

    public static void SetCf(Worksheet ws, XlsxStyles styles, string json)
    {
        var built = new List<ConditionalFormatting>();
        // a rule whose look is one the sheet's rules already have keeps that dxf (with whatever else it holds: borders, a number
        // format), so rewriting the list leaves styles.xml as it was
        static string Look(IEnumerable<(string Name, string Value)> look) => string.Join(";", look.Select(p => p.Name + "=" + p.Value).Order(StringComparer.Ordinal));
        var looks = new Dictionary<string, uint>();
        foreach (var old in ws.Elements<ConditionalFormatting>().SelectMany(cf => cf.Elements<ConditionalFormattingRule>()))
            if (old.FormatId?.Value is { } id) looks.TryAdd(Look(styles.Dxf(id)), id);
        using (var parsed = JsonDocument.Parse(json))
        {
            if (parsed.RootElement.ValueKind != JsonValueKind.Array)
                throw new WriterException(ErrorCode.Validation, "cf must be a JSON array of rules", "Example: [{\"range\":\"A2:A11\",\"type\":\"cellIs\",\"operator\":\"greaterThan\",\"value\":\"80\",\"fill\":\"FFC7CE\"}]");
            var priority = 1;
            foreach (var e in parsed.RootElement.EnumerateArray())
            {
                var range = Str(e, "range") ?? throw new WriterException(ErrorCode.Validation, "cf: every rule needs a range", "Example: {\"range\":\"A2:A11\",\"type\":\"duplicateValues\"}");
                var parts = range.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 0) throw new WriterException(ErrorCode.Validation, "cf: every rule needs a range", "Example: {\"range\":\"A2:A11\",\"type\":\"duplicateValues\"}");
                var first = Range(parts[0], "cf");
                foreach (var p in parts.Skip(1)) Range(p, "cf");
                var cell = XlsxCells.Reference(first.Col1, first.Row1);
                var type = Str(e, "type") ?? "cellIs";
                if (!CfTypes.Contains(type))
                    throw new WriterException(ErrorCode.Validation, $"cf: '{type}' is not a rule type", "Types: " + string.Join(", ", CfTypes) + ".");
                var rule = new ConditionalFormattingRule { Type = new ConditionalFormatValues(type), Priority = priority++ };
                string value = Str(e, "value") ?? "", text = Str(e, "text") ?? "";
                var value2 = Str(e, "value2");
                switch (type)
                {
                    case "cellIs":
                        var op = Str(e, "operator") ?? "equal";
                        if (!CfOps.Contains(op)) throw new WriterException(ErrorCode.Validation, $"cf: '{op}' is not a cellIs operator", "Operators: " + string.Join(", ", CfOps) + ".");
                        rule.Operator = new ConditionalFormattingOperatorValues(op);
                        rule.Append(new Formula(value));
                        if (op is "between" or "notBetween") rule.Append(new Formula(value2 ?? value));
                        break;
                    case "containsText" or "notContainsText" or "beginsWith" or "endsWith":
                        var q = "\"" + text.Replace("\"", "\"\"") + "\"";
                        rule.Operator = new ConditionalFormattingOperatorValues(type);
                        rule.Text = text;
                        rule.Append(new Formula(type switch
                        {
                            "containsText" => $"NOT(ISERROR(SEARCH({q},{cell})))",
                            "notContainsText" => $"ISERROR(SEARCH({q},{cell}))",
                            "beginsWith" => $"LEFT({cell},{text.Length.ToString(Inv)})={q}",
                            _ => $"RIGHT({cell},{text.Length.ToString(Inv)})={q}",
                        }));
                        break;
                    case "top10":
                        rule.Rank = (uint)Math.Max(1, Num(e, "rank") ?? 10);
                        if (Bool(e, "percent")) rule.Percent = true;
                        if (Bool(e, "bottom")) rule.Bottom = true;
                        break;
                    case "aboveAverage":
                        if (e.TryGetProperty("above", out var ab) && ab.ValueKind == JsonValueKind.False) rule.AboveAverage = false;
                        if (Num(e, "stdDev") is { } sd) rule.StdDev = (int)sd;
                        break;
                    case "dataBar":
                        rule.Append(new DataBar(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min }, new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max },
                            styles.Colors.Like<Color>(Units.ParseColor(Str(e, "color") ?? "638EC6"))));
                        break;
                    case "colorScale":
                        var colors = e.TryGetProperty("colors", out var cs) && cs.ValueKind == JsonValueKind.Array ? cs.EnumerateArray().Select(c => Units.ParseColor(c.GetString() ?? "")).ToList() : ["F8696B", "FFEB84", "63BE7B"];
                        if (colors.Count is not (2 or 3)) throw new WriterException(ErrorCode.Validation, "cf: a colorScale takes 2 or 3 colors", "Example: \"colors\":[\"F8696B\",\"FFEB84\",\"63BE7B\"]");
                        var sc = new ColorScale();
                        sc.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                        if (colors.Count == 3) sc.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percentile, Val = "50" });
                        sc.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                        foreach (var c in colors) sc.Append(styles.Colors.Like<Color>(c));
                        rule.Append(sc);
                        break;
                    case "iconSet":
                        var name = Str(e, "iconSet") ?? "3TrafficLights1";
                        var n = name.Length > 0 && char.IsDigit(name[0]) ? name[0] - '0' : 3;
                        IconSet set;
                        try { set = new IconSet { IconSetValue = new IconSetValues(name) }; }
                        catch (ArgumentOutOfRangeException) { throw new WriterException(ErrorCode.Validation, $"cf: '{name}' is not an icon set", "Excel's names, e.g. 3TrafficLights1, 3Arrows, 4Rating, 5Quarters."); }
                        if (Bool(e, "reverse")) set.Reverse = true;
                        for (var i = 0; i < n; i++) set.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent, Val = (i * 100 / n).ToString(Inv) });
                        rule.Append(set);
                        break;
                    case "timePeriod":
                        var period = Str(e, "period") ?? "today";
                        if (!Periods.TryGetValue(period, out var f)) throw new WriterException(ErrorCode.Validation, $"cf: '{period}' is not a time period", "Periods: " + string.Join(", ", Periods.Keys));
                        rule.TimePeriod = new TimePeriodValues(period);
                        rule.Append(new Formula(string.Format(Inv, f, cell)));
                        break;
                    case "expression":
                        rule.Append(new Formula(value));
                        break;
                }
                var fill = Str(e, "fill");
                var color = Str(e, "color");
                if (type is "dataBar") color = null;
                if (fill is not null || color is not null || Bool(e, "bold") || Bool(e, "italic"))
                {
                    string? f = fill is null ? null : Units.ParseColor(fill), c = color is null ? null : Units.ParseColor(color);
                    var look = new List<(string, string)>();
                    if (f is not null) look.Add(("fill", f));
                    if (c is not null) look.Add(("color", c));
                    if (Bool(e, "bold")) look.Add(("bold", "true"));
                    if (Bool(e, "italic")) look.Add(("italic", "true"));
                    rule.FormatId = looks.TryGetValue(Look(look), out var same) ? same : styles.DxfId(f, c, Bool(e, "bold"), Bool(e, "italic"));
                }
                var cf = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(parts.Select(p => new StringValue(p))) };
                cf.Append(rule);
                built.Add(cf);
            }
        }
        foreach (var old in ws.Elements<ConditionalFormatting>().ToList()) old.Remove();
        RemoveExt(ws, X14Cf);
        foreach (var cf in built) InsertBeforeAny(ws, cf, e => e is DataValidations or Hyperlinks || Tail(e));
    }

    // ---- data validation: [{range, type, operator, value, value2, values, source, error, errorTitle, prompt, promptTitle, allowBlank}] ----
    public static string? Validations(Worksheet ws)
    {
        var list = ws.GetFirstChild<DataValidations>()?.Elements<DataValidation>().ToList();
        if (list is not { Count: > 0 }) return null;
        return NodeJson.Compact(w =>
        {
            w.WriteStartArray();
            foreach (var v in list)
            {
                w.WriteStartObject();
                w.WriteString("range", v.SequenceOfReferences?.InnerText ?? "");
                var type = v.Type?.InnerText ?? "any";
                w.WriteString("type", type);
                var f1 = v.Formula1?.Text?.Trim() ?? "";
                var f2 = v.Formula2?.Text?.Trim() ?? "";
                if (type == "list")
                {
                    if (f1.Length >= 2 && f1[0] == '"' && f1[^1] == '"')
                    {
                        w.WritePropertyName("values");
                        w.WriteStartArray();
                        foreach (var s in f1[1..^1].Replace("\"\"", "\"").Split(',')) w.WriteStringValue(s.Trim());
                        w.WriteEndArray();
                    }
                    else if (f1.Length > 0) w.WriteString("source", f1.Replace("$", ""));
                }
                else if (type != "any")
                {
                    if (type != "custom") w.WriteString("operator", v.Operator?.InnerText ?? "between");
                    if (f1.Length > 0) w.WriteString("value", f1);
                    if (f2.Length > 0) w.WriteString("value2", f2);
                }
                if (v.ErrorTitle?.Value is { Length: > 0 } et) w.WriteString("errorTitle", et);
                if (v.Error?.Value is { Length: > 0 } er) w.WriteString("error", er);
                if (v.PromptTitle?.Value is { Length: > 0 } pt) w.WriteString("promptTitle", pt);
                if (v.Prompt?.Value is { Length: > 0 } pr) w.WriteString("prompt", pr);
                if (v.AllowBlank?.Value == false) w.WriteBoolean("allowBlank", false);
                w.WriteEndObject();
            }
            w.WriteEndArray();
        });
    }

    public static void SetValidations(Worksheet ws, string json)
    {
        var built = new List<DataValidation>();
        using (var parsed = JsonDocument.Parse(json))
        {
            if (parsed.RootElement.ValueKind != JsonValueKind.Array)
                throw new WriterException(ErrorCode.Validation, "validations must be a JSON array of rules", "Example: [{\"range\":\"A2:A20\",\"type\":\"list\",\"values\":[\"Yes\",\"No\"]}]");
            foreach (var e in parsed.RootElement.EnumerateArray())
            {
                var range = Str(e, "range") ?? throw new WriterException(ErrorCode.Validation, "validations: every rule needs a range", "Example: {\"range\":\"A2:A20\",\"type\":\"whole\",\"operator\":\"between\",\"value\":\"1\",\"value2\":\"100\"}");
                var parts = range.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                foreach (var p in parts) Range(p, "validations");
                var type = Str(e, "type") ?? "any";
                if (type == "any") continue;
                if (type is not ("list" or "whole" or "decimal" or "date" or "time" or "textLength" or "custom"))
                    throw new WriterException(ErrorCode.Validation, $"validations: '{type}' is not a validation type", "Types: list, whole, decimal, date, time, textLength, custom.");
                var v = new DataValidation
                {
                    Type = new DataValidationValues(type), SequenceOfReferences = new ListValue<StringValue>(parts.Select(p => new StringValue(p))),
                    AllowBlank = !(e.TryGetProperty("allowBlank", out var ab) && ab.ValueKind == JsonValueKind.False), ShowInputMessage = true, ShowErrorMessage = true,
                };
                if (Str(e, "errorTitle") is { } et) v.ErrorTitle = et;
                if (Str(e, "error") is { } er) v.Error = er;
                if (Str(e, "promptTitle") is { } pt) v.PromptTitle = pt;
                if (Str(e, "prompt") is { } pr) v.Prompt = pr;
                if (type == "list")
                {
                    if (e.TryGetProperty("values", out var vals) && vals.ValueKind == JsonValueKind.Array)
                        v.Formula1 = new Formula1("\"" + string.Join(",", vals.EnumerateArray().Select(x => (x.GetString() ?? "").Replace("\"", "\"\""))) + "\"");
                    else if (Str(e, "source") is { Length: > 0 } source) v.Formula1 = new Formula1(Absolute(source));
                    else throw new WriterException(ErrorCode.Validation, "validations: a list needs values or a source range", "Example: {\"range\":\"A2:A20\",\"type\":\"list\",\"values\":[\"Yes\",\"No\"]}");
                }
                else
                {
                    if (type != "custom" && Str(e, "operator") is { } op && op != "between")
                    {
                        if (!CfOps.Contains(op)) throw new WriterException(ErrorCode.Validation, $"validations: '{op}' is not an operator", "Operators: " + string.Join(", ", CfOps) + ".");
                        v.Operator = new DataValidationOperatorValues(op);
                    }
                    if (Str(e, "value") is { Length: > 0 } f1) v.Formula1 = new Formula1(f1);
                    if (Str(e, "value2") is { Length: > 0 } f2) v.Formula2 = new Formula2(f2);
                }
                built.Add(v);
            }
        }
        ws.GetFirstChild<DataValidations>()?.Remove();
        RemoveExt(ws, X14Dv);
        if (built.Count == 0) return;
        var dvs = new DataValidations { Count = (uint)built.Count };
        foreach (var v in built) dvs.Append(v);
        InsertBeforeAny(ws, dvs, e => e is Hyperlinks || Tail(e));
    }

    /// <summary>A range as a list source: $A$1:$A$5, so it means the same cells from every cell of the rule.</summary>
    static string Absolute(string source)
    {
        var bang = source.LastIndexOf('!');
        var range = Regex.Replace(source[(bang + 1)..].Replace("$", ""), "([A-Za-z]+)(\\d+)", "$$$1$$$2");
        return bang < 0 ? range : source[..(bang + 1)] + range;
    }

    // ---- autofilter criteria: {"A":{"values":[…]},"C":{"operator":"greaterThan","value":"100"}} ----
    static readonly string[] Compare = ["equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual"];

    public static string? Filters(Worksheet ws)
    {
        var af = ws.GetFirstChild<AutoFilter>();
        if (af?.Reference?.Value is not { } reference) return null;
        var cols = af.Elements<FilterColumn>().ToList();
        if (cols.Count == 0) return null;
        var box = XlsxCells.ParseRange(reference);
        return NodeJson.Compact(w =>
        {
            w.WriteStartObject();
            foreach (var fc in cols)
            {
                var col = XlsxCells.ColumnName(box.Col1 + (int)(fc.ColumnId?.Value ?? 0));
                if (fc.GetFirstChild<Filters>() is { } fs)
                {
                    w.WritePropertyName(col);
                    w.WriteStartObject();
                    w.WritePropertyName("values");
                    w.WriteStartArray();
                    foreach (var f in fs.Elements<Filter>()) w.WriteStringValue(f.Val?.Value ?? "");
                    if (fs.Blank?.Value == true) w.WriteStringValue("");
                    w.WriteEndArray();
                    w.WriteEndObject();
                }
                else if (fc.GetFirstChild<CustomFilters>() is { } cfs)
                {
                    var items = cfs.Elements<CustomFilter>().Select(c => (Op: c.Operator?.InnerText ?? "equal", Val: c.Val?.Value ?? "")).ToList();
                    if (items.Count == 0) continue;
                    w.WritePropertyName(col);
                    w.WriteStartObject();
                    if (items.Count == 2 && cfs.And?.Value == true && items[0].Op == "greaterThanOrEqual" && items[1].Op == "lessThanOrEqual")
                    {
                        w.WriteString("operator", "between");
                        w.WriteString("value", items[0].Val);
                        w.WriteString("value2", items[1].Val);
                    }
                    else
                    {
                        var (op, val) = Unwild(items[0].Op, items[0].Val);
                        w.WriteString("operator", op);
                        w.WriteString("value", val);
                    }
                    w.WriteEndObject();
                }
            }
            w.WriteEndObject();
        });
    }

    /// <summary>Excel spells text criteria as wildcards on equal / notEqual: *x* contains, x* begins with, *x ends with.</summary>
    static (string Op, string Val) Unwild(string op, string val)
    {
        if (op is "equal" or "notEqual" && val.Length > 1)
        {
            if (val.Length > 2 && val.StartsWith('*') && val.EndsWith('*')) return (op == "equal" ? "contains" : "notContains", val[1..^1]);
            if (op == "equal" && val.EndsWith('*')) return ("beginsWith", val[..^1]);
            if (op == "equal" && val.StartsWith('*')) return ("endsWith", val[1..]);
        }
        return (op, val);
    }

    public static void SetFilters(Worksheet ws, string json)
    {
        var columns = new List<FilterColumn>();
        var af = ws.GetFirstChild<AutoFilter>();
        using (var parsed = JsonDocument.Parse(json))
        {
            if (parsed.RootElement.ValueKind != JsonValueKind.Object)
                throw new WriterException(ErrorCode.Validation, "filters must be a JSON object keyed by column letter", "Example: {\"A\":{\"values\":[\"East\"]},\"C\":{\"operator\":\"greaterThan\",\"value\":\"100\"}}");
            var entries = parsed.RootElement.EnumerateObject().ToList();
            if (entries.Count > 0 && af?.Reference?.Value is null)
                throw new WriterException(ErrorCode.Validation, "filters need an autofilter range", "Set filter=A1:D20 first, or in the same command before filters.");
            var box = af?.Reference?.Value is { } reference ? XlsxCells.ParseRange(reference) : default;
            foreach (var p in entries)
            {
                var key = p.Name.Trim();
                if (key.Length is 0 or > 3 || !key.All(char.IsAsciiLetter))
                    throw new WriterException(ErrorCode.Validation, $"filters: '{p.Name}' is not a column letter", "Keys are column letters like A or AB.");
                var col = XlsxCells.ColumnIndex(key);
                if (col < box.Col1 || col > box.Col2)
                    throw new WriterException(ErrorCode.Validation, $"filters: column {key.ToUpperInvariant()} is outside the autofilter range {af!.Reference!.Value}", "Filter a column inside the range, or widen filter=.");
                var fc = new FilterColumn { ColumnId = (uint)(col - box.Col1) };
                var e = p.Value;
                if (e.ValueKind != JsonValueKind.Object) throw new WriterException(ErrorCode.Validation, $"filters: '{p.Name}' needs an object", "Example: {\"values\":[\"East\",\"West\"]} or {\"operator\":\"greaterThan\",\"value\":\"100\"}");
                if (e.TryGetProperty("values", out var vals) && vals.ValueKind == JsonValueKind.Array)
                {
                    var fs = new Filters();
                    foreach (var v in vals.EnumerateArray())
                    {
                        var s = v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : v.GetRawText();
                        if (s.Length == 0) fs.Blank = true;
                        else fs.Append(new Filter { Val = s });
                    }
                    fc.Append(fs);
                }
                else
                {
                    var op = Str(e, "operator") ?? "equal";
                    var value = Str(e, "value") ?? "";
                    var cfs = new CustomFilters();
                    switch (op)
                    {
                        case "between":
                            cfs.And = true;
                            cfs.Append(new CustomFilter { Operator = FilterOperatorValues.GreaterThanOrEqual, Val = value });
                            cfs.Append(new CustomFilter { Operator = FilterOperatorValues.LessThanOrEqual, Val = Str(e, "value2") ?? value });
                            break;
                        case "contains": cfs.Append(new CustomFilter { Operator = FilterOperatorValues.Equal, Val = "*" + value + "*" }); break;
                        case "notContains": cfs.Append(new CustomFilter { Operator = FilterOperatorValues.NotEqual, Val = "*" + value + "*" }); break;
                        case "beginsWith": cfs.Append(new CustomFilter { Operator = FilterOperatorValues.Equal, Val = value + "*" }); break;
                        case "endsWith": cfs.Append(new CustomFilter { Operator = FilterOperatorValues.Equal, Val = "*" + value }); break;
                        default:
                            if (!Compare.Contains(op)) throw new WriterException(ErrorCode.Validation, $"filters: '{op}' is not an operator", "Operators: " + string.Join(", ", Compare) + ", between, contains, notContains, beginsWith, endsWith.");
                            cfs.Append(new CustomFilter { Operator = new FilterOperatorValues(op), Val = value });
                            break;
                    }
                    fc.Append(cfs);
                }
                columns.Add(fc);
            }
        }
        if (af is null) return;
        foreach (var old in af.Elements<FilterColumn>().ToList()) old.Remove();
        foreach (var fc in columns.OrderBy(c => c.ColumnId!.Value)) af.Append(fc);
    }

    // ---- hidden rows and columns: {"rows":[3,4],"cols":["B"]} ----
    public static string? Hidden(Worksheet ws)
    {
        var rows = ws.GetFirstChild<SheetData>()?.Elements<Row>().Where(r => r.Hidden?.Value == true && r.RowIndex?.Value is not null).Select(r => (int)r.RowIndex!.Value).ToList() ?? [];
        var cols = new List<int>();
        foreach (var c in ws.GetFirstChild<Columns>()?.Elements<Column>() ?? [])
            if (c.Hidden?.Value == true && c.Min?.Value is { } min && c.Max?.Value is { } max)
                for (var i = min; i <= max && i <= 16384; i++) cols.Add((int)i);
        if (rows.Count == 0 && cols.Count == 0) return null;
        return NodeJson.Compact(w =>
        {
            w.WriteStartObject();
            w.WritePropertyName("rows");
            w.WriteStartArray();
            foreach (var r in rows) w.WriteNumberValue(r);
            w.WriteEndArray();
            w.WritePropertyName("cols");
            w.WriteStartArray();
            foreach (var c in cols) w.WriteStringValue(XlsxCells.ColumnName(c));
            w.WriteEndArray();
            w.WriteEndObject();
        });
    }

    public static void SetHidden(Worksheet ws, SheetData data, string json)
    {
        HashSet<int> rows = [], cols = [];
        using (var parsed = JsonDocument.Parse(json))
        {
            if (parsed.RootElement.ValueKind != JsonValueKind.Object)
                throw new WriterException(ErrorCode.Validation, "hidden must be a JSON object", "Example: {\"rows\":[3,4],\"cols\":[\"B\"]}");
            if (parsed.RootElement.TryGetProperty("rows", out var rs) && rs.ValueKind == JsonValueKind.Array)
                foreach (var r in rs.EnumerateArray())
                    rows.Add(r.ValueKind == JsonValueKind.Number && r.TryGetInt32(out var i) && i is >= 1 and <= 1048576 ? i
                        : throw new WriterException(ErrorCode.Validation, $"hidden: '{r}' is not a row number", "Rows are numbers like 3."));
            if (parsed.RootElement.TryGetProperty("cols", out var cs) && cs.ValueKind == JsonValueKind.Array)
                foreach (var c in cs.EnumerateArray())
                {
                    var k = (c.GetString() ?? "").Trim();
                    if (k.Length is 0 or > 3 || !k.All(char.IsAsciiLetter)) throw new WriterException(ErrorCode.Validation, $"hidden: '{c}' is not a column letter", "Columns are letters like B.");
                    cols.Add(XlsxCells.ColumnIndex(k));
                }
        }
        foreach (var r in data.Elements<Row>().ToList())
        {
            var i = (int)(r.RowIndex?.Value ?? 0);
            if (rows.Remove(i)) r.Hidden = true;
            else if (r.Hidden is not null)
            {
                r.Hidden = null;
                if (!r.HasChildren && r.GetAttributes().All(a => a.LocalName == "r")) r.Remove();
            }
        }
        foreach (var i in rows) XlsxCells.GetOrCreateRow(data, i).Hidden = true;
        var columns = ws.GetFirstChild<Columns>();
        var shown = new List<int>();
        foreach (var c in columns?.Elements<Column>().Where(c => c.Hidden?.Value == true).ToList() ?? [])
            for (var i = c.Min?.Value ?? 0; i <= (c.Max?.Value ?? 0) && i <= 16384; i++) if (!cols.Remove((int)i)) shown.Add((int)i);
        foreach (var i in shown)
        {
            var c = XlsxLayout.Isolate(columns!, i)!;
            c.Hidden = null;
            if (c.GetAttributes().All(a => a.LocalName is "min" or "max")) c.Remove();
        }
        foreach (var i in cols) XlsxLayout.ColumnAt(ws, ref columns, i).Hidden = true;
        if (columns is not null && !columns.HasChildren) columns.Remove();
    }

    // ---- tab colour ----
    public static string? TabColor(Worksheet ws, XlsxStyles styles) => styles.Colors.Resolve(ws.SheetProperties?.TabColor);

    public static void SetTabColor(Worksheet ws, XlsxStyles styles, string value)
    {
        var pr = ws.SheetProperties;
        if (value == "none")
        {
            if (pr is null) return;
            pr.TabColor = null;
            if (!pr.HasChildren && !pr.HasAttributes) pr.Remove();
            return;
        }
        pr ??= ws.SheetProperties = new SheetProperties();
        pr.TabColor = styles.Colors.Like<TabColor>(value);
    }

    // ---- helpers ----
    static string? Str(JsonElement e, string name) => e.ValueKind == JsonValueKind.Object && e.TryGetProperty(name, out var v) && v.ValueKind != JsonValueKind.Null
        ? (v.ValueKind == JsonValueKind.String ? v.GetString() : v.GetRawText()) : null;

    static double? Num(JsonElement e, string name) => e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetDouble()
        : v.ValueKind == JsonValueKind.String && double.TryParse(v.GetString(), NumberStyles.Float, Inv, out var d) ? d : null;

    static bool Bool(JsonElement e, string name) => e.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || (v.ValueKind == JsonValueKind.String && v.GetString() == "true"));

    static (int Col1, int Row1, int Col2, int Row2) Range(string reference, string prop)
    {
        try { return XlsxCells.ParseRange(reference); }
        catch (WriterException) { throw new WriterException(ErrorCode.Validation, $"{prop}: '{reference}' is not a cell range", "Use a range like A2:A11."); }
    }

    /// <summary>The worksheet children that follow hyperlinks in the schema.</summary>
    static bool Tail(OpenXmlElement e) => e is PrintOptions or PageMargins or PageSetup or HeaderFooter or RowBreaks or ColumnBreaks or CustomProperties or CellWatches
        or IgnoredErrors or Drawing or LegacyDrawing or LegacyDrawingHeaderFooter or DrawingHeaderFooter or Picture or OleObjects or Controls or WebPublishItems
        or TableParts or WorksheetExtensionList or AlternateContent;

    static void InsertBeforeAny(OpenXmlCompositeElement parent, OpenXmlElement element, Func<OpenXmlElement, bool> after)
    {
        var anchor = parent.ChildElements.FirstOrDefault(after);
        if (anchor is null) parent.Append(element);
        else parent.InsertBefore(element, anchor);
    }

    /// <summary>Drops an x14 extension the rewritten rules would leave orphaned (Excel repairs a file whose ext names rules that are gone).</summary>
    static void RemoveExt(Worksheet ws, string uri)
    {
        var list = ws.GetFirstChild<WorksheetExtensionList>();
        if (list is null) return;
        foreach (var ext in list.Elements<WorksheetExtension>().Where(x => string.Equals(x.Uri?.Value, uri, StringComparison.OrdinalIgnoreCase)).ToList()) ext.Remove();
        if (!list.HasChildren) list.Remove();
    }
}
