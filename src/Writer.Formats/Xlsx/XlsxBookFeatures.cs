using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

static class XlsxBookFeatures
{
    internal static string Names(XlsxDocument doc, uint? localSheet = null)
    {
        var names = new JsonObject();
        foreach (var n in doc.Workbook.Workbook!.DefinedNames?.Elements<DefinedName>() ?? []) if (n.LocalSheetId?.Value == localSheet && n.Name?.Value is { } name && !name.StartsWith("_xlnm.")) names[name] = n.Text;
        return names.ToJsonString();
    }
    internal static void SetNames(XlsxDocument doc, string json, uint? localSheet = null)
    {
        var values = JsonNode.Parse(json) as JsonObject ?? throw new WriterException(ErrorCode.Validation, "names must be an object", "Use name to formula pairs.");
        if (values.Select(p => p.Key).Distinct(StringComparer.OrdinalIgnoreCase).Count() != values.Count) throw new WriterException(ErrorCode.Validation, "Duplicate workbook names", "Names are case-insensitive.");
        foreach (var (name, node) in values)
            if (!Regex.IsMatch(name, @"^[\p{L}_\\][\p{L}\p{N}_.\\]*$") || Regex.IsMatch(name, @"^(?:[A-Za-z]{1,3}[1-9]\d*|[RrCc])$") || node is not JsonValue)
                throw new WriterException(ErrorCode.Validation, "Invalid defined name: " + name, "Use a name such as Sales_Total and a range or formula as its value.");
        var book = doc.Workbook.Workbook!;
        var list = book.DefinedNames ?? new DefinedNames();
        if (list.Parent is null) book.AddChild(list, true);
        foreach (var old in list.Elements<DefinedName>().Where(n => n.LocalSheetId?.Value == localSheet && !(n.Name?.Value ?? "").StartsWith("_xlnm.")).ToList()) old.Remove();
        foreach (var (name, node) in values)
        {
            var defined = new DefinedName(node!.GetValue<string>().TrimStart('=')) { Name = name };
            if (localSheet is { } index) defined.LocalSheetId = index;
            list.Append(defined);
        }
        if (!list.HasChildren) list.Remove();
        doc.RecalculateOnLoad();
    }
    static string? LocalName(XlsxDocument doc, int index, string name) => doc.Workbook.Workbook!.DefinedNames?.Elements<DefinedName>().FirstOrDefault(n => n.LocalSheetId?.Value == index && n.Name?.Value == name)?.Text;
    static void SetLocalName(XlsxDocument doc, int index, string name, string value)
    {
        var book = doc.Workbook.Workbook!; var list = book.DefinedNames ?? new DefinedNames();
        if (list.Parent is null) book.AddChild(list, true);
        var old = list.Elements<DefinedName>().FirstOrDefault(n => n.LocalSheetId?.Value == index && n.Name?.Value == name);
        if (string.IsNullOrWhiteSpace(value)) { old?.Remove(); if (!list.HasChildren) list.Remove(); return; }
        if (old is null) list.Append(new DefinedName(value) { Name = name, LocalSheetId = (uint)index }); else old.Text = value;
    }
    internal static string Print(XlsxDocument doc, XlsxSheet sheet)
    {
        var ws = sheet.Ws; var setup = ws.GetFirstChild<PageSetup>(); var margins = ws.GetFirstChild<PageMargins>(); var hf = ws.GetFirstChild<HeaderFooter>();
        var index = doc.Sheets.FindIndex(s => ReferenceEquals(s.Part, sheet.Part));
        return new JsonObject {
            ["orientation"] = setup?.Orientation?.InnerText ?? "portrait", ["paper"] = setup?.PaperSize?.Value ?? 9U,
            ["scale"] = setup?.Scale?.Value ?? 100U, ["fitWidth"] = setup?.FitToWidth?.Value, ["fitHeight"] = setup?.FitToHeight?.Value,
            ["left"] = margins?.Left?.Value ?? 0.7, ["right"] = margins?.Right?.Value ?? 0.7, ["top"] = margins?.Top?.Value ?? 0.75, ["bottom"] = margins?.Bottom?.Value ?? 0.75,
            ["header"] = hf?.OddHeader?.Text ?? "", ["footer"] = hf?.OddFooter?.Text ?? "", ["area"] = LocalName(doc, index, "_xlnm.Print_Area") ?? "", ["titles"] = LocalName(doc, index, "_xlnm.Print_Titles") ?? "",
            ["gridlines"] = ws.GetFirstChild<PrintOptions>()?.GridLines?.Value ?? false,
            ["rowBreaks"] = new JsonArray((ws.GetFirstChild<RowBreaks>()?.Elements<Break>() ?? []).Where(b => b.ManualPageBreak?.Value == true).Select(b => JsonValue.Create(b.Id?.Value)).Cast<JsonNode?>().ToArray()),
            ["colBreaks"] = new JsonArray((ws.GetFirstChild<ColumnBreaks>()?.Elements<Break>() ?? []).Where(b => b.ManualPageBreak?.Value == true).Select(b => JsonValue.Create(b.Id?.Value)).Cast<JsonNode?>().ToArray())
        }.ToJsonString();
    }
    internal static void SetPrint(XlsxDocument doc, XlsxSheet sheet, string json)
    {
        var o = JsonNode.Parse(json) as JsonObject ?? throw new WriterException(ErrorCode.Validation, "print must be an object", "Use paper, orientation, margins, area, titles, header and footer.");
        var ws = sheet.Ws; var setup = ws.GetFirstChild<PageSetup>() ?? new PageSetup(); if (setup.Parent is null) ws.AddChild(setup, true);
        if (o["orientation"] is { } orient) setup.Orientation = orient.GetValue<string>() == "landscape" ? OrientationValues.Landscape : OrientationValues.Portrait;
        if (o["paper"] is { } paper) setup.PaperSize = paper.GetValue<uint>();
        if (o["scale"] is { } scale) setup.Scale = Math.Clamp(scale.GetValue<uint>(), 10, 400);
        if (o.ContainsKey("fitWidth")) setup.FitToWidth = o["fitWidth"]?.GetValue<uint>();
        if (o.ContainsKey("fitHeight")) setup.FitToHeight = o["fitHeight"]?.GetValue<uint>();
        var properties = ws.SheetProperties ??= new SheetProperties();
        properties.PageSetupProperties ??= new PageSetupProperties(); properties.PageSetupProperties.FitToPage = setup.FitToWidth is not null || setup.FitToHeight is not null;
        var margins = ws.GetFirstChild<PageMargins>() ?? new PageMargins { Left = 0.7, Right = 0.7, Top = 0.75, Bottom = 0.75, Header = 0.3, Footer = 0.3 };
        if (margins.Parent is null) ws.AddChild(margins, true);
        if (o["left"] is { } left) margins.Left = Math.Max(0, left.GetValue<double>()); if (o["right"] is { } right) margins.Right = Math.Max(0, right.GetValue<double>());
        if (o["top"] is { } top) margins.Top = Math.Max(0, top.GetValue<double>()); if (o["bottom"] is { } bottom) margins.Bottom = Math.Max(0, bottom.GetValue<double>());
        var hf = ws.GetFirstChild<HeaderFooter>() ?? new HeaderFooter(); if (hf.Parent is null) ws.AddChild(hf, true);
        if (o["header"] is { } header) hf.OddHeader = new OddHeader(header.GetValue<string>()); if (o["footer"] is { } footer) hf.OddFooter = new OddFooter(footer.GetValue<string>());
        if (o["gridlines"] is { } grid) { var po = ws.GetFirstChild<PrintOptions>() ?? new PrintOptions(); if (po.Parent is null) ws.AddChild(po, true); po.GridLines = grid.GetValue<bool>(); }
        var index = doc.Sheets.FindIndex(s => ReferenceEquals(s.Part, sheet.Part));
        if (o.ContainsKey("area")) SetLocalName(doc, index, "_xlnm.Print_Area", o["area"]?.GetValue<string>() ?? "");
        if (o.ContainsKey("titles")) SetLocalName(doc, index, "_xlnm.Print_Titles", o["titles"]?.GetValue<string>() ?? "");
        SetBreaks<RowBreaks>(ws, o, "rowBreaks", 1048576, 16383);
        SetBreaks<ColumnBreaks>(ws, o, "colBreaks", 16384, 1048575);
    }
    static void SetBreaks<T>(Worksheet ws, JsonObject o, string key, uint limit, uint max) where T : OpenXmlCompositeElement, new()
    {
        if (!o.ContainsKey(key)) return;
        var values = o[key]?.AsArray().Select(x => x!.GetValue<uint>()).Distinct().Order().ToArray() ?? [];
        if (values.Any(x => x == 0 || x >= limit)) throw new WriterException(ErrorCode.Validation, "Invalid page break", "Place page breaks inside the worksheet.");
        ws.RemoveAllChildren<T>();
        if (values.Length == 0) return;
        var breaks = new T();
        breaks.SetAttribute(new OpenXmlAttribute("count", "", values.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        breaks.SetAttribute(new OpenXmlAttribute("manualBreakCount", "", values.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        foreach (var id in values) breaks.Append(new Break { Id = id, Min = 0, Max = max, ManualPageBreak = true });
        ws.AddChild(breaks, true);
    }
    internal static void SetProtection(Worksheet ws, string value)
    {
        if (value != "true") { ws.RemoveAllChildren<SheetProtection>(); return; }
        var p = ws.GetFirstChild<SheetProtection>() ?? new SheetProtection { Objects = true, Scenarios = true }; if (p.Parent is null) ws.AddChild(p, true); p.Sheet = true;
    }
}
