using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
namespace Writer.Formats.Xlsx;
static class XlsxOutline
{
    internal static JsonArray Groups(XlsxSheet sheet)
    {
        var groups = new JsonArray();
        var rows = sheet.Data.Elements<Row>().Where(r => r.RowIndex?.Value is not null).ToDictionary(r => (int)r.RowIndex!.Value - 1);
        var columns = new Dictionary<int, Column>();
        foreach (var col in sheet.Ws.GetFirstChild<Columns>()?.Elements<Column>() ?? []) for (var i = (int)(col.Min?.Value ?? 1) - 1; i < Math.Min(16384, col.Max?.Value ?? 0); i++) columns[i] = col;
        void Collect(string axis, Dictionary<int, byte> levels, Func<int, bool> collapsed, Func<int, bool> hidden)
        {
            for (var level = 1; level <= 7; level++)
            {
                var selected = levels.Where(p => p.Value >= level).Select(p => p.Key).Order().ToArray(); var start = -1; var end = -1;
                void Add() { if (start >= 0) groups.Add((JsonNode)new JsonObject { ["axis"] = axis, ["start"] = start, ["end"] = end, ["level"] = level, ["collapsed"] = collapsed(end + 1) || end == (axis == "r" ? 1048575 : 16383) && Enumerable.Range(start, end - start + 1).All(hidden) }); }
                foreach (var n in selected) { if (n != end + 1) { Add(); start = n; } else if (start < 0) start = n; end = n; } Add();
            }
        }
        Collect("r", rows.ToDictionary(p => p.Key, p => p.Value.OutlineLevel?.Value ?? 0), i => rows.GetValueOrDefault(i)?.Collapsed?.Value == true, i => rows.GetValueOrDefault(i)?.Hidden?.Value == true);
        Collect("c", columns.ToDictionary(p => p.Key, p => p.Value.OutlineLevel?.Value ?? 0), i => columns.GetValueOrDefault(i)?.Collapsed?.Value == true, i => columns.GetValueOrDefault(i)?.Hidden?.Value == true);
        return groups;
    }
    internal static void Write(XlsxSheet sheet, string json)
    {
        var list = JsonNode.Parse(json) as JsonArray ?? throw Bad();
        foreach (var n in list) { if (n is null || n["axis"]?.GetValue<string>() is not ("r" or "c") || n["start"]?.GetValue<int>() is not { } a || n["end"]?.GetValue<int>() is not { } b || a < 0 || b < a || b >= (n["axis"]!.GetValue<string>() == "r" ? 1048576 : 16384) || (n["level"]?.GetValue<int>() ?? 1) is < 1 or > 7) throw Bad(); }
        var old = Groups(sheet); var cols = sheet.Ws.GetFirstChild<Columns>(); var rows = sheet.Data.Elements<Row>().ToDictionary(r => (int)(r.RowIndex?.Value ?? 0));
        foreach (var r in rows.Values) { r.OutlineLevel = null; r.Collapsed = null; }
        foreach (var c in cols?.Elements<Column>() ?? []) { c.OutlineLevel = null; c.Collapsed = null; }
        Row Row(int index) { if (!rows.TryGetValue(index + 1, out var row)) rows[index + 1] = row = XlsxCells.GetOrCreateRow(sheet.Data, index + 1); return row; }
        foreach (var n in old.Where(n => n!["collapsed"]!.GetValue<bool>())) for (var i = n!["start"]!.GetValue<int>(); i <= n["end"]!.GetValue<int>(); i++) { if (n["axis"]!.GetValue<string>() == "r") Row(i).Hidden = null; else XlsxLayout.ColumnAt(sheet.Ws, ref cols, i + 1).Hidden = null; }
        byte maxR = 0, maxC = 0;
        foreach (var n in list)
        {
            var axis = n!["axis"]!.GetValue<string>(); var a = n["start"]!.GetValue<int>(); var b = n["end"]!.GetValue<int>(); var level = (byte)(n["level"]?.GetValue<int>() ?? 1); var collapsed = n["collapsed"]?.GetValue<bool>() == true;
            for (var i = a; i <= b; i++) if (axis == "r") { var r = Row(i); r.OutlineLevel = Math.Max(r.OutlineLevel?.Value ?? 0, level); if (collapsed) r.Hidden = true; } else { var c = XlsxLayout.ColumnAt(sheet.Ws, ref cols, i + 1); c.OutlineLevel = Math.Max(c.OutlineLevel?.Value ?? 0, level); if (collapsed) c.Hidden = true; }
            if (axis == "r") { maxR = Math.Max(maxR, level); if (collapsed && b < 1048575) Row(b + 1).Collapsed = true; } else { maxC = Math.Max(maxC, level); if (collapsed && b < 16383) XlsxLayout.ColumnAt(sheet.Ws, ref cols, b + 2).Collapsed = true; }
        }
        var format = sheet.Ws.SheetFormatProperties ??= new SheetFormatProperties { DefaultRowHeight = 15 }; format.OutlineLevelRow = maxR; format.OutlineLevelColumn = maxC;
        var props = sheet.Ws.SheetProperties ??= new SheetProperties(); props.OutlineProperties ??= new OutlineProperties(); props.OutlineProperties.SummaryBelow = true; props.OutlineProperties.SummaryRight = true;
        sheet.Doc.RecalculateOnLoad();
    }
    static WriterException Bad() => new(ErrorCode.Validation, "Invalid outline group", "Use {axis:r/c,start,end,level:1..7,collapsed} with zero-based inclusive indices.");
}
