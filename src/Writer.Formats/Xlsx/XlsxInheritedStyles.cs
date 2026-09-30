using System.Globalization;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

/// <summary>Sparse row/column defaults. The style catalog carries values, not live
/// stylesheet indices, so undo after save remains valid when unused styles are trimmed.</summary>
static class XlsxInheritedStyles
{
    internal static string Read(XlsxDocument doc, XlsxSheet sheet)
    {
        var rows = new JsonObject(); var cols = new JsonObject(); var styles = new JsonObject();
        string Add(uint index)
        {
            var id = index.ToString(CultureInfo.InvariantCulture);
            if (!styles.ContainsKey(id)) styles[id] = doc.Styles.Snapshot(index);
            return id;
        }
        foreach (var row in sheet.Data.Elements<Row>())
            if (row.RowIndex is { } r && row.StyleIndex is { } s && row.CustomFormat?.Value != false) rows[r.Value.ToString(CultureInfo.InvariantCulture)] = Add(s.Value);
        foreach (var col in sheet.Ws.GetFirstChild<Columns>()?.Elements<Column>() ?? [])
            if (col.Style is { } s && col.Min is { } min && col.Max is { } max)
                for (var i = min.Value; i <= Math.Min(max.Value, 16384); i++) cols[XlsxCells.ColumnName((int)i)] = Add(s.Value);
        return new JsonObject { ["rows"] = rows, ["cols"] = cols, ["styles"] = styles, ["base"] = doc.Styles.Snapshot(0) }.ToJsonString();
    }

    internal static void Write(XlsxDocument doc, XlsxSheet sheet, string json)
    {
        var input = JsonNode.Parse(json) as JsonObject;
        if (input?["rows"] is not JsonObject rows || input["cols"] is not JsonObject cols || input["styles"] is not JsonObject styles)
            throw Invalid();
        var rowIds = new Dictionary<int,string>(); var colIds = new Dictionary<int,string>();
        foreach (var (key,value) in rows)
        {
            if (!int.TryParse(key, NumberStyles.None, CultureInfo.InvariantCulture, out var i) || i is < 1 or > 1048576 || value is not JsonValue v || !v.TryGetValue<string>(out var id) || styles[id] is not JsonObject) throw Invalid();
            rowIds.Add(i,id);
        }
        foreach (var (key,value) in cols)
        {
            if (key.Length is < 1 or > 3 || key.Any(c => c is < 'A' or > 'Z') || XlsxCells.ColumnIndex(key) > 16384 || value is not JsonValue v || !v.TryGetValue<string>(out var id) || styles[id] is not JsonObject) throw Invalid();
            colIds.Add(XlsxCells.ColumnIndex(key),id);
        }
        var used = rowIds.Values.Concat(colIds.Values).Distinct().ToArray();
        // Parse all definitions before changing the worksheet or stylesheet.
        foreach (var id in used) XlsxStyles.ValidateSnapshot(styles[id]!.AsObject());
        var mapped = used.ToDictionary(id => id, id => doc.Styles.Restore(styles[id]!.AsObject()));
        foreach (var row in sheet.Data.Elements<Row>()) { row.StyleIndex = null; row.CustomFormat = null; }
        foreach (var (index,id) in rowIds) { var row = XlsxCells.GetOrCreateRow(sheet.Data,index); row.StyleIndex = mapped[id]; row.CustomFormat = true; }
        var columns = sheet.Ws.GetFirstChild<Columns>();
        foreach (var col in columns?.Elements<Column>() ?? []) col.Style = null;
        foreach (var (index,id) in colIds.OrderBy(x => x.Key)) XlsxLayout.ColumnAt(sheet.Ws,ref columns,index).Style = mapped[id];
    }
    static WriterException Invalid() => new(ErrorCode.Validation,"Invalid row/column style defaults","Use sparse rows, cols and a style catalog.");
}
