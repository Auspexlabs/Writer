using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Spreadsheet;
using Jint;
using Jint.Runtime.Interop;
using Writer.Core;

namespace Writer.Formats.Xlsx;

/// <summary>The same formula implementation is embedded in the CLI, browser engine and desktop engine.
/// Formula text is parsed by Calc, never evaluated as JavaScript. Only the checked-in implementation runs in Jint.</summary>
static class XlsxCalculation
{
    static readonly Lazy<Jint.Prepared<Acornima.Ast.Script>> Program = new(() => Engine.PrepareScript(Source()));
    static string Source()
    {
        var assembly = typeof(XlsxCalculation).Assembly;
        var source = new System.Text.StringBuilder();
        foreach (var name in new[] { "sheet-model", "sheet-statistics", "sheet-whatif", "sheet-pivot-values", "sheet-pivot-fields", "sheet-engine", "word-table-formula" })
        {
            using var reader = new StreamReader(assembly.GetManifestResourceStream("Writer.Calculation." + name + ".js")!);
            foreach (var line in reader.ReadToEnd().Split('\n'))
            {
                if (line.StartsWith("import ") || line.StartsWith("export {")) continue;
                source.AppendLine(line.StartsWith("export ") ? line[7..] : line);
            }
        }
        source.AppendLine("""
            function calculateWordTableBatch(json) { return JSON.stringify(calculateWordTables(JSON.parse(json).map(input=>Array.isArray(input)?{rows:input}:input))); }
            function calculatePivotFormulaBatch(json) { return JSON.stringify(pivotFormulaBatch(JSON.parse(json))); }
            function convertTableFormulaBatch(json) {
              const input=JSON.parse(json),calc=new Calc(input.doc),table=input.doc.sheets[input.sheet].tables.find(t=>t.name===input.table);
              return JSON.stringify(input.formulas.map(([si,r,c,formula])=>convertTableRefs(formula,table,calc,si,r,c)));
            }
            function calculateWorkbook(json) {
              const doc = JSON.parse(json);
              doc.sheets.forEach((sh, si) => {
                const cells = Object.create(null); let addresses;
                sh.cells = new Proxy(cells, { get(target, key) {
                  if (typeof key !== 'string') return undefined;
                  if (!Object.hasOwn(target, key)) target[key] = JSON.parse(loadCell(si, key));
                  return target[key];
                }, ownKeys() { return addresses || (addresses = JSON.parse(loadCellKeys(si))); },
                getOwnPropertyDescriptor(target, key) { return { enumerable: true, configurable: true }; }});
              });
              const calc = new Calc(doc), out = []; calc.affected = new Set(doc.dirty);
              doc.sheets.forEach((sh, si) => sh.formulas.forEach(a => { const p = parseA(a), value = calc.value(si, p.r, p.c), matrix = calc.arrays.get(`${si}:${p.r}:${p.c}`); if (doc.all || calc.affected.has(`${si}:${p.r}:${p.c}`)) out.push([si, a, value, matrix ? A(p.r + matrix.length - 1, p.c + matrix[0].length - 1) : null]); }));
              for (const [key, spill] of calc.spills) { const [si, r, c] = key.split(':').map(Number); if (doc.all || calc.affected.has(spill.owner)) out.push([si, A(r, c), spill.value, null]); }
              doc.sheets.forEach((sh,si)=>(sh.dataTables||[]).forEach(t=>{const b=whatIfBox(t);for(let r=b.r1+1;r<=b.r2;r++)for(let c=b.c1+1;c<=b.c2;c++)out.push([si,A(r,c),calc.value(si,r,c),null]);}));
              return JSON.stringify(out);
            }
            """);
        return source.ToString();
    }

    internal static JsonDocument CalculateWordTables(string json)
    {
        var engine = new Engine(o => o.TimeoutInterval(TimeSpan.FromSeconds(15)).LimitRecursion(256).LimitMemory(128_000_000));
        var program = Program.Value; engine.Execute(in program);
        return JsonDocument.Parse(engine.Call(engine.GetValue("calculateWordTableBatch"), Jint.Native.JsValue.Undefined, [new Jint.Native.JsString(json)]).AsString());
    }

    internal static JsonArray CalculatePivotFormulas(JsonObject input)
    {
        try {
            var engine = new Engine(o => o.TimeoutInterval(TimeSpan.FromSeconds(15)).LimitRecursion(256).LimitMemory(128_000_000));
            var program = Program.Value; engine.Execute(in program);
            return JsonNode.Parse(engine.Call(engine.GetValue("calculatePivotFormulaBatch"), Jint.Native.JsValue.Undefined, [new Jint.Native.JsString(input.ToJsonString())]).AsString())!.AsArray();
        } catch (Jint.Runtime.JavaScriptException e) { throw new WriterException(ErrorCode.Validation, e.Message, "Use source field names and arithmetic or supported scalar functions."); }
    }

    internal static void ConvertTableReferences(XlsxDocument doc, int sheet, string table)
    {
        var sheets = doc.Sheets.Select(s => new XlsxSheet(doc, s.Sheet, s.Part)).ToArray();
        var formulas = sheets.SelectMany((s, si) => s.Data.Descendants<Cell>().Where(c => c.CellFormula is not null).Select(c => (Cell: c, Sheet: si))).ToArray();
        if (formulas.Length == 0) return;
        var input = new JsonObject { ["sheet"] = sheet, ["table"] = table,
            ["doc"] = new JsonObject { ["sheets"] = new JsonArray(sheets.Select(s => (JsonNode)new JsonObject { ["name"] = s.SheetName, ["tables"] = JsonNode.Parse(XlsxTables.Read(s)) }).ToArray()) },
            ["formulas"] = new JsonArray(formulas.Select(f => { var (c, r) = XlsxCells.Position(f.Cell); return (JsonNode)new JsonArray(JsonValue.Create(f.Sheet), JsonValue.Create(r - 1), JsonValue.Create(c - 1), JsonValue.Create(f.Cell.CellFormula!.Text)); }).ToArray()) };
        var engine = new Engine(o => o.TimeoutInterval(TimeSpan.FromSeconds(30)).LimitRecursion(512).LimitMemory(256_000_000)); var program = Program.Value; engine.Execute(in program);
        var output = JsonNode.Parse(engine.Call(engine.GetValue("convertTableFormulaBatch"), Jint.Native.JsValue.Undefined, [new Jint.Native.JsString(input.ToJsonString())]).AsString())!.AsArray();
        for (var i = 0; i < formulas.Length; i++) formulas[i].Cell.CellFormula!.Text = output[i]!.GetValue<string>();
        doc.RecalculateOnLoad();
    }

    internal static void Recalculate(XlsxDocument doc)
    {
        var sheets = doc.Sheets.Select(s => new XlsxSheet(doc, s.Sheet, s.Part)).ToArray();
        var rows = sheets.Select(s => s.Data.Elements<Row>().ToDictionary(r => (int)(r.RowIndex?.Value ?? 0))).ToArray();
        var formulas = sheets.Select(s => s.Data.Elements<Row>().SelectMany(r => r.Elements<Cell>()).Where(c => c.CellFormula is not null).ToArray()).ToArray();
        if (formulas.All(f => f.Length == 0)) return;
        var plan=XlsxCalculationDependencies.Plan(doc,sheets,formulas);
        if(plan is {Length:0})return;
        var planned=plan?.ToHashSet();
        var nodes = new JsonArray();
        for (var i = 0; i < sheets.Length; i++)
        {
            var maxR = rows[i].Keys.DefaultIfEmpty(1).Max();
            var maxC = plan is not null?1:rows[i].Values.SelectMany(r => r.Elements<Cell>()).Select(c => XlsxCells.Position(c).Col).DefaultIfEmpty(1).Max();
            var hidden = JsonNode.Parse(XlsxRules.Hidden(sheets[i].Ws) ?? "{}")?["rows"] as JsonArray ?? [];
            var filter = sheets[i].Ws.GetFirstChild<AutoFilter>()?.Reference?.Value;
            var filterBox = filter is not null ? XlsxCells.ParseRange(filter) : (Col1: 0, Row1: 0, Col2: 0, Row2: 0);
            bool Filtered(JsonNode? n) => filter is not null && n!.GetValue<int>() > filterBox.Row1 && n.GetValue<int>() <= filterBox.Row2;
            nodes.Add((JsonNode)new JsonObject { ["name"] = sheets[i].SheetName,
                ["tables"] = JsonNode.Parse(XlsxTables.Read(sheets[i])),
                ["dataTables"] = XlsxDataTables.Read(sheets[i]),
                ["names"] = JsonNode.Parse(XlsxBookFeatures.Names(doc, (uint)i)),
                ["outline"] = XlsxOutline.Groups(sheets[i]),
                ["hiddenRows"] = new JsonArray(hidden.Where(n => !Filtered(n)).Select(n => (JsonNode?)JsonValue.Create(n!.GetValue<int>() - 1)).ToArray()),
                ["frows"] = new JsonArray(hidden.Where(Filtered).Select(n => (JsonNode?)JsonValue.Create(n!.GetValue<int>() - 1)).ToArray()),
                ["_used"] = new JsonObject { ["r1"] = 0, ["c1"] = 0, ["r2"] = maxR - 1, ["c2"] = maxC - 1 },
                ["formulas"] = new JsonArray((plan is null?formulas[i].Select(c=>c.CellReference!.Value!):plan.Where(k=>k.Sheet==i).Select(k=>XlsxCells.Reference(k.Col,k.Row))).Select(a=>(JsonNode?)JsonValue.Create(a)).ToArray()) });
        }
        var names = new JsonObject();
        foreach (var n in doc.Workbook.Workbook!.DefinedNames?.Elements<DefinedName>() ?? []) if (n.Name?.Value is { } name && n.LocalSheetId is null) names[name.ToUpperInvariant()] = n.Text;
        var dirty = new JsonArray(doc.CalculationChanges.Select(change => {
            var si = Array.FindIndex(sheets, s => s.SheetName == change.Sheet); var (c, r) = XlsxCells.Parse(change.Cell);
            return (JsonNode?)JsonValue.Create($"{si}:{r - 1}:{c - 1}"); }).ToArray());
        var input = new JsonObject { ["sheets"] = nodes, ["names"] = names, ["date1904"] = doc.Date1904, ["iteration"] = XlsxCalculationSettings.Read(doc), ["dirty"] = dirty, ["all"] = doc.CalculateAll || sheets.Any(s=>s.Data.Descendants<CellFormula>().Any(f=>f.FormulaType?.InnerText=="dataTable")) };
        var engine = new Engine(o => o.TimeoutInterval(TimeSpan.FromSeconds(30)).LimitRecursion(512).LimitMemory(512_000_000));
        engine.SetValue("loadCellKeys", (Jint.Native.JsValue)new ClrFunction(engine, "loadCellKeys", (_, args) =>
            new JsonArray(rows[(int)args[0].AsNumber()].Values.SelectMany(r => r.Elements<Cell>()).Select(c => (JsonNode?)JsonValue.Create(c.CellReference!.Value)).ToArray()).ToJsonString()));
        engine.SetValue("loadCell", (Jint.Native.JsValue)new ClrFunction(engine, "loadCell", (_, args) =>
        {
            var si = (int)args[0].AsNumber(); var address = args[1].AsString();
            var (col, row) = XlsxCells.Parse(address);
            var cell = rows[si].TryGetValue(row, out var r) ? XlsxCells.FindCell(r, col) : null;
            if (cell is null) return "null";
            // This scalar-only plan has proved untouched formulas independent of
            // the edit. Their typed cached results avoid recalculating the rest.
            if(planned is not null&&cell.CellFormula is not null&&!planned.Contains(new(si,row,col))) {
                var cached=cell.CellValue!.Text;var error=cell.DataType?.Value==CellValues.Error;var logical=cell.DataType?.Value==CellValues.Boolean;
                return new JsonObject{["v"]=error?"="+cached:logical?cached=="1"?"TRUE":"FALSE":cached,["literal"]=cell.DataType?.Value==CellValues.String,["s"]=new JsonObject{["fmt"]=cell.DataType?.Value==CellValues.String?"text":"general"}}.ToJsonString();
            }
            var formula = XlsxCells.Formula(sheets[si], cell);
            if (formula is null && sheets[si].ArrayOwner(col, row) is { } owner) return new JsonObject { ["v"] = "", ["spill"] = owner }.ToJsonString();
            var value = formula is { Length: > 0 } ? "=" + formula : XlsxCells.Display(doc, cell);
            var style = new JsonObject();
            if (XlsxCells.TypeOf(doc, cell) == "string") style["fmt"] = "text";
            else if (XlsxCells.TypeOf(doc, cell) == "date") style["fmt"] = "date";
            return new JsonObject { ["v"] = value, ["s"] = style, ["literal"] = formula is null && value.StartsWith('=') }.ToJsonString();
        }));
        try
        {
            var program = Program.Value;
            engine.Execute(in program);
            using var result = JsonDocument.Parse(engine.Call(engine.GetValue("calculateWorkbook"), Jint.Native.JsValue.Undefined, [new Jint.Native.JsString(input.ToJsonString())]).AsString());
            var rewritten = result.RootElement.EnumerateArray().Select(v => (v[0].GetInt32(), v[1].GetString())).ToHashSet();
            // Remove the previous array's cached children before writing the new spill footprint.
            if(plan is null)for (var i = 0; i < sheets.Length; i++) foreach (var cell in rows[i].Values.SelectMany(r => r.Elements<Cell>()))
            {
                var (c, r) = XlsxCells.Position(cell);
                if (cell.CellFormula is null && sheets[i].ArrayOwner(c, r) is { } owner && rewritten.Contains((i, owner))) { cell.CellValue = null; cell.DataType = null; }
            }
            foreach (var item in result.RootElement.EnumerateArray())
            {
                var si = item[0].GetInt32(); var (col, row) = XlsxCells.Parse(item[1].GetString()!);
                if (!rows[si].TryGetValue(row, out var dataRow)) { dataRow = XlsxCells.GetOrCreateRow(sheets[si].Data, row); rows[si][row] = dataRow; }
                var cell = XlsxCells.FindCell(dataRow, col);
                if (cell is null) { cell = new Cell { CellReference = XlsxCells.Reference(col, row) }; var next = dataRow.Elements<Cell>().FirstOrDefault(c => XlsxCells.Position(c).Col > col); if (next is null) dataRow.Append(cell); else dataRow.InsertBefore(cell, next); }
                if (cell.CellFormula is { } formula)
                {
                    if (item[3].ValueKind == JsonValueKind.String) { formula.FormulaType = CellFormulaValues.Array; formula.Reference = cell.CellReference!.Value + ":" + item[3].GetString(); }
                    else if (formula.FormulaType?.InnerText == "array") { formula.FormulaType = null; formula.Reference = null; }
                }
                var value = item[2];
                cell.RemoveAllChildren<InlineString>();
                switch (value.ValueKind)
                {
                    case JsonValueKind.Number: cell.DataType = null; cell.CellValue = new CellValue(value.GetDouble().ToString("G15", CultureInfo.InvariantCulture)); break;
                    case JsonValueKind.True: case JsonValueKind.False: cell.DataType = CellValues.Boolean; cell.CellValue = new CellValue(value.GetBoolean() ? "1" : "0"); break;
                    case JsonValueKind.String: cell.DataType = CellValues.String; cell.CellValue = new CellValue(value.GetString()!); break;
                    case JsonValueKind.Object when value.TryGetProperty("err", out var error):
                        cell.DataType = CellValues.Error; cell.CellValue = new CellValue(error.GetString()!); break;
                    default: cell.DataType = null; cell.CellValue = null; break;
                }
            }
        }
        catch (Exception e) when (e is not WriterException)
        {
            throw new WriterException(ErrorCode.FormatError, "Formula calculation could not finish: " + e.Message, "Reduce the formula range and save again.");
        }
    }
}
