using System.Globalization;
using System.Security;
using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
namespace Writer.Formats.Xlsx;

static class XlsxPivots
{
    const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    static string Esc(string? s) => SecurityElement.Escape(s ?? "") ?? "";
    static int Index(OpenXmlElement e, string key) => int.TryParse(e.GetAttributes().FirstOrDefault(a => a.LocalName == key).Value, out var i) ? i : -1;
    internal static JsonArray Read(XlsxSheet sheet)
    {
        var result = new JsonArray();
        foreach (var part in sheet.Part.PivotTableParts)
        {
            var pivot = part.PivotTableDefinition; if (pivot is null) continue;
            var cache = part.PivotTableCacheDefinitionPart?.PivotCacheDefinition; var source = cache?.CacheSource?.GetFirstChild<WorksheetSource>();
            var rows = pivot.GetFirstChild<RowFields>()?.ChildElements.Select(x => (JsonNode?)JsonValue.Create(Index(x, "x"))).ToArray() ?? [];
            var cols = pivot.GetFirstChild<ColumnFields>()?.ChildElements.Where(x => Index(x,"x") >= 0).Select(x => (JsonNode?)JsonValue.Create(Index(x, "x"))).ToArray() ?? [];
            var model = new JsonObject { ["id"] = part.Uri.ToString(), ["name"] = pivot.Name?.Value, ["sourceSheet"] = source?.Sheet?.Value,
                ["sourceRange"] = source?.Reference?.Value, ["sourceTable"] = source?.Name?.Value, ["range"] = pivot.Location?.Reference?.Value,
                ["target"] = pivot.Location?.Reference?.Value?.Split(':')[0] ?? "A1", ["rows"] = new JsonArray(rows), ["cols"] = new JsonArray(cols),
                ["values"] = new JsonArray((pivot.GetFirstChild<DataFields>()?.Elements<DataField>() ?? []).Select(v => (JsonNode)new JsonObject { ["field"] = v.Field is { } field ? (int)field.Value : -1, ["fn"] = v.Subtotal?.InnerText ?? "sum", ["name"] = v.Name?.Value }).ToArray()),
                ["fields"] = new JsonArray((cache?.CacheFields?.Elements<CacheField>() ?? []).Select(f => (JsonNode?)JsonValue.Create(f.Name?.Value)).ToArray()),
                ["style"] = pivot.GetFirstChild<PivotTableStyle>()?.Name?.Value ?? "PivotStyleMedium9" };
            var cacheFields=cache?.CacheFields?.Elements<CacheField>().ToArray()??[];
            var grouping=new JsonArray();var groupedBases=new Dictionary<int,int>();
            for(var i=0;i<cacheFields.Length;i++)if(cacheFields[i].FieldGroup is {Base:{} basis} fg&&fg.GetFirstChild<DiscreteProperties>() is {} mapping&&fg.GetFirstChild<GroupItems>() is {} groupItems&&basis.Value<cacheFields.Length){
                var sourceItems=cacheFields[basis.Value].SharedItems?.ChildElements.ToArray()??[];var labels=groupItems.ChildElements.ToArray();var members=new Dictionary<int,JsonArray>();var map=mapping.ChildElements.ToArray();
                for(var j=0;j<map.Length&&j<sourceItems.Length;j++){var index=Index(map[j],"v");if(index<0||index>=labels.Length)continue;if(!members.TryGetValue(index,out var items))members[index]=items=[];items.Add(ReadItem(sourceItems[j]));}
                // Identity mappings are ungrouped source items. Recreating them
                // as named text groups would merge numeric 1 and text "1".
                var itemsModel=new JsonArray(members.Where(pair=>pair.Value.Any(v=>!JsonNode.DeepEquals(v,ReadItem(labels[pair.Key])))).Select(pair=>(JsonNode)new JsonObject{["name"]=labels[pair.Key].GetAttributes().FirstOrDefault(a=>a.LocalName=="v").Value??"",["items"]=pair.Value}).ToArray());
                if(itemsModel.Count>0)grouping.Add((JsonNode)new JsonObject{["field"]=(int)basis.Value,["items"]=itemsModel});groupedBases[i]=(int)basis.Value;
            }
            if(grouping.Count>0)model["groups"]=grouping;
            var dataFields=pivot.GetFirstChild<DataFields>()?.Elements<DataField>().ToArray()??[];
            for(var i=0;i<dataFields.Length;i++) {
                var v=dataFields[i];if(v.ShowDataAs?.InnerText is { } mode && mode!="normal")model["values"]![i]!["showAs"]=mode;
                if(v.BaseField is not null)model["values"]![i]!["baseField"]=v.BaseField.Value;
                if(v.BaseItem?.Value==1048829)model["values"]![i]!["baseItem"]="next";
                if(v.Field?.Value is {} fieldIndex&&fieldIndex<cacheFields.Length&&cacheFields[fieldIndex].Formula?.Value is {Length:>0} formula){model["values"]![i]!["formula"]=formula;model["values"]![i]!.AsObject().Remove("field");}
            }
            var filters=new JsonArray();var pivotFields=pivot.PivotFields?.Elements<PivotField>().ToArray()??[];
            for(var i=0;i<pivotFields.Length&&i<cacheFields.Length;i++) {
                var items=pivotFields[i].Items?.Elements<Item>().ToArray()??[];
                if(!items.Any(item=>item.Hidden?.Value==true))continue;
                var shared=(cacheFields[i].FieldGroup?.GetFirstChild<GroupItems>()??(OpenXmlCompositeElement?)cacheFields[i].SharedItems)?.ChildElements.ToArray()??[];var selected=new JsonArray();
                foreach(var item in items.Where(item=>item.Hidden?.Value!=true&&item.Index is not null))if(item.Index!.Value<shared.Length) {
                    var value=shared[item.Index.Value];var text=value.GetAttributes().FirstOrDefault(a=>a.LocalName=="v").Value??"";
                    selected.Add(value.LocalName=="n"&&double.TryParse(text,NumberStyles.Float,CultureInfo.InvariantCulture,out var n)?(JsonNode?)JsonValue.Create(n):value.LocalName=="b"?JsonValue.Create(text=="1"):JsonValue.Create(text));
                }
                filters.Add((JsonNode)new JsonObject{["field"]=i,["items"]=selected});
            }
            if(filters.Count>0)model["filters"]=filters;
            if (source?.Sheet?.Value is { } sourceName && source.Reference?.Value is { } sourceRange)
            {
                var owner = sheet.Doc.Sheets.FirstOrDefault(s => string.Equals(s.Sheet.Name?.Value, sourceName, StringComparison.OrdinalIgnoreCase));
                if (owner.Part is not null)
                {
                    var box = XlsxCells.ParseRange(sourceRange); var header = owner.Part.Worksheet!.GetFirstChild<SheetData>()?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == box.Row1);
                    var fields = Enumerable.Range(box.Col1, box.Col2 - box.Col1 + 1).Select(c => header is not null && XlsxCells.FindCell(header, c) is { } cell ? XlsxCells.Display(sheet.Doc, cell) : "").ToArray();
                    var cached = model["fields"]!.AsArray().Select(f => f?.GetValue<string>() ?? "").ToArray();
                    int Map(int index) {index=groupedBases.GetValueOrDefault(index,index);return index >= 0 && index < cached.Length ? Array.FindIndex(fields, f => string.Equals(f, cached[index], StringComparison.OrdinalIgnoreCase)) : index;}
                    foreach (var key in new[] { "rows", "cols" }) model[key] = new JsonArray(model[key]!.AsArray().Select(v => (JsonNode?)JsonValue.Create(Map(v!.GetValue<int>()))).ToArray());
                    foreach (var v in model["values"]!.AsArray()) {if(v!["field"] is {} vf)v["field"] = Map(vf.GetValue<int>());if(v["baseField"] is {} bf)v["baseField"]=Map(bf.GetValue<int>());}
                    foreach(var f in filters)f!["field"]=Map(f["field"]!.GetValue<int>());
                    foreach(var g in grouping)g!["field"]=Map(g["field"]!.GetValue<int>());
                    model["fields"] = new JsonArray(fields.Select(f => (JsonNode?)JsonValue.Create(f)).ToArray());
                }
            }
            result.Add((JsonNode)model);
        }
        return result;
    }
    sealed record Value(string Type, string Text)
    {
        public string Xml => Type == "m" ? "<m/>" : $"<{Type} v=\"{Esc(Text)}\"/>";
        public double? Number => Type == "n" && double.TryParse(Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n : null;
    }
    static JsonNode? ReadItem(OpenXmlElement item) {var text=item.GetAttributes().FirstOrDefault(a=>a.LocalName=="v").Value??"";return item.LocalName=="n"&&double.TryParse(text,NumberStyles.Float,CultureInfo.InvariantCulture,out var n)?JsonValue.Create(n):item.LocalName=="b"?JsonValue.Create(text=="1"):JsonValue.Create(text);}
    static Value JsonItem(JsonNode? item) => item is JsonValue v&&v.TryGetValue<double>(out var n)&&double.IsFinite(n)?new("n",n.ToString("G15",CultureInfo.InvariantCulture)):item is JsonValue b&&b.TryGetValue<bool>(out var boolean)?new("b",boolean?"1":"0"):item is JsonValue s&&s.TryGetValue<string>(out var text)?new(text.Length==0?"m":"s",text):throw Bad("Group items must be finite numbers, text or booleans");
    internal static void Write(XlsxDocument doc, XlsxSheet sheet, string json)
    {
        var list = JsonNode.Parse(json) as JsonArray ?? throw Bad("pivots must be an array");
        var old = sheet.Part.PivotTableParts.ToList(); var keep = new HashSet<PivotTablePart>();
        var before = Read(sheet);
        foreach (var node in list)
        {
            var n = node as JsonObject ?? throw Bad("Each pivot must be an object");
            var name = n["name"]?.GetValue<string>() ?? "";
            if (string.IsNullOrWhiteSpace(name) || name.Length > 255 || doc.Sheets.Where(s => s.Part != sheet.Part).SelectMany(s => s.Part.PivotTableParts).Any(p => string.Equals(p.PivotTableDefinition?.Name?.Value, name, StringComparison.OrdinalIgnoreCase))) throw Bad("Invalid or duplicate pivot name");
            var part = old.FirstOrDefault(p => n["id"]?.GetValue<string>() is { } id ? p.Uri.ToString() == id : p.PivotTableDefinition?.Name?.Value == name);
            if (part is not null && keep.Contains(part)) throw Bad("Duplicate pivot table");
            if (part is not null && JsonNode.DeepEquals(n, before.FirstOrDefault(p => p!["id"]!.GetValue<string>() == part.Uri.ToString()))) { keep.Add(part); continue; }
            if (part is not null && before.FirstOrDefault(p => p!["id"]!.GetValue<string>() == part.Uri.ToString()) is { } prior)
            {
                var a = n.DeepClone().AsObject(); var b = prior.DeepClone().AsObject(); foreach (var key in new[] { "sourceSheet", "sourceRange", "target", "range", "name", "sourceChanged" }) { a.Remove(key); b.Remove(key); }
                if (n["sourceChanged"]?.GetValue<bool>() == true || JsonNode.DeepEquals(a, b)) { part.PivotTableDefinition!.Name = name; if (part.PivotTableDefinition.Location is { } location && n["range"] is { } range) location.Reference = range.GetValue<string>(); if (part.PivotTableCacheDefinitionPart?.PivotCacheDefinition?.CacheSource?.GetFirstChild<WorksheetSource>() is { } source) { source.Sheet = n["sourceSheet"]?.GetValue<string>(); source.Reference = n["sourceRange"]?.GetValue<string>(); } keep.Add(part); continue; }
            }
            Build(doc, sheet, n, part, out var created); keep.Add(created);
        }
        foreach (var part in old.Where(p => !keep.Contains(p))) { ClearOutput(sheet, part.PivotTableDefinition?.Location?.Reference?.Value); sheet.Part.DeletePart(part); }
        var used = doc.Sheets.SelectMany(s => s.Part.PivotTableParts).Select(p => p.PivotTableCacheDefinitionPart).ToHashSet();
        foreach (var cache in doc.Workbook.PivotTableCacheDefinitionParts.Where(c => !used.Contains(c) && c.PivotCacheDefinition?.RefreshedBy?.Value == "Writer").ToList()) { var rel = doc.Workbook.GetIdOfPart(cache); foreach (var item in doc.Workbook.Workbook!.PivotCaches?.Elements<PivotCache>().Where(c => c.Id?.Value == rel).ToList() ?? []) item.Remove(); doc.Workbook.DeletePart(cache); }
        if (doc.Workbook.Workbook!.PivotCaches?.HasChildren == false) doc.Workbook.Workbook.PivotCaches.Remove();
        doc.RecalculateOnLoad();
    }
    static void ClearOutput(XlsxSheet sheet, string? reference)
    {
        if (reference is null) return; var box = XlsxCells.ParseRange(reference);
        foreach (var row in sheet.Data.Elements<Row>().Where(r => r.RowIndex?.Value >= box.Row1 && r.RowIndex?.Value <= box.Row2)) foreach (var cell in row.Elements<Cell>().ToList()) { var (c, _) = XlsxCells.Position(cell); if (c >= box.Col1 && c <= box.Col2) { XlsxCells.SetValue(sheet.Doc, sheet, cell, ""); if (cell.StyleIndex is null) cell.Remove(); } }
    }
    sealed class AggregateState
    {
        int count, numbers; double sum, mean, m2, product = 1, min = double.PositiveInfinity, max = double.NegativeInfinity; string? error;
        internal void Add(Value value) { if (value.Type != "m") count++; if (value.Type == "e") error = value.Text; if (value.Number is { } n) { numbers++; sum += n; var delta = n - mean; mean += delta / numbers; m2 += delta * (n - mean); product *= n; min = Math.Min(min, n); max = Math.Max(max, n); } }
        internal string Result(string fn) {
            if (fn == "count") return count.ToString(CultureInfo.InvariantCulture);
            if (fn == "countNums") return numbers.ToString(CultureInfo.InvariantCulture);
            if (error is not null) return error; if (numbers == 0) return "";
            double result;
            if (fn is "var" or "varP" or "stdDev" or "stdDevP") { var df = numbers - (fn is "var" or "stdDev" ? 1 : 0); if (df <= 0) return "#DIV/0!"; result = Math.Max(0, m2 / df); if (fn.StartsWith("std")) result = Math.Sqrt(result); }
            else result = fn switch { "sum" => sum, "average" => mean, "min" => min, "max" => max, "product" => product, _ => sum };
            return double.IsFinite(result) ? result.ToString("G15", CultureInfo.InvariantCulture) : "#NUM!";
        }
    }

    static void Build(XlsxDocument doc, XlsxSheet target, JsonObject spec, PivotTablePart? previous, out PivotTablePart part)
    {
        var previousModes=previous?.PivotTableDefinition?.DataFields?.Elements<DataField>().Select(v=>v.ShowDataAs?.InnerText??"normal").ToArray()??[];
        var sourceName = spec["sourceSheet"]?.GetValue<string>(); var reference = spec["sourceRange"]?.GetValue<string>();
        if (spec["sourceTable"]?.GetValue<string>() is { Length: > 0 } tableName)
            foreach (var s in doc.Sheets) if (s.Part.TableDefinitionParts.Select(p => p.Table).FirstOrDefault(t => string.Equals(t?.DisplayName?.Value, tableName, StringComparison.OrdinalIgnoreCase)) is { } table) { sourceName = s.Sheet.Name!.Value; var b = XlsxCells.ParseRange(table.Reference!.Value!); reference = XlsxCells.Reference(b.Col1, b.Row1) + ":" + XlsxCells.Reference(b.Col2, b.Row2 - (int)(table.TotalsRowCount?.Value ?? 0)); break; }
        var source = doc.Sheets.FirstOrDefault(s => string.Equals(s.Sheet.Name?.Value, sourceName, StringComparison.OrdinalIgnoreCase));
        if (source.Part is null || reference is null) throw Bad("The pivot needs a worksheet source");
        var box = XlsxCells.ParseRange(reference); var width = box.Col2 - box.Col1 + 1;var sourceWidth=width; var rowFields = spec["rows"]?.AsArray().Select(n => n!.GetValue<int>()).ToArray() ?? [];
        var colFields = spec["cols"]?.AsArray().Select(n => n!.GetValue<int>()).ToArray() ?? []; var values = spec["values"] as JsonArray ?? [];
        var valueFields = values.Select(v => v?["field"]?.GetValue<int>() ?? -1).ToArray();
        var functions = values.Select(v => v?["fn"]?.GetValue<string>() ?? "sum").ToArray();
        var formulas=values.Select(v=>v?["formula"]?.GetValue<string>()).ToArray();
        if (rowFields.Length == 0 || values.Count == 0 || rowFields.Concat(colFields).Concat(valueFields.Where((_,i)=>formulas[i] is null)).Any(f => f < 0 || f >= width) || rowFields.Concat(colFields).Distinct().Count() != rowFields.Length + colFields.Length || functions.Any(f => f is not ("sum" or "count" or "countNums" or "average" or "min" or "max" or "product" or "stdDev" or "stdDevP" or "var" or "varP"))||formulas.Select((f,i)=>(f,i)).Any(x=>x.f is not null&&(functions[x.i]!="sum"||string.IsNullOrWhiteSpace(values[x.i]?["name"]?.GetValue<string>())))) throw Bad("Invalid pivot fields or aggregation");
        var modes=values.Select(v=>v?["showAs"]?.GetValue<string>()??"normal").ToArray();
        var baseFields=values.Select(v=>v?["baseField"]?.GetValue<int>()??-1).ToArray();var filters=(spec["filters"]?.DeepClone() as JsonArray)??[];
        if(modes.Any(m=>m is not ("normal" or "percentOfRow" or "percentOfCol" or "percentOfTotal" or "index" or "runTotal" or "difference" or "percent" or "percentDiff"))||modes.Select((m,i)=>(m,i)).Any(x=>x.m is "runTotal" or "difference" or "percent" or "percentDiff"&&!rowFields.Concat(colFields).Contains(baseFields[x.i]))||filters.Any(f=>!rowFields.Concat(colFields).Contains(f?["field"]?.GetValue<int>()??-1)||f?["items"] is not JsonArray))throw Bad("Invalid pivot display calculation or filter");
        doc.Calculate();
        var sourceSheet = new XlsxSheet(doc, source.Sheet, source.Part); var sourceRows = sourceSheet.Data.Elements<Row>().Where(r => r.RowIndex?.Value >= box.Row1 && r.RowIndex?.Value <= box.Row2).ToDictionary(r => (int)r.RowIndex!.Value, r => r.Elements<Cell>().ToDictionary(c => XlsxCells.Position(c).Col));
        Value Get(int r, int c) { var cell = sourceRows.GetValueOrDefault(r)?.GetValueOrDefault(c); if (cell is null) return new("m", ""); var text = XlsxCells.Display(doc, cell); if (text.Length == 0) return new("m", ""); var type = XlsxCells.TypeOf(doc, cell); return type == "number" || type == "date" ? new("n", double.TryParse(cell.CellValue?.Text??text,NumberStyles.Float,CultureInfo.InvariantCulture,out var numeric)?numeric.ToString("G15",CultureInfo.InvariantCulture):text) : type == "bool" ? new("b", text == "true" ? "1" : "0") : type == "error" ? new("e", text) : new("s", text); }
        var fields = Enumerable.Range(box.Col1, width).Select(c => Get(box.Row1, c).Text).ToArray();
        if (fields.Any(string.IsNullOrWhiteSpace) || fields.Distinct(StringComparer.OrdinalIgnoreCase).Count() != fields.Length) throw Bad("Pivot source headers must be nonempty and unique");
        var data = Enumerable.Range(box.Row1 + 1, Math.Max(0, box.Row2 - box.Row1)).Select(r => Enumerable.Range(box.Col1, width).Select(c => Get(r, c)).ToArray()).ToArray();
        var sourceFields=fields.ToArray();var grouped=new Dictionary<int,(int Base,Dictionary<Value,Value> Map)>();var fieldFormulas=new Dictionary<int,string>();
        foreach(var group in spec["groups"] as JsonArray??[]){
            var basis=group?["field"]?.GetValue<int>()??-1;
            if(!rowFields.Concat(colFields).Contains(basis)||group?["items"] is not JsonArray groupItems||grouped.Values.Any(g=>g.Base==basis))throw Bad("Invalid pivot grouping field");
            var map=new Dictionary<Value,Value>();var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(var item in groupItems){var label=item?["name"]?.GetValue<string>();if(string.IsNullOrWhiteSpace(label)||label.Length>255||!names.Add(label)||item?["items"] is not JsonArray {Count:>0} members)throw Bad("Group names must be unique and contain source items");foreach(var member in members)if(!map.TryAdd(JsonItem(member),new("s",label)))throw Bad("Each source item may belong to only one group");}
            var fieldName=fields[basis]+" (Group)";while(fields.Contains(fieldName,StringComparer.OrdinalIgnoreCase))fieldName+="_";
            grouped[width]=(basis,map);fields=[..fields,fieldName];data=data.Select(row=>row.Append(map.GetValueOrDefault(row[basis],row[basis])).ToArray()).ToArray();
            rowFields=rowFields.Select(f=>f==basis?width:f).ToArray();colFields=colFields.Select(f=>f==basis?width:f).ToArray();baseFields=baseFields.Select(f=>f==basis?width:f).ToArray();foreach(var filter in filters)if(filter!["field"]!.GetValue<int>()==basis)filter["field"]=width;width++;
        }
        for(var j=0;j<formulas.Length;j++)if(formulas[j] is {} formula){var name=values[j]!["name"]!.GetValue<string>();if(fields.Contains(name,StringComparer.OrdinalIgnoreCase))throw Bad("Calculated field names must be unique and different from source fields");fieldFormulas[width]=formula.Trim().TrimStart('=');valueFields[j]=width;fields=[..fields,name];data=data.Select(row=>row.Append(new Value("m","")).ToArray()).ToArray();width++;}
        var shared = Enumerable.Range(0, width).Select(c => data.Select(r => r[c]).Distinct().ToArray()).ToArray(); var indices = shared.Select(items => items.Select((v, i) => (v, i)).ToDictionary(x => x.v, x => x.i)).ToArray();
        bool Included(Value value,JsonArray selected)=>selected.Any(item=>item is JsonValue j&&(value.Type=="n"&&j.TryGetValue<double>(out var num)&&num==value.Number||value.Type=="b"&&j.TryGetValue<bool>(out var boolean)&&(boolean?"1":"0")==value.Text||value.Type is "s" or "m"&&j.TryGetValue<string>(out var text)&&text==value.Text));
        bool Visible(Value[] row)=>filters.All(f=>Included(row[f!["field"]!.GetValue<int>()],f["items"]!.AsArray()));
        var visible=data.Where(Visible).ToArray();
        string Key(Value[] row, int[] axis) => string.Join(",", axis.Select(f => indices[f][row[f]]));
        int[][] Keys(int[] axis) => visible.Select(row => axis.Select(f => indices[f][row[f]]).ToArray()).DistinctBy(k => string.Join(",",k)).ToArray();
        var rowKeys = Keys(rowFields); var colKeys = colFields.Length > 0 ? Keys(colFields) : [Array.Empty<int>()];
        if (colKeys.Length == 0) colKeys = [Array.Empty<int>()];
        var rowIndices = rowKeys.Select((k,i)=>(Key:string.Join(",",k),Index:i)).ToDictionary(x=>x.Key,x=>x.Index);
        var colIndices = colKeys.Select((k,i)=>(Key:string.Join(",",k),Index:i)).ToDictionary(x=>x.Key,x=>x.Index);
        var nv = values.Count; var headerRows = Math.Max(1, colFields.Length + (nv > 1 ? 1 : 0));
        var (targetCol, targetRow) = XlsxCells.Parse(spec["target"]?.GetValue<string>() ?? "A1");
        var endCol = targetCol + rowFields.Length + (colKeys.Length + (colFields.Length > 0 ? 1 : 0)) * nv - 1; var endRow = targetRow + headerRows + rowKeys.Length;
        if (endCol > 16384 || endRow > 1048576) throw Bad("Pivot output exceeds the worksheet");
        var outputRange = XlsxCells.Reference(targetCol, targetRow) + ":" + XlsxCells.Reference(endCol, endRow);
        var groups = new Dictionary<(int R, int C, int V), AggregateState>();
        void Add(int r, int c, int j, Value value) { if (!groups.TryGetValue((r,c,j), out var state)) groups[(r,c,j)] = state = new AggregateState(); state.Add(value); }
        foreach (var row in visible) { var r = rowIndices[Key(row,rowFields)]; var c = colIndices[Key(row,colFields)]; for(var j=0;j<nv+(fieldFormulas.Count>0?sourceWidth:0);j++) { if(j<nv&&formulas[j] is not null)continue;var v=row[j<nv?valueFields[j]:j-nv]; Add(r,c,j,v); Add(r,-1,j,v); Add(-1,c,j,v); Add(-1,-1,j,v); } }
        var calculated=new Dictionary<(int R,int C,int J),string>();
        if(fieldFormulas.Count>0){
            var contexts=Enumerable.Range(-1,rowKeys.Length+1).SelectMany(r=>Enumerable.Range(-1,colKeys.Length+1).Select(c=>(r,c))).ToArray();
            JsonNode Total(int r,int c,int f){var value=groups.GetValueOrDefault((r,c,nv+f))?.Result("sum")??"";return value.StartsWith('#')?new JsonObject{["err"]=value}:JsonValue.Create(double.TryParse(value,NumberStyles.Float,CultureInfo.InvariantCulture,out var n)?n:0)!;}
            var formulaIndices=Enumerable.Range(0,nv).Where(j=>formulas[j] is not null).ToArray();
            var computed=XlsxCalculation.CalculatePivotFormulas(new JsonObject{["fields"]=new JsonArray(sourceFields.Select(f=>(JsonNode?)JsonValue.Create(f)).ToArray()),["formulas"]=new JsonArray(formulaIndices.Select(j=>(JsonNode?)JsonValue.Create(formulas[j])).ToArray()),["totals"]=new JsonArray(contexts.Select(x=>(JsonNode)new JsonArray(Enumerable.Range(0,sourceWidth).Select(f=>Total(x.r,x.c,f)).ToArray())).ToArray())});
            for(var i=0;i<contexts.Length;i++)for(var j=0;j<formulaIndices.Length;j++){var value=computed[i]![j];calculated[(contexts[i].r,contexts[i].c,formulaIndices[j])]=value is JsonValue v&&v.TryGetValue<string>(out var text)?text:value?.ToJsonString()??"";}
        }
        string Aggregate(int r, int c, int j) => formulas[j] is not null?calculated.GetValueOrDefault((r,c,j),""):groups.GetValueOrDefault((r,c,j))?.Result(functions[j]) ?? "";
        if (source.Part == target.Part && targetCol <= box.Col2 && endCol >= box.Col1 && targetRow <= box.Row2 && endRow >= box.Row1) throw Bad("Pivot output overlaps its source");
        var oldRange = previous?.PivotTableDefinition?.Location?.Reference?.Value; var oldBox = oldRange is null ? (Col1: 0, Row1: 0, Col2: 0, Row2: 0) : XlsxCells.ParseRange(oldRange);
        foreach (var r in target.Data.Elements<Row>().Where(r => r.RowIndex?.Value >= targetRow && r.RowIndex?.Value <= endRow)) foreach (var cell in r.Elements<Cell>()) { var (c, rr) = XlsxCells.Position(cell); if (c >= targetCol && c <= endCol && !(c >= oldBox.Col1 && c <= oldBox.Col2 && rr >= oldBox.Row1 && rr <= oldBox.Row2) && (cell.CellValue is not null || cell.InlineString is not null || cell.CellFormula is not null)) throw Bad("Pivot output would overwrite existing cells"); }
        var caches = doc.Workbook.Workbook!.PivotCaches ??= new PivotCaches(); var cacheId = caches.Elements<PivotCache>().Select(c => c.CacheId?.Value ?? 0).DefaultIfEmpty().Max() + 1;
        var cachePart = doc.Workbook.AddNewPart<PivotTableCacheDefinitionPart>(); var recordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
        string SharedXml(int c) => $"<sharedItems count=\"{shared[c].Length}\" containsString=\"{(shared[c].Any(v => v.Type == "s") ? 1 : 0)}\" containsNumber=\"{(shared[c].Any(v => v.Type == "n") ? 1 : 0)}\" containsBlank=\"{(shared[c].Any(v => v.Type == "m") ? 1 : 0)}\" containsSemiMixedTypes=\"{(shared[c].Select(v => v.Type).Distinct().Count() > 1 ? 1 : 0)}\">{string.Concat(shared[c].Select(v => v.Xml))}</sharedItems>";
        var fieldXml = string.Concat(fields.Select((f,c)=>{
            if(fieldFormulas.TryGetValue(c,out var formula))return $"<cacheField name=\"{Esc(f)}\" databaseField=\"0\" formula=\"{Esc(formula)}\"/>";
            if(grouped.TryGetValue(c,out var group))return $"<cacheField name=\"{Esc(f)}\" databaseField=\"0\"><fieldGroup base=\"{group.Base}\"><discretePr count=\"{shared[group.Base].Length}\">{string.Concat(shared[group.Base].Select(v=>$"<x v=\"{indices[c][group.Map.GetValueOrDefault(v,v)]}\"/>"))}</discretePr><groupItems count=\"{shared[c].Length}\">{string.Concat(shared[c].Select(v=>v.Xml))}</groupItems></fieldGroup></cacheField>";
            var parent=grouped.FirstOrDefault(g=>g.Value.Base==c);return $"<cacheField name=\"{Esc(f)}\">{SharedXml(c)}{(parent.Key>=sourceWidth?$"<fieldGroup par=\"{parent.Key}\"/>":"")}</cacheField>";
        }));
        cachePart.PivotCacheDefinition = new PivotCacheDefinition($"<pivotCacheDefinition xmlns=\"{Ns}\" xmlns:r=\"{Rel}\" r:id=\"{cachePart.GetIdOfPart(recordsPart)}\" recordCount=\"{data.Length}\" refreshedBy=\"Writer\" createdVersion=\"3\" refreshedVersion=\"3\" minRefreshableVersion=\"3\"><cacheSource type=\"worksheet\"><worksheetSource sheet=\"{Esc(sourceName)}\" ref=\"{Esc(reference)}\"/></cacheSource><cacheFields count=\"{width}\">{fieldXml}</cacheFields></pivotCacheDefinition>");
        recordsPart.PivotCacheRecords = new PivotCacheRecords($"<pivotCacheRecords xmlns=\"{Ns}\" count=\"{data.Length}\">{string.Concat(data.Select(row => "<r>" + string.Concat(row.Select((v,c) => fieldFormulas.ContainsKey(c)?"<m/>":$"<x v=\"{indices[c][v]}\"/>")) + "</r>"))}</pivotCacheRecords>");
        caches.Append(new PivotCache { CacheId = cacheId, Id = doc.Workbook.GetIdOfPart(cachePart) });
        part = previous ?? target.Part.AddNewPart<PivotTablePart>(); if (part.PivotTableCacheDefinitionPart is { } oldCache) part.DeletePart(oldCache); part.AddPart(cachePart);
        var fieldDefs = string.Concat(fields.Select((_, c) => { var axis = rowFields.Contains(c) ? "axisRow" : colFields.Contains(c) ? "axisCol" : null; var df = valueFields.Contains(c) ? 1 : 0; return axis is null ? $"<pivotField dataField=\"{df}\" showAll=\"0\"/>" : $"<pivotField axis=\"{axis}\" dataField=\"{df}\" compact=\"0\" outline=\"0\" showAll=\"0\" defaultSubtotal=\"0\"><items count=\"{shared[c].Length}\">{string.Concat(shared[c].Select((v,i) => $"<item x=\"{i}\"{(filters.Any(f=>f!["field"]!.GetValue<int>()==c&&!Included(v,f["items"]!.AsArray()))?" h=\"1\"":"")}/>"))}</items></pivotField>"; }));
        string Axis(string tag, int[] axis) => $"<{tag} count=\"{axis.Length}\">{string.Concat(axis.Select(f=>$"<field x=\"{f}\"/>"))}</{tag}>";
        string Items(string tag, int[][] keys, bool column) {
            var items = new List<string>(); var multiple = column && nv > 1;
            foreach (var key in keys) for (var j=0;j<(multiple?nv:1);j++) items.Add($"<i i=\"{j}\">{string.Concat(key.Select(k=>$"<x v=\"{k}\"/>"))}{(multiple?$"<x v=\"{j}\"/>":"")}</i>");
            if (!column || colFields.Length > 0) for (var j=0;j<(multiple?nv:1);j++) items.Add($"<i t=\"grand\" i=\"{j}\"><x/></i>");
            return $"<{tag} count=\"{items.Count}\">{string.Concat(items)}</{tag}>";
        }
        string Caption(int j) => values[j]?["name"]?.GetValue<string>() ?? functions[j] + " " + fields[valueFields[j]];
        var columnAxis = colFields.Concat(nv > 1 ? [-2] : Array.Empty<int>()).ToArray();
        var dataDefs = string.Concat(valueFields.Select((f,j)=>$"<dataField name=\"{Esc(Caption(j))}\" fld=\"{f}\" subtotal=\"{functions[j]}\" showDataAs=\"{modes[j]}\"{(baseFields[j]>=0?$" baseField=\"{baseFields[j]}\" baseItem=\"{(values[j]?["baseItem"]?.GetValue<string>()=="next"?1048829:1048828)}\"":"")}{(modes[j].StartsWith("percent")?" numFmtId=\"10\"":"")}/>"));
        part.PivotTableDefinition = new PivotTableDefinition($"<pivotTableDefinition xmlns=\"{Ns}\" name=\"{Esc(spec["name"]!.GetValue<string>())}\" cacheId=\"{cacheId}\" dataCaption=\"Values\" rowGrandTotals=\"1\" colGrandTotals=\"1\" compact=\"0\" compactData=\"0\" outline=\"0\" gridDropZones=\"1\" createdVersion=\"3\" updatedVersion=\"3\" minRefreshableVersion=\"3\"><location ref=\"{outputRange}\" firstHeaderRow=\"0\" firstDataRow=\"{headerRows}\" firstDataCol=\"{rowFields.Length}\"/><pivotFields count=\"{width}\">{fieldDefs}</pivotFields>{Axis("rowFields",rowFields)}{Items("rowItems",rowKeys,false)}{(columnAxis.Length>0?Axis("colFields",columnAxis)+Items("colItems",colKeys,true):"")}<dataFields count=\"{nv}\">{dataDefs}</dataFields><pivotTableStyleInfo name=\"{Esc(spec["style"]?.GetValue<string>() ?? "PivotStyleMedium9")}\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"1\" showColStripes=\"0\" showLastColumn=\"1\"/></pivotTableDefinition>");
        ClearOutput(target, oldRange);
        var dataCol = targetCol + rowFields.Length;
        void Cell(int r, int c, string value, bool text = false, bool bold = false) { var cell = XlsxCells.GetOrCreateCell(target.Data, c, r); XlsxCells.SetValue(doc, target, cell, value); if (text) XlsxCells.SetString(doc, cell, value); else if (value.StartsWith('#')) { cell.DataType = CellValues.Error; cell.CellValue = new CellValue(value); }
            if (cell.StyleIndex is null && bold) { var node = new XlsxCell(doc, target, c, r); node.SetProp("bold", "true"); node.SetProp("fill", "E8EEF7"); }
        }
        string Display(int r,int c,int j) {
            var mode=modes[j];var value=Aggregate(r,c,j);if(mode=="normal"||value.StartsWith('#'))return value;
            double Number(string v)=>double.TryParse(v,NumberStyles.Float,CultureInfo.InvariantCulture,out var n)?n:0;
            string Divide(double a,string b)=>b.StartsWith('#')?b:Number(b)==0?"#DIV/0!":(a/Number(b)).ToString("G15",CultureInfo.InvariantCulture);
            if(mode=="percentOfRow")return Divide(Number(value),Aggregate(r,-1,j));
            if(mode=="percentOfCol")return Divide(Number(value),Aggregate(-1,c,j));
            if(mode=="percentOfTotal")return Divide(Number(value),Aggregate(-1,-1,j));
            if(mode=="index")return Divide(Number(value)*Number(Aggregate(-1,-1,j)),(Number(Aggregate(r,-1,j))*Number(Aggregate(-1,c,j))).ToString("G15",CultureInfo.InvariantCulture));
            var rows=rowFields.Contains(baseFields[j]);var axis=rows?rowKeys:colKeys;var fieldsOnAxis=rows?rowFields:colFields;var index=rows?r:c;var pos=Array.IndexOf(fieldsOnAxis,baseFields[j]);
            if(index<0||pos<0)return mode=="runTotal"?value:"";
            var key=axis[index];var related=axis.Select((k,i)=>(k,i)).Where(x=>x.k.Select((v,k)=>k==pos||v==key[k]).All(b=>b)).Select(x=>x.i).ToArray();
            var here=Array.IndexOf(related,index);string Get(int i)=>rows?Aggregate(i,c,j):Aggregate(r,i,j);
            if(mode=="runTotal") {double sum=0;foreach(var i in related.Take(here+1)){var v=Get(i);if(v.StartsWith('#'))return v;sum+=Number(v);}return sum.ToString("G15",CultureInfo.InvariantCulture);}
            var offset=here+(values[j]?["baseItem"]?.GetValue<string>()=="next"?1:-1);if(offset<0||offset>=related.Length)return "";var baseline=Get(related[offset]);if(baseline.StartsWith('#'))return baseline;
            return mode=="difference"?(Number(value)-Number(baseline)).ToString("G15",CultureInfo.InvariantCulture):Divide(mode=="percent"?Number(value):Number(value)-Number(baseline),baseline);
        }
        for(var k=0;k<rowFields.Length;k++) Cell(targetRow+headerRows-1,targetCol+k,fields[grouped.TryGetValue(rowFields[k],out var rowGroup)?rowGroup.Base:rowFields[k]],true,true);
        for(var c=0;c<colKeys.Length;c++) for(var j=0;j<nv;j++) { var x=dataCol+c*nv+j; for(var k=0;k<colFields.Length;k++) Cell(targetRow+k,x,colKeys[c].Length>k?shared[colFields[k]][colKeys[c][k]].Text:"",true,true); if(colFields.Length==0||nv>1)Cell(targetRow+headerRows-1,x,Caption(j),true,true); }
        if(colFields.Length>0)for(var j=0;j<nv;j++){ var x=dataCol+colKeys.Length*nv+j; Cell(targetRow,x,"Grand Total",true,true);if(headerRows>1)Cell(targetRow+headerRows-1,x,nv>1?Caption(j):"Grand Total",true,true); }
        for(var r=0;r<rowKeys.Length;r++) { var y=targetRow+headerRows+r; for(var k=0;k<rowFields.Length;k++)Cell(y,targetCol+k,shared[rowFields[k]][rowKeys[r][k]].Text,true); for(var c=0;c<colKeys.Length;c++)for(var j=0;j<nv;j++)Cell(y,dataCol+c*nv+j,Display(r,c,j));if(colFields.Length>0)for(var j=0;j<nv;j++)Cell(y,dataCol+colKeys.Length*nv+j,Display(r,-1,j),false,true); }
        Cell(endRow,targetCol,"Grand Total",true,true);for(var c=0;c<colKeys.Length;c++)for(var j=0;j<nv;j++)Cell(endRow,dataCol+c*nv+j,Display(-1,c,j),false,true);if(colFields.Length>0)for(var j=0;j<nv;j++)Cell(endRow,dataCol+colKeys.Length*nv+j,Display(-1,-1,j),false,true);
        for(var r=targetRow+headerRows;r<=endRow;r++)for(var c=dataCol;c<=endCol;c++){var j=(c-dataCol)%nv;if(modes[j].StartsWith("percent"))new XlsxCell(doc,target,c,r).SetProp("format","0.00%");else if(j<previousModes.Length&&previousModes[j].StartsWith("percent"))new XlsxCell(doc,target,c,r).SetProp("format","General");}


    }
    static WriterException Bad(string message) => new(ErrorCode.Validation, message, "Use a named pivot, worksheet source, target, rows:[fields], cols:[fields], values:[{field,fn,name}]. Field indices start at zero.");
}
