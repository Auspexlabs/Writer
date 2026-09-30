using System.Text.Json.Nodes;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Writer.Formats.Xlsx;
static class XlsxSparklines
{
    const string Uri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}", Xm = "http://schemas.microsoft.com/office/excel/2006/main", Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    static readonly string[] Flags = ["markers","high","low","first","last","negative","displayXAxis","displayHidden","rightToLeft"];
    static readonly string[] Colors = ["colorSeries","colorNegative","colorAxis","colorMarkers","colorFirst","colorLast","colorHigh","colorLow"];
    internal static string Read(Worksheet ws)
    {
        var result = new JsonArray();
        foreach (var group in ws.Descendants<X14.SparklineGroup>()) {
            string? A(string key) => group.GetAttributes().FirstOrDefault(a=>a.LocalName==key&&a.NamespaceUri=="").Value;
            var item = new JsonObject { ["xml"] = group.OuterXml, ["type"] = A("type") ?? "line", ["empty"] = A("displayEmptyCellsAs") ?? "zero" };
            foreach (var key in Flags) item[key] = A(key) is "1" or "true";
            foreach (var key in new[]{"manualMin","manualMax","lineWeight"}) if (double.TryParse(A(key), NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) item[key] = n;
            foreach (var key in new[]{"minAxisType","maxAxisType"}) if (A(key) is { } value) item[key] = value;
            foreach (var key in Colors) if (group.ChildElements.FirstOrDefault(e => e.LocalName == key)?.GetAttributes().FirstOrDefault(a=>a.LocalName=="rgb").Value is { Length: > 0 } color) item[key] = "#" + color[^6..];
            if (group.ChildElements.FirstOrDefault(e => e.LocalName == "f" && e.NamespaceUri == Xm)?.InnerText is { Length: > 0 } dates) item["dateSource"] = dates;
            var lines = new JsonArray(); foreach (var line in group.Descendants<X14.Sparkline>()) lines.Add((JsonNode)new JsonObject { ["source"] = line.ChildElements.FirstOrDefault(e => e.LocalName == "f")?.InnerText ?? "", ["cell"] = line.ChildElements.FirstOrDefault(e => e.LocalName == "sqref")?.InnerText ?? "" });
            item["sparklines"] = lines; result.Add((JsonNode)item);
        }
        return result.ToJsonString();
    }
    internal static void Write(Worksheet ws, string json)
    {
        var items = JsonNode.Parse(json) as JsonArray ?? throw Bad("Sparklines must be an array of groups");
        var groups = new X14.SparklineGroups(); var used = new HashSet<string>();
        foreach (var node in items) {
            var item = node as JsonObject ?? throw Bad("Invalid sparkline group");
            string? S(string key) => item[key]?.GetValue<string>();
            var type = S("type") ?? "line"; var empty = S("empty") ?? "gap";
            if (type is not ("line" or "column" or "stacked") || empty is not ("gap" or "zero" or "span")) throw Bad("Invalid sparkline type or empty-cell mode");
            var group = S("xml") is { } raw ? new X14.SparklineGroup(raw) : new X14.SparklineGroup();
            void Attr(string key, string value) => group.SetAttribute(new OpenXmlAttribute("", key, "", value));
            Attr("type", type); Attr("displayEmptyCellsAs", empty);
            foreach (var key in Flags) Attr(key, item[key]?.GetValue<bool>() == true ? "1" : "0");
            foreach (var key in new[]{"manualMin","manualMax","lineWeight"}) if (item[key] is { } value) { var n = value.GetValue<double>(); if (!double.IsFinite(n) || key == "lineWeight" && (n <= 0 || n > 20)) throw Bad("Invalid sparkline scale"); Attr(key,n.ToString(CultureInfo.InvariantCulture)); }
            foreach (var key in new[]{"minAxisType","maxAxisType"}) if (S(key) is { } value) { if (value is not ("individual" or "group" or "custom")) throw Bad("Invalid sparkline axis mode"); Attr(key,value); }
            // Preserve imported theme colors unless a new RGB color was explicitly supplied.
            foreach (var key in Colors) if (S(key) is { } color) { color = color.TrimStart('#'); if (!System.Text.RegularExpressions.Regex.IsMatch(color,"^[0-9A-Fa-f]{6}$")) throw Bad("Sparkline colors use #RRGGBB"); foreach(var old in group.ChildElements.Where(e=>e.LocalName==key).ToArray())old.Remove(); OpenXmlElement colorNode=key switch {"colorSeries"=>new X14.SeriesColor(),"colorNegative"=>new X14.NegativeColor(),"colorAxis"=>new X14.AxisColor(),"colorMarkers"=>new X14.MarkersColor(),"colorFirst"=>new X14.FirstMarkerColor(),"colorLast"=>new X14.LastMarkerColor(),"colorHigh"=>new X14.HighMarkerColor(),_=>new X14.LowMarkerColor()};colorNode.SetAttribute(new OpenXmlAttribute("","rgb","","FF"+color.ToUpperInvariant()));group.AddChild(colorNode); }
            foreach(var old in group.ChildElements.Where(e=>e.LocalName is "sparklines" or "f").ToArray())old.Remove();
            if(S("dateSource") is { Length: > 0 } dates){ Attr("dateAxis","1");group.AddChild(new DocumentFormat.OpenXml.Office.Excel.Formula(dates)); } else Attr("dateAxis","0");
            var lines = new X14.Sparklines();
            foreach(var line in item["sparklines"] as JsonArray ?? throw Bad("Missing sparkline entries")) {
                var source = line?["source"]?.GetValue<string>() ?? ""; var cell = line?["cell"]?.GetValue<string>()?.Replace("$", "").ToUpperInvariant() ?? "";
                if(!System.Text.RegularExpressions.Regex.IsMatch(cell,"^[A-Z]{1,3}[1-9][0-9]{0,6}$") || !used.Add(cell) || source.Length == 0 || source.Length > 4096)throw Bad("Sparkline entries need a source and a unique destination cell");
                var address=XlsxCells.ParseRange(cell); if(address.Row1>1048576||address.Col1>16384)throw Bad("Sparkline destination is outside the worksheet");
                lines.Append(new X14.Sparkline(new DocumentFormat.OpenXml.Office.Excel.Formula(source),new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence(cell)));
            }
            if(!lines.HasChildren)continue;group.AddChild(lines);groups.Append(group);
        }
        var extensions=ws.GetFirstChild<WorksheetExtensionList>();foreach(var old in extensions?.Elements<WorksheetExtension>().Where(e=>e.Uri?.Value==Uri).ToArray()??[])old.Remove();
        if(groups.HasChildren){extensions??=ws.AppendChild(new WorksheetExtensionList());extensions.Append(new WorksheetExtension(groups){Uri=Uri});}
        if(extensions is {HasChildren:false})extensions.Remove();
    }
    static WriterException Bad(string message)=>new(ErrorCode.Validation,message,"Use groups containing type and sparklines:[{source,cell}].");
}
