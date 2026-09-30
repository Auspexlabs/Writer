using System.Globalization;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>Editable vector strokes and polygons, stored as real DrawingML custom geometry.</summary>
static class PptxFreeform
{
    static readonly XNamespace Ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    internal static string? Read(P.ShapeProperties? properties)
    {
        if (properties?.GetFirstChild<A.CustomGeometry>() is not { } geometry) return null;
        var root = XElement.Parse(geometry.OuterXml); var paths = new JsonArray();
        foreach (var path in root.Element(Ns + "pathLst")?.Elements(Ns + "path") ?? [])
        {
            if (!double.TryParse(path.Attribute("w")?.Value, Inv, out var w) || w <= 0 ||
                !double.TryParse(path.Attribute("h")?.Value, Inv, out var h) || h <= 0) return null;
            var points = new JsonArray(); var controls = new JsonObject(); var closed = false;
            foreach (var command in path.Elements())
            {
                if (command.Name == Ns + "close" && !closed) { closed = true; continue; }
                var kind=command.Name.LocalName;var expected=kind=="cubicBezTo"?3:kind=="quadBezTo"?2:1;
                if (closed || (points.Count==0?kind!="moveTo":kind is not ("lnTo" or "cubicBezTo" or "quadBezTo"))) return null;
                var coordinates=new JsonArray();
                foreach(var pt in command.Elements(Ns+"pt"))
                {
                    if(!double.TryParse(pt.Attribute("x")?.Value,Inv,out var x)||!double.TryParse(pt.Attribute("y")?.Value,Inv,out var y)||!double.IsFinite(x)||!double.IsFinite(y))return null;
                    coordinates.Add((JsonNode)new JsonArray(JsonValue.Create(x*1000/w),JsonValue.Create(y*1000/h)));
                }
                if(coordinates.Count!=expected)return null;
                var endpoint=coordinates[^1]!.DeepClone();coordinates.RemoveAt(coordinates.Count-1);
                if(coordinates.Count>0)controls[points.Count.ToString(Inv)]=coordinates;
                points.Add(endpoint);
            }
            if (points.Count == 0) continue;
            var fill = path.Attribute("fill")?.Value ?? "norm";
            if (fill is not ("none" or "norm")) return null;
            var data=new JsonObject { ["points"] = points, ["closed"] = closed, ["fill"] = fill,
                ["stroke"] = path.Attribute("stroke")?.Value is not ("false" or "0") };
            if(controls.Count>0)data["controls"]=controls;paths.Add((JsonNode)data);
        }
        return paths.Count == 0 ? null : new JsonObject { ["w"] = 1000, ["h"] = 1000, ["paths"] = paths }.ToJsonString();
    }

    internal static void Write(P.Shape shape, string json)
    {
        var spec = JsonNode.Parse(json) as JsonObject ?? throw new ArgumentException("Expected freeform geometry.");
        var width = spec["w"]?.GetValue<double>() ?? 1000; var height = spec["h"]?.GetValue<double>() ?? 1000;
        if (!double.IsFinite(width) || !double.IsFinite(height) || width <= 0 || height <= 0) throw new ArgumentException("Freeform dimensions must be positive.");
        var paths = spec["paths"] as JsonArray ?? throw new ArgumentException("Freeform paths are required.");
        if (paths.Count is < 1 or > 1000) throw new ArgumentException("Expected 1–1000 freeform paths.");
        var list = new XElement(Ns + "pathLst"); var count = 0;
        foreach (var path in paths)
        {
            var points = path?["points"] as JsonArray ?? throw new ArgumentException("Freeform points are required.");
            if (points.Count == 0 || (count += points.Count) > 100000) throw new ArgumentException("Expected at most 100000 nonempty freeform points.");
            var fill = path?["fill"]?.GetValue<string>() ?? "norm";
            if (fill is not ("none" or "norm")) throw new ArgumentException("Freeform fill must be norm or none.");
            var native = new XElement(Ns + "path", new XAttribute("w",1000000), new XAttribute("h",1000000), new XAttribute("fill",fill), new XAttribute("stroke",path?["stroke"]?.GetValue<bool>() ?? true));
            var controls=path?["controls"] as JsonObject;
            if(path?["controls"] is not null&&controls is null)throw new ArgumentException("Invalid curve controls.");
            foreach(var (key,value) in controls??new JsonObject())
                if(!int.TryParse(key,out var index)||index<1||index>=points.Count||key!=index.ToString(Inv)||value is not JsonArray a||a.Count is <1 or >2)throw new ArgumentException("Curve controls require a valid segment and one or two points.");
            XElement Coordinate(JsonNode? value)
            {
                var point=value as JsonArray;if(point?.Count!=2)throw new ArgumentException("Each freeform point needs x and y.");
                var x=point[0]!.GetValue<double>()/width*1000000;var y=point[1]!.GetValue<double>()/height*1000000;
                if(!double.IsFinite(x)||!double.IsFinite(y)||Math.Abs(x)>1e12||Math.Abs(y)>1e12)throw new ArgumentException("Invalid freeform coordinate.");
                return new XElement(Ns+"pt",new XAttribute("x",Math.Round(x)),new XAttribute("y",Math.Round(y)));
            }
            for (var i = 0; i < points.Count; i++)
            {
                var curve=controls?[i.ToString(Inv)] as JsonArray;
                if((count+=curve?.Count??0)>100000)throw new ArgumentException("Too many curve points.");
                var command=new XElement(Ns+(i==0?"moveTo":curve?.Count==2?"cubicBezTo":curve?.Count==1?"quadBezTo":"lnTo"));
                if(curve is not null)foreach(var control in curve)command.Add(Coordinate(control));
                command.Add(Coordinate(points[i]));native.Add(command);
            }
            if (path?["closed"]?.GetValue<bool>() == true) native.Add(new XElement(Ns + "close"));
            list.Add(native);
        }
        var xml = new XElement(Ns + "custGeom", new XAttribute(XNamespace.Xmlns+"a",Ns), new XElement(Ns+"avLst"), new XElement(Ns+"gdLst"), new XElement(Ns+"ahLst"), new XElement(Ns+"cxnLst"),
            new XElement(Ns+"rect",new XAttribute("l","0"),new XAttribute("t","0"),new XAttribute("r","w"),new XAttribute("b","h")), list);
        var geometry = new A.CustomGeometry(xml.ToString(SaveOptions.DisableFormatting));
        var properties = shape.ShapeProperties ??= new P.ShapeProperties();
        properties.RemoveAllChildren<A.PresetGeometry>(); properties.RemoveAllChildren<A.CustomGeometry>();
        if (properties.Transform2D is { } transform) properties.InsertAfter(geometry,transform); else properties.PrependChild(geometry);
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties is { } nv) nv.TextBox = null;
    }
}
