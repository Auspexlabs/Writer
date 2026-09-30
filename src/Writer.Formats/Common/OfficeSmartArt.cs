using System.Globalization;
using System.Security;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using A = DocumentFormat.OpenXml.Drawing;
using D = DocumentFormat.OpenXml.Drawing.Diagrams;

namespace Writer.Formats.Common;

static class OfficeSmartArt
{
    const string Dgm = OfficeGraphics.DiagramUri, An = "http://schemas.openxmlformats.org/drawingml/2006/main", Dsp = "http://schemas.microsoft.com/office/drawing/2008/diagram", Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const long Width=5486400,Height=3086100;
    static string Esc(string? text)=>SecurityElement.Escape(text??"")??"";
    static string Id()=>"{"+Guid.NewGuid().ToString().ToUpperInvariant()+"}";
    internal sealed record Item(string Id,string Text,string? Parent);
    static XDocument Xml(OpenXmlPart part){using var stream=part.GetStream();return XDocument.Load(stream);}
    static void Save(OpenXmlPart part,XDocument xml){using var stream=new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));part.FeedData(stream);part.RootElement?.Reload();}
    static void Save(OpenXmlPart part,string xml)=>Save(part,XDocument.Parse(xml));
    internal static OpenXmlPart? Related(OpenXmlPart owner,string? id)=>id is not null&&owner.TryGetPartById(id,out var p)?p:null;
    internal static OpenXmlPart? Drawing(OpenXmlPart owner,OpenXmlPart data){var id=OfficeGraphics.SmartArtDrawingId(data);return Related(data,id)??Related(owner,id);}
    internal static (List<Item> Nodes,string Layout) Read(OpenXmlPart data)
    {
        XNamespace d=Dgm,a=An;var xml=Xml(data);var points=xml.Descendants(d+"pt").Where(p=>(string?)p.Attribute("type") is null or "node").ToArray();var ids=points.Select(p=>(string?)p.Attribute("modelId")??"").ToHashSet();
        var parents=xml.Descendants(d+"cxn").Where(c=>(string?)c.Attribute("type") is null or "parOf").GroupBy(c=>(string?)c.Attribute("destId")??"").ToDictionary(g=>g.Key,g=>(string?)g.First().Attribute("srcId"));
        var nodes=points.Select(p=>{var id=(string?)p.Attribute("modelId")??"";var parent=parents.GetValueOrDefault(id);return new Item(id,string.Join("\n",p.Element(d+"t")?.Elements(a+"p").Select(p=>string.Concat(p.Descendants(a+"t").Select(t=>t.Value)))??[]),parent is not null&&ids.Contains(parent)?parent:null);}).ToList();
        var layout=xml.Descendants(d+"prSet").Select(p=>(string?)p.Attribute("loTypeId")).FirstOrDefault(v=>v?.StartsWith("urn:writer:smartart:")==true)?.Split(':').Last()??"preserve";
        return(nodes,layout);
    }
    internal static A.GraphicData Write(OpenXmlPart owner,JsonElement spec,A.GraphicData? existing=null)
    {
        var nodes=spec.GetProperty("nodes").EnumerateArray().Select(n=>new Item(n.TryGetProperty("id",out var i)&&!string.IsNullOrEmpty(i.GetString())?i.GetString()!:Id(),n.GetProperty("text").GetString()??"",n.TryGetProperty("parent",out var p)&&p.ValueKind!=JsonValueKind.Null?p.GetString():null)).ToList();
        if(nodes.Count==0||nodes.Count>1000||nodes.Select(n=>n.Id).Distinct().Count()!=nodes.Count)throw Bad("SmartArt needs 1–1000 distinct nodes");
        var byId=nodes.ToDictionary(n=>n.Id);foreach(var n in nodes){var seen=new HashSet<string>{n.Id};var parent=n.Parent;while(parent is not null){if(!byId.TryGetValue(parent,out var p)||!seen.Add(parent))throw Bad("Invalid SmartArt hierarchy");parent=p.Parent;}}
        var rels=existing?.GetFirstChild<D.RelationshipIds>();var data=Related(owner,rels?.DataPart?.Value) as DiagramDataPart;
        var layout=spec.TryGetProperty("layout",out var l)?l.GetString()??"hierarchy":"hierarchy";
        if(layout=="preserve"&&data is not null){var before=Read(data);if(!before.Nodes.Select(n=>(n.Id,n.Parent)).SequenceEqual(nodes.Select(n=>(n.Id,n.Parent))))throw Bad("Choose a layout when changing the diagram hierarchy");UpdateText(owner,data,nodes);return existing!;}
        if(layout is not ("list" or "process" or "cycle" or "hierarchy"))throw Bad("Choose list, process, cycle or hierarchy");
        // Diagram layout's descendant axis walks the hierarchy in preorder. Keep its
        // presentation-point indices aligned even when a client sends children first.
        var ordered=new List<Item>();void Visit(Item n){ordered.Add(n);foreach(var child in nodes.Where(x=>x.Parent==n.Id))Visit(child);}
        foreach(var n in nodes.Where(n=>n.Parent is null))Visit(n);nodes=ordered;
        data??=owner.AddNewPart<DiagramDataPart>();
        var drawing=Drawing(owner,data) as DiagramPersistLayoutPart??owner.AddNewPart<DiagramPersistLayoutPart>();
        var layoutPart=Related(owner,rels?.LayoutPart?.Value) as DiagramLayoutDefinitionPart??owner.AddNewPart<DiagramLayoutDefinitionPart>();
        var colors=Related(owner,rels?.ColorPart?.Value) as DiagramColorsPart??owner.AddNewPart<DiagramColorsPart>();
        var style=Related(owner,rels?.StylePart?.Value) as DiagramStylePart??owner.AddNewPart<DiagramStylePart>();
        var root=Id();var pres=nodes.ToDictionary(n=>n.Id,_=>Id());var edges=Edges(nodes,layout);
        string Text(string text)=>$"<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"zh-CN\" sz=\"1600\"/><a:t>{Esc(text)}</a:t></a:r></a:p></dgm:t>";
        var points=$"<dgm:pt modelId=\"{root}\" type=\"doc\"><dgm:prSet loTypeId=\"urn:writer:smartart:{layout}\"/></dgm:pt>"+string.Concat(nodes.Select(n=>$"<dgm:pt modelId=\"{Esc(n.Id)}\" type=\"node\"><dgm:prSet/>{Text(n.Text)}</dgm:pt><dgm:pt modelId=\"{pres[n.Id]}\" type=\"pres\"><dgm:prSet presAssocID=\"{Esc(n.Id)}\" presName=\"node{nodes.IndexOf(n)}\" presStyleLbl=\"node0\"/></dgm:pt>"));
        points+=string.Concat(edges.Select((e,i)=>$"<dgm:pt modelId=\"{e.Id}\" type=\"pres\"><dgm:prSet presName=\"edge{i}\" presStyleLbl=\"edge\"/></dgm:pt>"));
        var connections=string.Concat(nodes.Select((n,i)=>$"<dgm:cxn modelId=\"{Id()}\" type=\"parOf\" srcId=\"{Esc(n.Parent??root)}\" destId=\"{Esc(n.Id)}\" srcOrd=\"{nodes.Take(i).Count(p=>p.Parent==n.Parent)}\" destOrd=\"0\"/><dgm:cxn modelId=\"{Id()}\" type=\"presOf\" srcId=\"{Esc(n.Id)}\" destId=\"{pres[n.Id]}\" srcOrd=\"0\" destOrd=\"0\"/>"));
        Save(data,$"<dgm:dataModel xmlns:dgm=\"{Dgm}\" xmlns:a=\"{An}\"><dgm:ptLst>{points}</dgm:ptLst><dgm:cxnLst>{connections}</dgm:cxnLst><dgm:bg/><dgm:whole/><dgm:extLst><a:ext uri=\"{Dsp}\"><dsp:dataModelExt xmlns:dsp=\"{Dsp}\" relId=\"{owner.GetIdOfPart(drawing)}\" minVer=\"{Dgm}\"/></a:ext></dgm:extLst></dgm:dataModel>");
        Save(layoutPart,Layout(nodes,layout,edges));
        Save(colors,$"<dgm:colorsDef xmlns:dgm=\"{Dgm}\" xmlns:a=\"{An}\" uniqueId=\"urn:writer:smartart:colors\"><dgm:title val=\"Writer\"/><dgm:desc val=\"Writer colors\"/><dgm:styleLbl name=\"node0\"><dgm:fillClrLst><a:srgbClr val=\"4472C4\"/></dgm:fillClrLst><dgm:linClrLst><a:srgbClr val=\"FFFFFF\"/></dgm:linClrLst><dgm:txFillClrLst><a:srgbClr val=\"FFFFFF\"/></dgm:txFillClrLst></dgm:styleLbl><dgm:styleLbl name=\"edge\"><dgm:linClrLst><a:srgbClr val=\"4472C4\"/></dgm:linClrLst></dgm:styleLbl></dgm:colorsDef>");
        Save(style,$"<dgm:styleDef xmlns:dgm=\"{Dgm}\" xmlns:a=\"{An}\" uniqueId=\"urn:writer:smartart:style\"><dgm:title val=\"Writer\"/><dgm:desc val=\"Writer style\"/><dgm:styleLbl name=\"node0\"><dgm:style><a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></dgm:style></dgm:styleLbl><dgm:styleLbl name=\"edge\"><dgm:style><a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef></dgm:style></dgm:styleLbl></dgm:styleDef>");
        Save(drawing,Render(nodes,pres,layout,edges));
        return new A.GraphicData(new D.RelationshipIds{DataPart=owner.GetIdOfPart(data),LayoutPart=owner.GetIdOfPart(layoutPart),ColorPart=owner.GetIdOfPart(colors),StylePart=owner.GetIdOfPart(style)}){Uri=Dgm};
    }
    static void UpdateText(OpenXmlPart owner,OpenXmlPart data,List<Item> nodes)
    {
        XNamespace d=Dgm,a=An,s=Dsp;var xml=Xml(data);var map=nodes.ToDictionary(n=>n.Id,n=>n.Text);
        void Replace(XElement body,string text){var ps=body.Elements(a+"p").ToList();var style=ps.FirstOrDefault()?.Element(a+"pPr");var run=body.Descendants(a+"rPr").FirstOrDefault();foreach(var p in ps)p.Remove();foreach(var line in text.Split('\n'))body.Add(new XElement(a+"p",style is null?null:new XElement(style),new XElement(a+"r",run is null?null:new XElement(run),new XElement(a+"t",line))));}
        foreach(var point in xml.Descendants(d+"pt")){var id=(string?)point.Attribute("modelId");if(id is not null&&map.TryGetValue(id,out var text)){var body=point.Element(d+"t");if(body is null){body=new XElement(d+"t",new XElement(a+"bodyPr"),new XElement(a+"lstStyle"));point.Add(body);}Replace(body,text);}}
        if(Drawing(owner,data) is { } drawing){var cache=Xml(drawing);var assoc=xml.Descendants(d+"pt").Where(p=>p.Element(d+"prSet")?.Attribute("presAssocID") is not null).ToDictionary(p=>(string)p.Attribute("modelId")!,p=>(string)p.Element(d+"prSet")!.Attribute("presAssocID")!);foreach(var shape in cache.Descendants(s+"sp")){var id=(string?)shape.Attribute("modelId");if(id is null)continue;id=assoc.GetValueOrDefault(id,id);if(map.TryGetValue(id,out var text)&&shape.Element(s+"txBody") is { } body)Replace(body,text);}Save(drawing,cache);}
        Save(data,xml);
    }
    static Dictionary<string,(long X,long Y,long W,long H)> Boxes(List<Item> nodes,string layout)
    {
        var depth=new Dictionary<string,int>();int Depth(Item n)=>depth.TryGetValue(n.Id,out var d)?d:depth[n.Id]=n.Parent is null?0:Depth(nodes.First(p=>p.Id==n.Parent))+1;
        foreach(var n in nodes)Depth(n);var levels=nodes.GroupBy(n=>depth[n.Id]).ToDictionary(g=>g.Key,g=>g.ToArray());var boxes=new Dictionary<string,(long X,long Y,long W,long H)>();
        for(var i=0;i<nodes.Count;i++){var n=nodes[i];double x,y,w,h;if(layout=="cycle"){var angle=-Math.PI/2+i*2*Math.PI/nodes.Count;w=Width*.2;h=Height*.2;x=Width/2+Width*.36*Math.Cos(angle)-w/2;y=Height/2+Height*.36*Math.Sin(angle)-h/2;}else if(layout=="hierarchy"){var level=levels[depth[n.Id]];w=Width*.85/Math.Max(1,level.Length);h=Height*.65/levels.Count;x=(Array.IndexOf(level,n)+.5)*Width/level.Length-w/2;y=(depth[n.Id]+.3)*Height/levels.Count;}else if(layout=="list"){w=Width*.86;h=Height*.8/nodes.Count;x=Width*.07;y=(i+.1)*Height/nodes.Count;}else{w=Width*.8/nodes.Count;h=Height*.6;x=(i+.1)*Width/nodes.Count;y=Height*.2;}boxes[n.Id]=((long)x,(long)y,(long)w,(long)h);}
        return boxes;
    }
    sealed record Edge(string Id,long X,long Y,long W,double Rotation);
    static List<Edge> Edges(List<Item> nodes,string layout)
    {
        var boxes=Boxes(nodes,layout);var pairs=new List<(Item From,Item To)>();
        if(layout=="hierarchy")foreach(var n in nodes.Where(n=>n.Parent is not null))pairs.Add((nodes.First(p=>p.Id==n.Parent),n));
        if(layout is "process" or "cycle")for(var i=0;i<nodes.Count-1;i++)pairs.Add((nodes[i],nodes[i+1]));
        if(layout=="cycle"&&nodes.Count>1)pairs.Add((nodes[^1],nodes[0]));
        var edges=new List<Edge>();
        foreach(var (from,to) in pairs) {
            var a=boxes[from.Id];var b=boxes[to.Id];double ax=a.X+a.W/2d,ay=a.Y+a.H/2d,bx=b.X+b.W/2d,by=b.Y+b.H/2d,dx=bx-ax,dy=by-ay;
            var distance=Math.Sqrt(dx*dx+dy*dy);if(distance<1)continue;var ux=dx/distance;var uy=dy/distance;
            double Inset(long w,long h)=>Math.Min(Math.Abs(ux)<1e-10?double.PositiveInfinity:w/2d/Math.Abs(ux),Math.Abs(uy)<1e-10?double.PositiveInfinity:h/2d/Math.Abs(uy));
            var first=Inset(a.W,a.H);var last=Inset(b.W,b.H);var length=distance-first-last;if(length<1)continue;
            ax+=ux*first;ay+=uy*first;bx-=ux*last;by-=uy*last;
            edges.Add(new Edge(Id(),(long)((ax+bx-length)/2),(long)((ay+by)/2),(long)length,Math.Atan2(dy,dx)*180/Math.PI));
        }
        return edges;
    }
    static string Layout(List<Item> nodes,string layout,List<Edge> edges)
    {
        var boxes=Boxes(nodes,layout);
        string F(double x)=>x.ToString("0.########",CultureInfo.InvariantCulture);
        var constraints=new StringBuilder();var children=new StringBuilder();
        foreach(var (edge,i) in edges.Select((e,i)=>(e,i))) {
            foreach(var (type,reference,factor) in new[]{("w","w",(double)edge.W/Width),("h","h",1d/Height),("l","w",(double)edge.X/Width),("t","h",(double)edge.Y/Height)})
                constraints.Append($"<dgm:constr type=\"{type}\" for=\"ch\" forName=\"edge{i}\" refType=\"{reference}\" fact=\"{F(factor)}\"/>");
            children.Append($"<dgm:layoutNode name=\"edge{i}\" styleLbl=\"edge\"><dgm:alg type=\"composite\"/><dgm:presOf/><dgm:shape type=\"line\" rot=\"{F(edge.Rotation)}\" lkTxEntry=\"1\" zOrderOff=\"-1\"/></dgm:layoutNode>");
        }
        // Explicit composite constraints preserve the chosen diagram geometry when Office
        // recalculates the layout; using a linear algorithm for every preset would flatten trees.
        foreach(var (n,i) in nodes.Select((n,i)=>(n,i))) {
            var b=boxes[n.Id];
            foreach(var (type,reference,factor) in new[]{("w","w",(double)b.W/Width),("h","h",(double)b.H/Height),("l","w",(double)b.X/Width),("t","h",(double)b.Y/Height)})
                constraints.Append($"<dgm:constr type=\"{type}\" for=\"ch\" forName=\"node{i}\" refType=\"{reference}\" fact=\"{F(factor)}\"/>");
            constraints.Append($"<dgm:constr type=\"primFontSz\" for=\"ch\" forName=\"node{i}\" op=\"equ\" val=\"16\"/>");
            children.Append($"<dgm:layoutNode name=\"node{i}\" styleLbl=\"node0\"><dgm:alg type=\"tx\"/><dgm:presOf axis=\"des\" ptType=\"node\" st=\"{i}\" cnt=\"1\"/><dgm:shape type=\"roundRect\"/></dgm:layoutNode>");
        }
        return $"<dgm:layoutDef xmlns:dgm=\"{Dgm}\" xmlns:a=\"{An}\" uniqueId=\"urn:writer:smartart:{layout}\" minVer=\"{Dgm}\"><dgm:title val=\"Writer {layout}\"/><dgm:desc val=\"Editable diagram\"/><dgm:layoutNode name=\"root\"><dgm:alg type=\"composite\"/><dgm:presOf/><dgm:shape/><dgm:constrLst>{constraints}</dgm:constrLst>{children}</dgm:layoutNode></dgm:layoutDef>";
    }
    static string Render(List<Item> nodes,Dictionary<string,string> pres,string layout,List<Edge> edges)
    {
        var boxes=Boxes(nodes,layout);
        var shapes=new StringBuilder();uint id=1;
        foreach(var edge in edges)shapes.Append($"<dsp:sp modelId=\"{edge.Id}\"><dsp:nvSpPr><dsp:cNvPr id=\"{id++}\" name=\"Connector\"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm rot=\"{(int)Math.Round(edge.Rotation*60000)}\"><a:off x=\"{edge.X}\" y=\"{edge.Y}\"/><a:ext cx=\"{edge.W}\" cy=\"1\"/></a:xfrm><a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"4472C4\"/></a:solidFill>{(layout is "process" or "cycle"?"<a:tailEnd type=\"triangle\"/>":"")}</a:ln></dsp:spPr></dsp:sp>");
        foreach(var n in nodes){var b=boxes[n.Id];shapes.Append($"<dsp:sp modelId=\"{pres[n.Id]}\"><dsp:nvSpPr><dsp:cNvPr id=\"{id++}\" name=\"{Esc(n.Text)}\"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x=\"{b.X}\" y=\"{b.Y}\"/><a:ext cx=\"{b.W}\" cy=\"{b.H}\"/></a:xfrm><a:prstGeom prst=\"roundRect\"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=\"4472C4\"/></a:solidFill><a:ln><a:noFill/></a:ln></dsp:spPr><dsp:txBody><a:bodyPr anchor=\"ctr\"/><a:lstStyle/><a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr sz=\"1600\"><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill></a:rPr><a:t>{Esc(n.Text)}</a:t></a:r></a:p></dsp:txBody></dsp:sp>");}
        return $"<dsp:drawing xmlns:dsp=\"{Dsp}\" xmlns:a=\"{An}\"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id=\"0\" name=\"\"/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{Width}\" cy=\"{Height}\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"{Width}\" cy=\"{Height}\"/></a:xfrm></dsp:grpSpPr>{shapes}</dsp:spTree></dsp:drawing>";
    }
    static WriterException Bad(string message)=>new(ErrorCode.Validation,message,"Use {layout,nodes:[{id,text,parent?}]} with an acyclic hierarchy.");
}
