using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

static class XlsxSplitPanes
{
    internal static string Read(Worksheet ws)
    {
        var p=ws.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.Pane;
        if(p is null||p.State?.InnerText is "frozen" or "frozenSplit")return "null";
        return new JsonObject{["x"]=p.HorizontalSplit?.Value??0,["y"]=p.VerticalSplit?.Value??0,["topLeft"]=p.TopLeftCell?.Value??"A1",["activePane"]=p.ActivePane?.InnerText??"bottomRight"}.ToJsonString();
    }
    internal static void Write(Worksheet ws,string json)
    {
        var node=JsonNode.Parse(json);var value=node as JsonObject;
        if(node is not null&&value is null)throw Invalid();
        var x=value?["x"]?.GetValue<double>()??0;var y=value?["y"]?.GetValue<double>()??0;
        var topLeft=value?["topLeft"]?.GetValue<string>()??"A1";
        var active=value?["activePane"]?.GetValue<string>()??(x>0&&y>0?"bottomRight":x>0?"topRight":"bottomLeft");
        if(!double.IsFinite(x)||!double.IsFinite(y)||x<0||y<0||x>100000||y>100000)throw Invalid();
        var (col,row)=XlsxCells.Parse(topLeft);
        if(col<1||col>16384||row<1||row>1048576||active is not ("topLeft" or "topRight" or "bottomLeft" or "bottomRight"))throw Invalid();
        var views=ws.GetFirstChild<SheetViews>();var view=views?.GetFirstChild<SheetView>();
        if(view is null){if(x==0&&y==0)return;if(views is null)ws.AddChild(views=new SheetViews(),true);view=views.AppendChild(new SheetView{WorkbookViewId=0U});}
        // Clearing a split must not remove a newly set freeze in the same save.
        if(x==0&&y==0){if(view.Pane?.State?.InnerText is not ("frozen" or "frozenSplit")){view.Pane=null;foreach(var s in view.Elements<Selection>().Where(s=>s.Pane is not null).ToList())s.Remove();}return;}
        view.Pane=new Pane{HorizontalSplit=x,VerticalSplit=y,TopLeftCell=XlsxCells.Reference(col,row),State=PaneStateValues.Split,ActivePane=active switch{"topLeft"=>PaneValues.TopLeft,"topRight"=>PaneValues.TopRight,"bottomLeft"=>PaneValues.BottomLeft,_=>PaneValues.BottomRight}};
        foreach(var s in view.Elements<Selection>().Where(s=>s.Pane is not null).ToList())s.Remove();
        view.Append(new Selection{Pane=view.Pane.ActivePane,ActiveCell=topLeft,SequenceOfReferences=new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue>{InnerText=topLeft}});
    }
    static WriterException Invalid()=>new(ErrorCode.Validation,"Invalid split panes","Use x/y positions in twips and a topLeft cell inside the worksheet.");
}
