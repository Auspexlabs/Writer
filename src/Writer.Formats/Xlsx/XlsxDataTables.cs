using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
namespace Writer.Formats.Xlsx;
static class XlsxDataTables
{
 internal static JsonArray Read(XlsxSheet sheet)
 {
  var list=new JsonArray();foreach(var f in sheet.Data.Descendants<CellFormula>().Where(f=>f.FormulaType?.InnerText=="dataTable"&&f.Reference is not null)){
   var b=XlsxCells.ParseRange(f.Reference!.Value!);if(b.Col1<2||b.Row1<2)continue;var table=new JsonObject{["range"]=XlsxCells.Reference(b.Col1-1,b.Row1-1)+":"+XlsxCells.Reference(b.Col2,b.Row2)};
   if(f.DataTable2D?.Value==true){table["rowInput"]=f.Input1Deleted?.Value==true?"#REF!":f.R1?.Value;table["colInput"]=f.Input2Deleted?.Value==true?"#REF!":f.R2?.Value;}
   else table[f.DataTableRow?.Value==true?"rowInput":"colInput"]=f.Input1Deleted?.Value==true?"#REF!":f.R1?.Value;list.Add((JsonNode)table);
  }return list;
 }
 internal static void Write(XlsxDocument doc,XlsxSheet sheet,string json)
 {
  var tables=JsonNode.Parse(json) as JsonArray??throw Bad("Data tables must be an array");var prepared=new List<(int Col,int Row,CellFormula Formula)>();var boxes=new List<(int C1,int R1,int C2,int R2)>();
  foreach(var t in tables){var b=XlsxCells.ParseRange(t?["range"]?.GetValue<string>()??"");if(b.Col1>=b.Col2||b.Row1>=b.Row2||b.Col2>16384||b.Row2>1048576||(long)(b.Col2-b.Col1)*(b.Row2-b.Row1)>100000)throw Bad("Invalid or oversized data table");if(boxes.Any(o=>b.Col1<=o.C2&&b.Col2>=o.C1&&b.Row1<=o.R2&&b.Row2>=o.R1))throw Bad("Data tables overlap");boxes.Add((b.Col1,b.Row1,b.Col2,b.Row2));
   var row=t!["rowInput"]?.GetValue<string>();var col=t["colInput"]?.GetValue<string>();foreach(var input in new[]{row,col}.Where(x=>!string.IsNullOrEmpty(x)&&x!="#REF!")){var (c,r)=XlsxCells.Parse(input!);if(c>16384||r>1048576||c>=b.Col1&&c<=b.Col2&&r>=b.Row1&&r<=b.Row2)throw Bad("Input cell is invalid or inside the data table");}if(string.IsNullOrEmpty(row)&&string.IsNullOrEmpty(col)||row==col)throw Bad("Supply one or two distinct input cells");
   var f=new CellFormula{FormulaType=CellFormulaValues.DataTable,Reference=XlsxCells.Reference(b.Col1+1,b.Row1+1)+":"+XlsxCells.Reference(b.Col2,b.Row2),CalculateCell=true,R1=(row??col)=="#REF!"?null:row??col,Input1Deleted=(row??col)=="#REF!",DataTable2D=row is not null&&col is not null,DataTableRow=row is not null&&col is null};if(row is not null&&col is not null){f.R2=col=="#REF!"?null:col;f.Input2Deleted=col=="#REF!";}prepared.Add((b.Col1+1,b.Row1+1,f));
  }
  foreach(var f in sheet.Data.Descendants<CellFormula>().Where(f=>f.FormulaType?.InnerText=="dataTable").ToList())f.Remove();
  foreach(var (c,r,f) in prepared)XlsxCells.GetOrCreateCell(sheet.Data,c,r).CellFormula=f;doc.RecalculateOnLoad();
 }
 static WriterException Bad(string message)=>new(ErrorCode.Validation,message,"Use [{range:'A1:D10',rowInput:'F1',colInput:'F2'}].");
}
