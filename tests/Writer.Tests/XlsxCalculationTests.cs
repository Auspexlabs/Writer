using Writer.Core;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxCalculationTests
{
    [Fact] public void Large_cached_scalar_workbook_recalculates_only_affected_chain_and_preserves_types()
    {
        using var file=new MemoryStream();
        using(var native=DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(file,DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook,true)) {
            var workbook=native.AddWorkbookPart();workbook.Workbook=new DocumentFormat.OpenXml.Spreadsheet.Workbook();var part=workbook.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();var data=new DocumentFormat.OpenXml.Spreadsheet.SheetData();part.Worksheet=new DocumentFormat.OpenXml.Spreadsheet.Worksheet(data);
            workbook.Workbook.Append(new DocumentFormat.OpenXml.Spreadsheet.Sheets(new DocumentFormat.OpenXml.Spreadsheet.Sheet{Name="Data",SheetId=1,Id=workbook.GetIdOfPart(part)}));
            for(var r=1;r<=10000;r++){var row=new DocumentFormat.OpenXml.Spreadsheet.Row{RowIndex=(uint)r};row.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="A"+r,CellValue=new("2")},new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="B"+r,CellFormula=new("$A"+r+"*2"),CellValue=new("4")},new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="C"+r,CellFormula=new("B"+r+"+1"),CellValue=new("5")});data.Append(row);}
            var first=data.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().First();
            first.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="H1",CellFormula=new("'Data'!$A$1+SUM (2,3E1)+IF(TRUE,0,1)+SUM(B9999)"),CellValue=new("38")},new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="I1",CellFormula=new("\"A1\"&A1"),CellValue=new("A12"),DataType=DocumentFormat.OpenXml.Spreadsheet.CellValues.String});
            first.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="D1",CellFormula=new("1/0"),CellValue=new("#DIV/0!"),DataType=DocumentFormat.OpenXml.Spreadsheet.CellValues.Error},new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="E1",CellFormula=new("IFERROR(D1,0)+A1"),CellValue=new("2")},new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="F1",CellFormula=new("TRUE()"),CellValue=new("1"),DataType=DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean},new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="G1",CellFormula=new("IF(F1,A1,0)"),CellValue=new("2")});
            var far=workbook.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();far.Worksheet=new(new DocumentFormat.OpenXml.Spreadsheet.SheetData(new DocumentFormat.OpenXml.Spreadsheet.Row(new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="XFD1048576",CellValue=new("9")}){RowIndex=1048576}));
            workbook.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>()!.Append(new DocumentFormat.OpenXml.Spreadsheet.Sheet{Name="Far",SheetId=2,Id=workbook.GetIdOfPart(far)});
            first.Append(new DocumentFormat.OpenXml.Spreadsheet.Cell{CellReference="J1",CellFormula=new("'Far'!$XFD$1048576+A1+A1"),CellValue=new("13")});
        }
        using var doc=new XlsxAdapter().Open(new MemoryStream(file.ToArray()));Set(doc,"A1","value","7");Set(doc,"A9876","value","12");using var opened=Reopen(doc);
        Assert.Equal("14",Value(opened,"B1"));Assert.Equal("15",Value(opened,"C1"));Assert.Equal("24",Value(opened,"B9876"));Assert.Equal("25",Value(opened,"C9876"));Assert.Equal("4",Value(opened,"B9999"));Assert.Equal("#DIV/0!",Value(opened,"D1"));Assert.Equal("7",Value(opened,"E1"));Assert.Equal("7",Value(opened,"G1"));
        Assert.Equal("43",Value(opened,"H1"));Assert.Equal("A17",Value(opened,"I1"));
        Assert.Equal("23",Value(opened,"J1"));PathResolver.Single(opened.Root,"/sheet[2]/cell[XFD1048576]").SetProp("value","20");
        PathResolver.Single(opened.Root,"/sheet[1]/cell[A9876]").Remove();using var cleared=Reopen(opened);Assert.Equal("0",Value(cleared,"B9876"));Assert.Equal("1",Value(cleared,"C9876"));
        Assert.Equal("34",Value(cleared,"J1"));
    }
    [Fact] public void Native_iteration_settings_and_circular_formula_caches_roundtrip()
    {
        using var doc = new XlsxAdapter().Create();
        doc.Root.SetProp("iteration", "{\"enabled\":true,\"count\":200,\"delta\":1e-10}");
        Set(doc,"A1","formula","(B1+2)/2"); Set(doc,"B1","formula","A1/2"); Set(doc,"C1","formula","A1+B1");
        using var opened = Reopen(doc);
        using var settings = System.Text.Json.JsonDocument.Parse(opened.Root.GetProps()["iteration"]);
        Assert.True(settings.RootElement.GetProperty("enabled").GetBoolean());
        Assert.Equal(200,settings.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(2,double.Parse(Value(opened,"C1"),System.Globalization.CultureInfo.InvariantCulture),8);
        var previous = opened.Root.GetProps()["iteration"];
        Assert.Throws<ArgumentException>(()=>opened.Root.SetProp("iteration","{\"enabled\":true,\"count\":0}"));
        Assert.Equal(previous,opened.Root.GetProps()["iteration"]);
        opened.Root.SetProp("iteration","{\"enabled\":false}");
        using var disabled = Reopen(opened); Assert.Equal("#CIRC!",Value(disabled,"A1"));
    }
    static void Set(Document doc, string a, string key, string value) => PathResolver.Single(doc.Root, "/sheet[1]/cell[" + a + "]").SetProp(key, value);
    static Document Reopen(Document doc) { var ms = new MemoryStream(); doc.Save(ms); ms.Position = 0; return new XlsxAdapter().Open(ms); }
    static string Value(Document doc, string a) => PathResolver.Single(doc.Root, "/sheet[1]/cell[" + a + "]").GetProps()["value"];
    [Fact] public void Regression_distributions_and_bonds_recalculate_in_the_embedded_engine()
    {
        using var d = new XlsxAdapter().Create();
        Set(d,"A1","formula","LINEST({1;9;5;7},{0;4;2;3},TRUE,TRUE)");
        Set(d,"D1","formula","T.TEST({1,2,3},{2,3,4},2,3)");
        Set(d,"D2","formula","PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)");
        Set(d,"D3","formula","T.INV.2T(0.05,6)");
        Set(d,"F1","value","=SEQUENCE(100)"); Set(d,"F1","type","string");
        Set(d,"G1","formula","ISFORMULA(F1)"); Set(d,"G2","formula","F1");
        using var opened=Reopen(d);
        static double N(string s)=>double.Parse(s,System.Globalization.CultureInfo.InvariantCulture);
        Assert.Equal(2,N(Value(opened,"A1")),9); Assert.Equal(1,N(Value(opened,"B1")),9);
        Assert.Equal(.28786413472669,N(Value(opened,"D1")),10);
        Assert.Equal(94.63,Math.Round(N(Value(opened,"D2")),2)); Assert.Equal(2.44691185114497,N(Value(opened,"D3")),9);
        Assert.Equal("false",Value(opened,"G1")); Assert.Equal("=SEQUENCE(100)",Value(opened,"G2")); Assert.Equal("",Value(opened,"F2"));
    }
    [Fact] public void Formulas_and_dependents_recalculate_on_save_with_number_text_boolean_and_error_types()
    {
        using var d = new XlsxAdapter().Create(); Set(d, "A1", "value", "4"); Set(d, "B1", "formula", "A1*2"); Set(d, "C1", "formula", "B1+3");
        Set(d, "D1", "formula", "\"hello\""); Set(d, "E1", "formula", "A1>2"); Set(d, "F1", "formula", "1/0");
        using var a = Reopen(d); Assert.Equal("8", Value(a, "B1")); Assert.Equal("11", Value(a, "C1")); Assert.Equal("hello", Value(a, "D1")); Assert.Equal("true", Value(a, "E1")); Assert.Equal("#DIV/0!", Value(a, "F1"));
        Set(a, "A1", "value", "7"); using var b = Reopen(a); Assert.Equal("14", Value(b, "B1")); Assert.Equal("17", Value(b, "C1"));
        PathResolver.Single(b.Root, "/sheet[1]/cell[A1]").Remove(); using var c = Reopen(b); Assert.Equal("0", Value(c, "B1"));
    }
    [Fact] public void Dynamic_array_results_save_reopen_and_shrink_without_stale_children()
    {
        using var d = new XlsxAdapter().Create(); Set(d, "A1", "formula", "SEQUENCE(3,2)"); Set(d, "D1", "formula", "SUM(A1#)");
        using var a = Reopen(d); Assert.Equal("6", Value(a, "B3")); Assert.Equal("21", Value(a, "D1")); Assert.Equal("A1", PathResolver.Single(a.Root, "/sheet[1]/cell[B3]").GetProps()["spill"]);
        Set(a, "A1", "formula", "SEQUENCE(2,1,10)"); using var b = Reopen(a); Assert.Equal("11", Value(b, "A2")); Assert.Equal("", Value(b, "B3")); Assert.Equal("21", Value(b, "D1"));
    }
    [Fact] public void Full_column_reducers_recalculate_after_deleting_a_value_and_lambda_arrays_roundtrip()
    {
        using var doc = new XlsxAdapter().Create();
        Set(doc,"A1","value","7");Set(doc,"A1048576","value","2");Set(doc,"B1","formula","SUM(A1:A1048576)");
        Set(doc,"C1","formula","MAP({1,2},LAMBDA(x,x*3))");
        using var first=Reopen(doc);Assert.Equal("9",Value(first,"B1"));Assert.Equal("3",Value(first,"C1"));Assert.Equal("6",Value(first,"D1"));
        PathResolver.Single(first.Root,"/sheet[1]/cell[A1]").Remove();
        using var next=Reopen(first);Assert.Equal("2",Value(next,"B1"));
    }
    [Fact] public void Empty_text_counts_and_array_errors_and_sparse_conditional_sums_save_correct_caches()
    {
        using var doc = new XlsxAdapter().Create();
        Set(doc,"A1","value","1"); Set(doc,"A2","formula","\"\""); Set(doc,"A3","value","2");
        Set(doc,"D1","formula","COUNTA(A1:A3)");
        Set(doc,"D2","formula","IFERROR(MAP({1,0,2},LAMBDA(x,1/x)),99)");
        Set(doc,"D3","formula","SUMIFS(A1:A1048576,A1:A1048576,\">0\")");
        using var first=Reopen(doc);
        Assert.Equal("3",Value(first,"D1")); Assert.Equal("1",Value(first,"D2"));
        Assert.Equal("99",Value(first,"E2")); Assert.Equal("0.5",Value(first,"F2")); Assert.Equal("3",Value(first,"D3"));
        PathResolver.Single(first.Root,"/sheet[1]/cell[A3]").Remove();
        using var next=Reopen(first); Assert.Equal("2",Value(next,"D1")); Assert.Equal("1",Value(next,"D3"));
    }
    [Fact] public void Worksheet_names_keep_their_scope_through_recalculation_and_save()
    {
        using var doc = new XlsxAdapter().Create();
        Mutations.Set(doc.Root,new Dictionary<string,string>{{"names","{\"Rate\":\"2\"}"}});
        var first=PathResolver.Single(doc.Root,"/sheet[1]");
        Mutations.Set(first,new Dictionary<string,string>{{"names","{\"Rate\":\"3\"}"}});
        var second=Mutations.Add(doc.Root,"sheet",new Dictionary<string,string>{{"name","Second"}},null);
        Set(doc,"A1","formula","Rate*10");PathResolver.Single(doc.Root,second.Path+"/cell[A1]").SetProp("formula","Rate*10");
        var firstName = first.GetProps()["name"].Replace("'", "''");
        PathResolver.Single(doc.Root,second.Path+"/cell[B1]").SetProp("formula",$"'{firstName}'!Rate*10");
        using var reopened=Reopen(doc);Assert.Equal("30",Value(reopened,"A1"));
        Assert.Equal("20",PathResolver.Single(reopened.Root,second.Path+"/cell[A1]").GetProps()["value"]);
        Assert.Equal("30",PathResolver.Single(reopened.Root,second.Path+"/cell[B1]").GetProps()["value"]);
        Assert.Contains("Rate",PathResolver.Single(reopened.Root,"/sheet[1]").GetProps()["names"]);
    }

}
