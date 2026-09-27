using Writer.Core;
using Writer.Formats.Xlsx;
namespace Writer.Tests;
public class XlsxCalculationTests
{
    static void Set(Document doc, string a, string key, string value) => PathResolver.Single(doc.Root, "/sheet[1]/cell[" + a + "]").SetProp(key, value);
    static Document Reopen(Document doc) { var ms = new MemoryStream(); doc.Save(ms); ms.Position = 0; return new XlsxAdapter().Open(ms); }
    static string Value(Document doc, string a) => PathResolver.Single(doc.Root, "/sheet[1]/cell[" + a + "]").GetProps()["value"];
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
}
