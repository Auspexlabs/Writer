using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Writer.Formats.Xlsx;

// A conservative dependency plan for large scalar workbooks. Any unrecognised
// syntax falls back to the full shared calculator; it never guesses dependencies.
static partial class XlsxCalculationDependencies
{
    internal readonly record struct Key(int Sheet,int Row,int Col);
    readonly record struct Edge(int Target,int Next);
    static ulong Pack(int sheet,int row,int col)=>((ulong)sheet<<34)|((ulong)(row-1)<<14)|(uint)(col-1);
    static Key Unpack(ulong key)=>new((int)(key>>34),(int)((key>>14)&0xfffff)+1,(int)(key&0x3fff)+1);
    static readonly HashSet<string> ScalarFunctions = new(StringComparer.OrdinalIgnoreCase) {
        "ABS","ROUND","ROUNDUP","ROUNDDOWN","INT","TRUNC","MOD","POWER","SQRT","EXP","LN","LOG","LOG10","SIGN","PI",
        "SUM","AVERAGE","MIN","MAX","COUNT","COUNTA","PRODUCT","IF","IFS","IFERROR","IFNA","AND","OR","NOT","XOR",
        "LEN","LEFT","RIGHT","MID","UPPER","LOWER","TRIM","CONCAT","CONCATENATE","EXACT","VALUE","N","T",
        "ISNUMBER","ISTEXT","ISLOGICAL","ISERROR","ISERR","ISNA","ISBLANK","ISEVEN","ISODD","NA","TRUE","FALSE"
    };
    // Numeric tokens precede references so scientific notation cannot turn into E5.
    [GeneratedRegex("\\G(?:\\s+|(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?|\"(?:[^\"]|\"\")*\"|(?<fn>[A-Za-z][A-Za-z0-9.]*)\\s*(?=\\()|(?:(?<sheet>'(?:[^']|'')*'|[A-Za-z_\\u4e00-\\u9fff][A-Za-z0-9_.\\u4e00-\\u9fff]*)!)?(?<cell>\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]*)|(?<bool>TRUE|FALSE)|[+*/^&%=(),<>-])",RegexOptions.IgnoreCase|RegexOptions.CultureInvariant)]
    private static partial Regex Tokens();
    internal static Key[]? Plan(XlsxDocument doc,XlsxSheet[] sheets,Cell[][] formulas)
    {
        var count=formulas.Sum(f=>f.Length);
        if(doc.CalculateAll||doc.CalculationChanges.Count==0||count<10000)return null;
        var names=sheets.Select((s,i)=>(s.SheetName,i)).ToDictionary(x=>x.SheetName,x=>x.i,StringComparer.OrdinalIgnoreCase);
        var queue=new Queue<ulong>();var affected=new HashSet<ulong>();var affectedFormulas=new HashSet<int>();
        foreach(var change in doc.CalculationChanges){if(!names.TryGetValue(change.Sheet,out var si))return null;var (c,r)=XlsxCells.Parse(change.Cell);var key=Pack(si,r,c);if(affected.Add(key))queue.Enqueue(key);}
        // One linked adjacency array avoids a List and backing array for every
        // referenced cell. Edges store formula indices, not repeated coordinates.
        var reverse=new Dictionary<ulong,int>();var edges=new List<Edge>(count);var targets=new ulong[count];var targetIndex=0;
        for(var si=0;si<sheets.Length;si++)foreach(var cell in formulas[si]) {
            if(cell.CellFormula!.FormulaType?.InnerText is { } type&&type!="normal"||cell.CellValue is null)return null;
            var (col,row)=XlsxCells.Position(cell);var target=Pack(si,row,col);targets[targetIndex]=target;if(affected.Contains(target))affectedFormulas.Add(targetIndex);
            var formula=cell.CellFormula.Text.AsSpan();var offset=0;
            // ValueMatch enumerates spans without allocating Match/Group objects
            // for every token in a million-formula workbook.
            foreach(var match in Tokens().EnumerateMatches(formula)) {
                if(match.Index!=offset)return null;offset+=match.Length;
                var token=formula.Slice(match.Index,match.Length).Trim();if(token.IsEmpty)continue;
                var first=token[0];
                if(first=='"'||char.IsDigit(first)||first=='.'||"+*/^&%=(),<>-".Contains(first))continue;
                var bang=token.LastIndexOf('!');
                if(bang<0 && offset<formula.Length && formula[offset]=='('){if(!ScalarFunctions.Contains(token.ToString()))return null;continue;}
                if(token.Equals("TRUE",StringComparison.OrdinalIgnoreCase)||token.Equals("FALSE",StringComparison.OrdinalIgnoreCase))continue;
                var owner=si;
                if(bang>=0){var sheetName=token[..bang].ToString().Trim('\'').Replace("''","'");if(!names.TryGetValue(sheetName,out owner))return null;token=token[(bang+1)..];}
                if(!Reference(token,out var c,out var r))return null;
                var source=Pack(owner,r,c);
                var previous=reverse.GetValueOrDefault(source,-1);reverse[source]=edges.Count;edges.Add(new(targetIndex,previous));
            }
            if(offset!=formula.Length)return null;
            targetIndex++;
        }
        while(queue.TryDequeue(out var key))if(reverse.TryGetValue(key,out var first))for(var i=first;i>=0;i=edges[i].Next){var index=edges[i].Target;affectedFormulas.Add(index);var child=targets[index];if(affected.Add(child))queue.Enqueue(child);}
        return affectedFormulas.Select(i=>Unpack(targets[i])).ToArray();
    }
    static bool Reference(ReadOnlySpan<char> value,out int col,out int row)
    {
        col=0;row=0;var i=0;if(value[i]=='$')i++;
        for(;i<value.Length&&char.IsAsciiLetter(value[i]);i++)col=col*26+(char.ToUpperInvariant(value[i])-'A'+1);
        if(i<value.Length&&value[i]=='$')i++;
        for(;i<value.Length;i++){if(!char.IsAsciiDigit(value[i])||row>1048576)return false;row=row*10+value[i]-'0';}
        return col is >0 and <=16384&&row is >0 and <=1048576;
    }
}
