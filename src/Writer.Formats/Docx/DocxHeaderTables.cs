using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

// Header tables retain their original OOXML layout and relationships. Editable
// cells carry their inline HTML separately from the rendered table preview.
static partial class DocxHeaderTables
{
    internal static IEnumerable<Node> Cells(Node node)
    {
        if (node.Kind == "cell") yield return node;
        foreach (var child in node.Children)
            foreach (var cell in Cells(child)) yield return cell;
    }
    static bool Editable(Node node) => node.Anchor is W.TableCell cell
        && !cell.Elements<W.Table>().Any()
        && !cell.Descendants().Any(e => e is W.Drawing or W.Picture or W.SimpleField or W.FieldChar or W.SdtElement or W.Hyperlink);

    internal static string Preview(DocxDocument doc, W.Table table, ref int ordinal)
    {
        var node = new DocxTable(doc, table);
        var cells = new Queue<Node>(Cells(node));
        var index = ordinal;
        var html = CellTag().Replace(Writer.Formats.Html.HtmlWriter.Fragment(node), m => {
            if (!cells.TryDequeue(out var cell)) return m.Value;
            var id = index++;
            return Editable(cell) ? m.Value[..^1] + " data-w-hf-cell=\"" + id.ToString(CultureInfo.InvariantCulture)
                + "\" data-w-hf-html=\"" + WebUtility.HtmlEncode(cell.GetProps()["html"]) + "\">" : m.Value;
        });
        ordinal = index;
        return html;
    }
    internal static void Apply(DocxDocument doc, OpenXmlCompositeElement container, string html)
    {
        var cells = container.Elements<W.Table>().SelectMany(t => Cells(new DocxTable(doc, t))).ToList();
        var seen = new HashSet<int>();
        foreach (Match tag in CellTag().Matches(html))
        {
            var id = CellId().Match(tag.Value); var payload = CellHtml().Match(tag.Value);
            if (!id.Success || !payload.Success || !int.TryParse(id.Groups[1].Value, out var i) || i >= cells.Count || !seen.Add(i) || !Editable(cells[i])) continue;
            var value = WebUtility.HtmlDecode(payload.Groups[1].Value);
            if (cells[i].GetProps()["html"] != value) cells[i].SetProp("html", value);
        }
    }
    internal static string WithoutPreviews(string html)
    {
        var output = new System.Text.StringBuilder(); var cursor = 0; var depth = 0;
        foreach (Match tag in DivTag().Matches(html))
        {
            var closing = tag.Groups[1].Success;
            if (depth == 0)
            {
                if (!closing && tag.Value.Contains("data-w-hf-table", StringComparison.OrdinalIgnoreCase)) { output.Append(html[cursor..tag.Index]); depth = 1; }
            }
            else if (!closing) depth++;
            else if (--depth == 0) cursor = tag.Index + tag.Length;
        }
        if (depth == 0) output.Append(html[cursor..]);
        return output.ToString();
    }
    [GeneratedRegex("<t[dh]\\b[^>]*>", RegexOptions.IgnoreCase)] private static partial Regex CellTag();
    [GeneratedRegex("\\bdata-w-hf-cell=\"(\\d+)\"", RegexOptions.IgnoreCase)] private static partial Regex CellId();
    [GeneratedRegex("\\bdata-w-hf-html=\"([^\"]*)\"", RegexOptions.IgnoreCase)] private static partial Regex CellHtml();
    [GeneratedRegex("<(/)?div\\b[^>]*>", RegexOptions.IgnoreCase)] private static partial Regex DivTag();
}
