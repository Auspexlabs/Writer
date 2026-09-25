using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;
using Writer.Formats.Docx;
using Writer.Formats.Pptx;
using Writer.Formats.Xlsx;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Compat;

/// <summary>Turns the compatibility model into a real docx, xlsx or pptx through the engine's own node API, so the
/// converted document is exactly what the editors and the CLI would have written themselves.</summary>
static class Writers
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    static string Cm(double v) => v.ToString("0.###", Inv) + "cm";
    /// <summary>The engine reads the real type from the bytes; the media type in the URL is only a formality.</summary>
    static string DataUrl(byte[] bytes) => "data:image/png;base64," + Convert.ToBase64String(bytes);

    static Dictionary<string, string> Props(params (string Name, object? Value)[] props)
    {
        var d = new Dictionary<string, string>();
        foreach (var (name, value) in props)
            if (value is not null) d[name] = value is IFormattable f ? f.ToString(null, Inv) : value.ToString()!;
        return d;
    }

    // ---------- docx ----------

    public static Document Docx(IReadOnlyList<Block> blocks, List<string> warnings)
    {
        var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.First(c => c.Kind == "body");
        Blocks(body, blocks, warnings);
        return doc;
    }

    static void Blocks(Node parent, IReadOnlyList<Block> blocks, List<string> warnings)
    {
        foreach (var b in blocks)
        {
            try
            {
                switch (b.Kind)
                {
                    case BlockKind.Heading:
                        Mutations.Add(parent, "heading", Props(("html", b.Html), ("level", b.Level), ("align", b.Align)), null);
                        break;
                    case BlockKind.Paragraph:
                        Mutations.Add(parent, "paragraph", Props(("html", b.Html), ("list", b.List), ("level", b.List is null ? null : b.Level), ("align", b.Align), ("style", b.Style)), null);
                        break;
                    case BlockKind.Table when b.Rows is { Count: > 0 }:
                        Table(parent, b.Rows, warnings);
                        break;
                    case BlockKind.Image when b.Image is { Length: > 0 }:
                        Mutations.Add(parent, "image", Props(("src", DataUrl(b.Image)), ("width", b.WidthCm is > 0 ? Cm(b.WidthCm.Value) : null), ("height", b.HeightCm is > 0 ? Cm(b.HeightCm.Value) : null)), null);
                        break;
                    case BlockKind.PageBreak:
                        Mutations.Add(parent, "pagebreak", [], null);
                        break;
                }
            }
            catch (WriterException ex)
            {
                warnings.Add($"{b.Kind}: {ex.Message}");
            }
        }
    }

    /// <summary>Cells take the next free grid column, as in html; spans reserve the columns to the right and below.
    /// The text goes in first, then the merges (merging moves absorbed cells' text in, and those are empty).</summary>
    static void Table(Node parent, List<List<TableCell>> rows, List<string> warnings)
    {
        var origins = new List<(int R, int C, TableCell Cell)>();
        var reserved = new HashSet<(int R, int C)>();
        for (var r = 0; r < rows.Count; r++)
        {
            var c = 0;
            foreach (var cell in rows[r])
            {
                while (reserved.Contains((r, c))) c++;
                origins.Add((r, c, cell));
                for (var dr = 0; dr < Math.Max(1, cell.RowSpan); dr++)
                    for (var dc = 0; dc < Math.Max(1, cell.ColSpan); dc++)
                        reserved.Add((r + dr, c + dc));
                c += Math.Max(1, cell.ColSpan);
            }
        }
        if (origins.Count == 0) return;
        var cols = reserved.Max(x => x.C) + 1;
        var rowCount = Math.Max(rows.Count, reserved.Max(x => x.R) + 1);
        var table = Mutations.Add(parent, "table", Props(("rows", rowCount), ("cols", cols)), null);
        var trs = table.Children.Where(x => x.Kind == "row").ToList();
        foreach (var (r, c, cell) in origins)
        {
            var td = trs[r].Children.Where(x => x.Kind == "cell").ElementAtOrDefault(c);
            if (td is null) continue;
            try { Mutations.Set(td, Props(("html", cell.Html), ("fill", cell.Fill))); }
            catch (WriterException ex) { warnings.Add($"table cell: {ex.Message}"); }
        }
        // bottom-right first: a merge only moves cells to its right and below, so earlier positions stay addressable
        for (var i = origins.Count - 1; i >= 0; i--)
        {
            var (r, c, cell) = origins[i];
            if (cell.ColSpan <= 1 && cell.RowSpan <= 1) continue;
            try
            {
                var td = table.Children.Where(x => x.Kind == "row").ElementAt(r).Children.Where(x => x.Kind == "cell").ElementAtOrDefault(c);
                if (td is null) continue;
                if (cell.ColSpan > 1) td = Mutations.Set(td, Props(("colspan", cell.ColSpan)));
                if (cell.RowSpan > 1) Mutations.Set(td, Props(("rowspan", cell.RowSpan)));
            }
            catch (WriterException ex) { warnings.Add($"table merge: {ex.Message}"); }
        }
    }

    // ---------- xlsx ----------

    public static Document Xlsx(IReadOnlyList<SheetModel> sheets, List<string> warnings)
    {
        var doc = (XlsxDocument)new XlsxAdapter().Create();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var recalc = false;
        for (var i = 0; i < sheets.Count; i++)
        {
            var model = sheets[i];
            var node = i == 0 ? doc.Root.Children[0] : Mutations.Add(doc.Root, "sheet", [], null);
            var name = XlsxDocument.ValidName(model.Name is { Length: > 0 } ? model.Name : $"Sheet{i + 1}");
            for (var n = 2; !names.Add(name); n++) name = XlsxDocument.ValidName($"{model.Name} ({n})");
            try { Mutations.Set(node, Props(("name", name))); } catch (WriterException ex) { warnings.Add($"sheet name: {ex.Message}"); }
            var data = doc.Sheets[i].Part.Worksheet!.GetFirstChild<SheetData>()!;
            foreach (var c in model.Cells)
            {
                if (c.Row < 1 || c.Col < 1) continue;
                var cell = XlsxCells.GetOrCreateCell(data, c.Col, c.Row);
                Style(doc, cell, c, warnings); // the format first: a date then keeps it instead of taking a default one on the way
                if (c.Formula is { Length: > 0 })
                {
                    cell.CellFormula = new CellFormula(c.Formula);
                    recalc = true;
                    // the cached result stays as the value shown until Excel recalculates
                    switch (c.Value)
                    {
                        case double d: XlsxCells.SetNumber(cell, d); break;
                        case bool b: XlsxCells.SetBool(cell, b); break;
                        case DateTime t: XlsxCells.SetNumber(cell, XlsxCells.ToSerial(t, false)); break;
                        case string s: cell.DataType = CellValues.String; cell.CellValue = new CellValue(s); break;
                    }
                }
                else
                {
                    switch (c.Value)
                    {
                        case string s: XlsxCells.SetString(doc, cell, s); break;
                        case double d: XlsxCells.SetNumber(cell, d); break;
                        case bool b: XlsxCells.SetBool(cell, b); break;
                        case DateTime t: XlsxCells.SetDate(doc, cell, t); break;
                    }
                }
            }
            try
            {
                var layout = Props(
                    ("merges", model.Merges.Count > 0 ? "[" + string.Join(",", model.Merges.Select(m => $"\"{m}\"")) + "]" : null),
                    ("widths", model.ColWidths.Count > 0 ? "{" + string.Join(",", model.ColWidths.Select(w => $"\"{XlsxCells.ColumnName(w.Key)}\":{w.Value.ToString("0.##", Inv)}")) + "}" : null),
                    ("heights", model.RowHeights.Count > 0 ? "{" + string.Join(",", model.RowHeights.Select(h => $"\"{h.Key}\":{h.Value.ToString("0.##", Inv)}")) + "}" : null));
                if (layout.Count > 0) Mutations.Set(node, layout);
            }
            catch (WriterException ex) { warnings.Add($"sheet layout: {ex.Message}"); }
            foreach (var img in model.Images)
            {
                try
                {
                    // ponytail: anchors assume default column widths and row heights; exact when the sheet keeps them
                    var x = Enumerable.Range(1, Math.Max(0, img.Col - 1)).Sum(col => (model.ColWidths.GetValueOrDefault(col, 8.43) + 0.71) * 0.19);
                    var y = Enumerable.Range(1, Math.Max(0, img.Row - 1)).Sum(row => model.RowHeights.GetValueOrDefault(row, 15) * 2.54 / 72);
                    Mutations.Add(node, "image", Props(("src", DataUrl(img.Bytes)), ("x", Cm(x)), ("y", Cm(y)), ("w", Cm(Math.Max(0.5, img.WidthCm))), ("h", Cm(Math.Max(0.5, img.HeightCm)))), null);
                }
                catch (WriterException ex) { warnings.Add($"picture: {ex.Message}"); }
            }
        }
        if (recalc) doc.RecalculateOnLoad();
        return doc;
    }

    static void Style(XlsxDocument doc, Cell cell, CellModel c, List<string> warnings)
    {
        var props = Props(
            ("bold", c.Bold ? "true" : null), ("italic", c.Italic ? "true" : null), ("underline", c.Underline ? "true" : null), ("strike", c.Strike ? "true" : null),
            ("wrap", c.Wrap ? "true" : null), ("font", c.Font), ("size", c.SizePt is > 0 ? c.SizePt.Value.ToString("0.#", Inv) : null), // the style setter takes points as a bare number
            ("color", c.Color), ("fill", c.Fill), ("format", c.Format), ("align", c.Align), ("valign", c.VAlign));
        try { doc.Styles.Set(cell, props.ToList()); } // one format per cell, not one per prop
        catch (WriterException)
        {
            foreach (var p in props)
            {
                try { doc.Styles.Set(cell, [p]); }
                catch (WriterException ex) { warnings.Add($"{cell.CellReference?.Value} {p.Key}: {ex.Message}"); }
            }
        }
    }

    // ---------- pptx ----------

    public static Document Pptx(DeckModel deck, List<string> warnings)
    {
        var doc = (PptxDocument)new PptxAdapter().Create();
        if (deck.WidthCm > 0 && deck.HeightCm > 0)
            doc.Presentation.Presentation!.SlideSize = new P.SlideSize { Cx = (int)Math.Round(deck.WidthCm * 360000), Cy = (int)Math.Round(deck.HeightCm * 360000) };
        var blank = doc.Root.Children.FirstOrDefault(c => c.Kind == "slide");
        foreach (var s in deck.Slides)
        {
            var slide = Mutations.Add(doc.Root, "slide", Props(("layout", "Blank")), null);
            if (s.Background is not null)
                try { Mutations.Set(slide, Props(("background", s.Background))); } catch (WriterException ex) { warnings.Add($"background: {ex.Message}"); }
            foreach (var sh in s.Shapes) Shape(slide, sh, warnings);
            if (s.Notes is { Length: > 0 })
                try { Mutations.Set(slide, Props(("notes", s.Notes))); } catch (WriterException ex) { warnings.Add($"notes: {ex.Message}"); }
        }
        if (deck.Slides.Count > 0) blank?.Remove(); // the template's own first slide
        return doc;
    }

    static void Shape(Node slide, ShapeModel sh, List<string> warnings)
    {
        var box = new (string, object?)[] { ("x", Cm(sh.X)), ("y", Cm(sh.Y)), ("w", Cm(Math.Max(0.1, sh.W))), ("h", Cm(Math.Max(0.1, sh.H))) };
        try
        {
            switch (sh.Kind)
            {
                case ShapeKind.Image when sh.Image is { Length: > 0 }:
                    Mutations.Add(slide, "image", Props([("src", DataUrl(sh.Image)), .. box]), null);
                    break;
                case ShapeKind.Table when sh.Rows is { Count: > 0 }:
                    var rows = sh.Rows.Select(r => "[" + string.Join(",", r.Select(c => NodeJson.Compact(w => w.WriteStringValue(Plain(c.Html))))) + "]");
                    Mutations.Add(slide, "table", Props([("data", "[" + string.Join(",", rows) + "]"), .. box]), null);
                    break;
                case ShapeKind.Text:
                    var geometry = sh.Geometry is "rect" or "roundRect" or "ellipse" or "triangle" or "diamond" or "rightArrow" ? sh.Geometry : "textbox";
                    var node = Mutations.Add(slide, "shape", Props([.. box, ("geometry", geometry), ("fill", sh.Fill), ("line", sh.Line), ("size", sh.SizePt is > 0 ? sh.SizePt.Value.ToString("0.#", Inv) + "pt" : null)]), null);
                    var paragraphs = sh.Paragraphs.Count > 0 ? sh.Paragraphs : [Block.Paragraph("")];
                    var first = node.Children.FirstOrDefault(c => c.Kind == "paragraph");
                    foreach (var p in paragraphs)
                    {
                        var props = Props(("html", p.Html), ("list", p.List), ("level", p.List is null ? null : p.Level), ("align", p.Align));
                        if (first is not null) { Mutations.Set(first, props); first = null; }
                        else Mutations.Add(node, "paragraph", props, null);
                    }
                    break;
            }
        }
        catch (WriterException ex)
        {
            warnings.Add($"{sh.Kind}: {ex.Message}");
        }
    }

    static readonly Regex Tags = new("<[^>]+>", RegexOptions.Compiled);

    /// <summary>Plain text of inline html: tags off, line breaks kept, entities decoded.</summary>
    public static string Plain(string html) =>
        System.Net.WebUtility.HtmlDecode(Tags.Replace(html.Replace("<br>", "\n").Replace("<br/>", "\n").Replace("<br />", "\n"), ""));
}
