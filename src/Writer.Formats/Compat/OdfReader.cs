using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Writer.Formats.Common;
using Writer.Formats.Xlsx;

namespace Writer.Formats.Compat;

/// <summary>OpenDocument text, spreadsheet and presentation (.odt .ods .odp, LibreOffice / OpenOffice / WPS export):
/// content.xml and styles.xml → blocks, sheets or a deck. Styles are resolved through their parent chain, so a run
/// that is bold through its paragraph style comes out bold.</summary>
public static class OdfReader
{
    static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    static readonly XNamespace TextNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    static readonly XNamespace Draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    static readonly XNamespace Xlink = "http://www.w3.org/1999/xlink";
    static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
    static readonly XNamespace Number = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
    static readonly XNamespace Presentation = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    static XDocument Xml(ZipArchive zip, string name)
    {
        var entry = zip.GetEntry(name) ?? throw new Core.WriterException(Core.ErrorCode.FormatError, $"The document has no {name}", "The file is damaged.");
        using var s = entry.Open();
        return XDocument.Load(s, LoadOptions.PreserveWhitespace);
    }

    static XDocument? TryXml(ZipArchive zip, string name) => zip.GetEntry(name) is null ? null : Xml(zip, name);

    // ---------- odt ----------

    public static List<Block> Text(ZipArchive zip, List<string> warnings)
    {
        var content = Xml(zip, "content.xml");
        var styles = new Styles(content, TryXml(zip, "styles.xml"));
        var body = content.Root?.Element(Office + "body")?.Element(Office + "text") ?? throw Damaged();
        var w = new TextWalker(styles, zip, warnings);
        w.Blocks(body, 0, null);
        return w.Out;
    }

    static Core.WriterException Damaged() => new(Core.ErrorCode.FormatError, "The document has no body", "The file is damaged.");

    sealed class TextWalker(Styles styles, ZipArchive zip, List<string> warnings)
    {
        public readonly List<Block> Out = [];
        public ZipArchive ZipOf => zip;
        public List<string> WarningsOf => warnings;

        public void Blocks(XElement parent, int listDepth, string? listType)
        {
            foreach (var e in parent.Elements())
            {
                var n = e.Name;
                if (n == TextNs + "h" || n == TextNs + "p")
                {
                    var ps = styles.Para(e.Attribute(TextNs + "style-name")?.Value);
                    if (ps.PageBreakBefore) Out.Add(Block.PageBreak());
                    var html = Inline(e, ps.Text, null, out var pictures);
                    Block b;
                    if (n == TextNs + "h") b = Block.Heading(int.TryParse(e.Attribute(TextNs + "outline-level")?.Value, out var lvl) ? lvl : 1, html);
                    else
                    {
                        b = Block.Paragraph(html);
                        if (listType is not null) { b.List = listType; b.Level = Math.Min(8, listDepth - 1); }
                        else if (ps.Heading > 0) { b.Kind = BlockKind.Heading; b.Level = Math.Clamp(ps.Heading, 1, 6); }
                        else if (ps.Style is not null) b.Style = ps.Style;
                    }
                    b.Align = ps.Align;
                    Out.Add(b);
                    Out.AddRange(pictures);
                }
                else if (n == TextNs + "list")
                {
                    var type = styles.ListType(e.Attribute(TextNs + "style-name")?.Value, listDepth) ?? listType ?? "bullet";
                    foreach (var item in e.Elements().Where(x => x.Name == TextNs + "list-item" || x.Name == TextNs + "list-header"))
                        Blocks(item, listDepth + 1, type);
                }
                else if (n == Table + "table") Out.Add(TableBlock(e));
                else if (n == TextNs + "section" || n == TextNs + "index-body" || n == TextNs + "table-of-content" || n == TextNs + "illustration-index" || n == TextNs + "alphabetical-index")
                    Blocks(e, listDepth, listType);
                else if (n == Draw + "frame")
                {
                    var pic = Picture(e);
                    if (pic is not null) Out.Add(pic);
                    else foreach (var box in e.Elements(Draw + "text-box")) Blocks(box, listDepth, listType);
                }
                else if (n == TextNs + "soft-page-break" || n == Office + "forms" || n == TextNs + "sequence-decls" || n == TextNs + "variable-decls" || n == TextNs + "user-field-decls") { }
                else if (n == TextNs + "tracked-changes" || n == Office + "annotation") { }
                else if (e.HasElements) Blocks(e, listDepth, listType);
            }
        }

        Block TableBlock(XElement table)
        {
            var rows = new List<List<TableCell>>();
            foreach (var row in table.Descendants(Table + "table-row"))
            {
                var cells = new List<TableCell>();
                foreach (var cell in row.Elements().Where(c => c.Name == Table + "table-cell"))
                {
                    var cs = styles.Cell(cell.Attribute(Table + "style-name")?.Value);
                    var parts = new List<string>();
                    foreach (var p in cell.Elements())
                    {
                        if (p.Name == TextNs + "p" || p.Name == TextNs + "h") parts.Add(Inline(p, styles.Para(p.Attribute(TextNs + "style-name")?.Value).Text, null, out _));
                        else if (p.Name == TextNs + "list") parts.AddRange(p.Descendants(TextNs + "p").Select(x => "• " + Inline(x, cs.Text, null, out _)));
                        else if (p.Name == Table + "table") parts.AddRange(p.Descendants(TextNs + "p").Select(x => Inline(x, cs.Text, null, out _)));
                    }
                    cells.Add(new TableCell(string.Join("<br>", parts))
                    {
                        ColSpan = Math.Max(1, Int(cell.Attribute(Table + "number-columns-spanned")?.Value)),
                        RowSpan = Math.Max(1, Int(cell.Attribute(Table + "number-rows-spanned")?.Value)),
                        Fill = cs.Fill,
                    });
                }
                if (cells.Count > 0) rows.Add(cells);
            }
            return Block.Table(rows);
        }

        /// <summary>Inline html of a paragraph: spans through their styles, spaces, tabs, line breaks, links. Pictures
        /// anchored inside the paragraph come out separately, after it.</summary>
        public string Inline(XElement p, TextStyle inherited, string? link, out List<Block> pictures)
        {
            var sb = new StringBuilder();
            var pics = new List<Block>();
            Walk(p, inherited, link, sb, pics);
            pictures = pics;
            return sb.ToString();
        }

        void Walk(XElement e, TextStyle st, string? link, StringBuilder sb, List<Block> pics)
        {
            foreach (var node in e.Nodes())
            {
                if (node is XText t) { sb.Append(Run(t.Value, st, link)); continue; }
                if (node is not XElement c) continue;
                var n = c.Name;
                if (n == TextNs + "span") Walk(c, styles.Text(c.Attribute(TextNs + "style-name")?.Value, st), link, sb, pics);
                else if (n == TextNs + "a") Walk(c, st, c.Attribute(Xlink + "href")?.Value ?? link, sb, pics);
                else if (n == TextNs + "s") sb.Append(Run(new string(' ', Math.Max(1, Int(c.Attribute(TextNs + "c")?.Value))), st, link));
                else if (n == TextNs + "tab") sb.Append(Run(" ", st, link));
                else if (n == TextNs + "line-break") sb.Append("<br>");
                else if (n == TextNs + "note" || n == Office + "annotation" || n == TextNs + "soft-page-break" || n == TextNs + "tracked-changes" || n == TextNs + "bookmark" || n == TextNs + "bookmark-start" || n == TextNs + "bookmark-end" || n == TextNs + "reference-mark") { }
                else if (n == Draw + "frame") { var pic = Picture(c); if (pic is not null) pics.Add(pic); else foreach (var box in c.Elements(Draw + "text-box")) Walk(box, st, link, sb, pics); }
                else if (n == TextNs + "list" || n == TextNs + "p") { if (sb.Length > 0) sb.Append("<br>"); Walk(c, st, link, sb, pics); }
                else Walk(c, st, link, sb, pics);
            }
        }

        static string Run(string text, TextStyle s, string? link) =>
            Writer.Formats.Compat.Inline.Run(text, s.Bold == true, s.Italic == true, s.Underline == true, s.Strike == true, s.Color, s.SizePt, s.Font, s.Highlight, link);

        public Block? Picture(XElement frame)
        {
            var image = frame.Element(Draw + "image");
            var href = image?.Attribute(Xlink + "href")?.Value;
            if (image is null) return null;
            byte[]? bytes = null;
            if (href is { Length: > 0 } && zip.GetEntry(href.TrimStart('.', '/')) is { } entry)
            {
                using var s = entry.Open();
                var ms = new MemoryStream();
                s.CopyTo(ms);
                bytes = ms.ToArray();
            }
            else if (image.Element(Office + "binary-data") is { } b64)
            {
                try { bytes = Convert.FromBase64String(Regex.Replace(b64.Value, @"\s+", "")); } catch (FormatException) { }
            }
            if (bytes is null || bytes.Length < 8) { warnings.Add($"picture {href} dropped: not in the file"); return Block.Paragraph(""); }
            try { ImageInfo.Read(bytes); }
            catch (Core.WriterException) { warnings.Add($"picture {href} dropped: not a PNG, JPEG, GIF or BMP"); return Block.Paragraph(""); }
            return Block.Picture(bytes, Length(frame.Attribute(Svg + "width")?.Value), Length(frame.Attribute(Svg + "height")?.Value));
        }
    }

    // ---------- ods ----------

    public static List<SheetModel> Sheets(ZipArchive zip, List<string> warnings)
    {
        var content = Xml(zip, "content.xml");
        var styles = new Styles(content, TryXml(zip, "styles.xml"));
        var body = content.Root?.Element(Office + "body")?.Element(Office + "spreadsheet") ?? throw Damaged();
        var sheets = new List<SheetModel>();
        var walker = new TextWalker(styles, zip, warnings);
        foreach (var table in body.Elements(Table + "table"))
        {
            var sheet = new SheetModel { Name = table.Attribute(Table + "name")?.Value ?? $"Sheet{sheets.Count + 1}" };
            var col = 1;
            foreach (var c in table.Descendants(Table + "table-column"))
            {
                var repeat = Math.Clamp(Int(c.Attribute(Table + "number-columns-repeated")?.Value, 1), 1, 1024);
                var width = styles.ColumnWidthCm(c.Attribute(Table + "style-name")?.Value);
                if (width is > 0 && repeat < 200) for (var i = 0; i < repeat; i++) sheet.ColWidths[col + i] = Math.Round(width.Value / 0.19 - 0.71, 2);
                col += repeat;
            }
            var row = 1;
            foreach (var r in table.Descendants(Table + "table-row"))
            {
                var repeat = Math.Max(1, Int(r.Attribute(Table + "number-rows-repeated")?.Value, 1));
                var height = styles.RowHeightPt(r.Attribute(Table + "style-name")?.Value);
                var cells = new List<CellModel>();
                var merges = new List<(int C, int R, int Cs, int Rs)>();
                var ci = 1;
                foreach (var cell in r.Elements().Where(x => x.Name == Table + "table-cell" || x.Name == Table + "covered-table-cell"))
                {
                    var creps = Math.Max(1, Int(cell.Attribute(Table + "number-columns-repeated")?.Value, 1));
                    if (cell.Name == Table + "table-cell")
                    {
                        var model = CellOf(cell, styles, walker);
                        var cs = Int(cell.Attribute(Table + "number-columns-spanned")?.Value, 1);
                        var rs = Int(cell.Attribute(Table + "number-rows-spanned")?.Value, 1);
                        if (model is not null)
                            for (var k = 0; k < Math.Min(creps, 1000); k++)
                            {
                                var copy = k == 0 ? model : Clone(model);
                                copy.Col = ci + k;
                                cells.Add(copy);
                            }
                        if (cs > 1 || rs > 1) merges.Add((ci, 0, cs, rs));
                    }
                    ci += creps;
                    if (ci > 16384) break;
                }
                var has = cells.Count > 0 || merges.Count > 0 || height is > 0;
                var reps = has ? Math.Min(repeat, 1000) : 0;
                for (var k = 0; k < reps && row + k <= 1048576; k++)
                {
                    foreach (var c in cells) { var copy = k == 0 ? c : Clone(c); copy.Row = row + k; sheet.Cells.Add(copy); }
                    foreach (var (mc, _, cs, rs) in merges) sheet.Merges.Add($"{XlsxCells.Reference(mc, row + k)}:{XlsxCells.Reference(mc + cs - 1, row + k + rs - 1)}");
                    if (height is > 0 && Math.Abs(height.Value - 12.8) > 0.3) sheet.RowHeights[row + k] = Math.Round(height.Value, 2);
                }
                row += repeat;
                if (row > 1048576) break;
            }
            foreach (var frame in table.Descendants(Draw + "frame"))
            {
                if (walker.Picture(frame) is not { Kind: BlockKind.Image } img) continue;
                var pic = new SheetImage { Bytes = img.Image!, WidthCm = img.WidthCm ?? 8, HeightCm = img.HeightCm ?? 6 };
                if (frame.Attribute(Table + "end-cell-address")?.Value is { } anchor && Regex.Match(anchor, @"\.\$?([A-Z]+)\$?(\d+)$") is { Success: true } m)
                { pic.Col = Math.Max(1, XlsxCells.ColumnIndex(m.Groups[1].Value)); pic.Row = Math.Max(1, int.Parse(m.Groups[2].Value)); }
                sheet.Images.Add(pic);
            }
            sheets.Add(sheet);
        }
        return sheets;
    }

    static CellModel Clone(CellModel c) => c.MemberwiseCloneOf();

    static CellModel? CellOf(XElement cell, Styles styles, TextWalker walker)
    {
        var type = cell.Attribute(Office + "value-type")?.Value;
        var formula = cell.Attribute(Table + "formula")?.Value;
        var cs = styles.Cell(cell.Attribute(Table + "style-name")?.Value);
        object? value = null;
        string? format = cs.Format;
        switch (type)
        {
            case "float" or "percentage" or "currency":
                if (double.TryParse(cell.Attribute(Office + "value")?.Value, NumberStyles.Float, Inv, out var d)) value = d;
                if (type == "percentage" && format is null) format = "0%";
                break;
            case "boolean":
                value = cell.Attribute(Office + "boolean-value")?.Value == "true";
                break;
            case "date":
                if (DateTime.TryParse(cell.Attribute(Office + "date-value")?.Value, Inv, DateTimeStyles.None, out var date)) { value = date; format ??= date.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm"; }
                break;
            case "time":
                var tm = cell.Attribute(Office + "time-value")?.Value ?? "";
                var tmatch = Regex.Match(tm, @"PT(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?");
                if (tmatch.Success) { value = new DateTime(1899, 12, 30).AddHours(Int(tmatch.Groups[1].Value)).AddMinutes(Int(tmatch.Groups[2].Value)).AddSeconds(double.TryParse(tmatch.Groups[3].Value, NumberStyles.Float, Inv, out var sec) ? sec : 0); format ??= "hh:mm:ss"; }
                break;
            default:
                var paragraphs = cell.Elements(TextNs + "p").Select(p => Writers.Plain(walker.Inline(p, new TextStyle(), null, out _))).ToList();
                if (paragraphs.Count > 0) value = string.Join("\n", paragraphs);
                break;
        }
        if (value is null && formula is null && cs.IsPlain) return null;
        return new CellModel
        {
            Value = value, Formula = formula is { Length: > 0 } ? Formula(formula) : null, Format = format,
            Bold = cs.Text.Bold == true, Italic = cs.Text.Italic == true, Underline = cs.Text.Underline == true, Strike = cs.Text.Strike == true,
            Color = cs.Text.Color, Fill = cs.Fill, Font = cs.Text.Font, SizePt = cs.Text.SizePt, Align = cs.Align, VAlign = cs.VAlign, Wrap = cs.Wrap,
        };
    }

    /// <summary>of:=SUM([.A1:.B2];[Sheet2.C3]) → SUM(A1:B2,Sheet2!C3).</summary>
    static string Formula(string f)
    {
        var s = f;
        var colon = s.IndexOf(":=", StringComparison.Ordinal);
        if (colon >= 0 && colon < 6) s = s[(colon + 2)..];
        else if (s.StartsWith('=')) s = s[1..];
        s = Regex.Replace(s, @"\[([^\]]*)\]", m =>
        {
            var r = m.Groups[1].Value;
            return string.Join(":", r.Split(':').Select(part =>
            {
                var dot = part.LastIndexOf('.');
                if (dot < 0) return part;
                var sheet = part[..dot].TrimStart('$');
                var cell = part[(dot + 1)..];
                if (sheet.Length == 0) return cell;
                if (sheet.StartsWith('\'')) return sheet + "!" + cell;
                return (Regex.IsMatch(sheet, @"^[A-Za-z0-9_一-鿿]+$") ? sheet : "'" + sheet + "'") + "!" + cell;
            }));
        });
        // ODF separates arguments with ; (outside strings)
        var sb = new StringBuilder();
        var quoted = false;
        foreach (var ch in s)
        {
            if (ch == '"') quoted = !quoted;
            sb.Append(ch == ';' && !quoted ? ',' : ch);
        }
        return sb.ToString();
    }

    // ---------- odp ----------

    public static DeckModel Deck(ZipArchive zip, List<string> warnings)
    {
        var content = Xml(zip, "content.xml");
        var stylesXml = TryXml(zip, "styles.xml");
        var styles = new Styles(content, stylesXml);
        var body = content.Root?.Element(Office + "body")?.Element(Office + "presentation") ?? throw Damaged();
        var deck = new DeckModel { WidthCm = 28, HeightCm = 21 };
        var walker = new TextWalker(styles, zip, warnings);
        var pages = body.Elements(Draw + "page").ToList();
        if (pages.Count > 0 && stylesXml is not null && PageSize(stylesXml, pages[0].Attribute(Draw + "master-page-name")?.Value) is { } size) (deck.WidthCm, deck.HeightCm) = size;
        foreach (var page in pages)
        {
            var slide = new SlideModel();
            foreach (var shape in page.Elements()) Shapes(shape, slide, styles, walker, warnings);
            if (page.Element(Presentation + "notes") is { } notes)
            {
                var text = string.Join("\n", notes.Descendants(TextNs + "p").Select(p => Writers.Plain(walker.Inline(p, new TextStyle(), null, out _))));
                if (text.Trim().Length > 0) slide.Notes = text;
            }
            deck.Slides.Add(slide);
        }
        return deck;
    }

    static (double W, double H)? PageSize(XDocument stylesXml, string? master)
    {
        var masters = stylesXml.Root?.Element(Office + "master-styles")?.Elements(Style + "master-page").ToList() ?? [];
        var mp = masters.FirstOrDefault(m => m.Attribute(Style + "name")?.Value == master) ?? masters.FirstOrDefault();
        var layoutName = mp?.Attribute(Style + "page-layout-name")?.Value;
        var layout = stylesXml.Root?.Element(Office + "automatic-styles")?.Elements(Style + "page-layout").FirstOrDefault(l => l.Attribute(Style + "name")?.Value == layoutName)?.Element(Style + "page-layout-properties");
        var w = Length(layout?.Attribute(Fo + "page-width")?.Value);
        var h = Length(layout?.Attribute(Fo + "page-height")?.Value);
        return w is > 0 && h is > 0 ? (w.Value, h.Value) : null;
    }

    static void Shapes(XElement e, SlideModel slide, Styles styles, TextWalker walker, List<string> warnings)
    {
        var n = e.Name;
        if (n == Draw + "g") { foreach (var child in e.Elements()) Shapes(child, slide, styles, walker, warnings); return; }
        if (n == Presentation + "notes" || n == Office + "forms" || n == Draw + "line" || n == Draw + "connector") return;
        var box = new ShapeModel
        {
            X = Length(e.Attribute(Svg + "x")?.Value) ?? 0, Y = Length(e.Attribute(Svg + "y")?.Value) ?? 0,
            W = Length(e.Attribute(Svg + "width")?.Value) ?? 0, H = Length(e.Attribute(Svg + "height")?.Value) ?? 0,
        };
        var cls = e.Attribute(Presentation + "class")?.Value;
        var gs = styles.Graphic(e.Attribute(Draw + "style-name")?.Value ?? e.Attribute(Presentation + "style-name")?.Value);
        box.Fill = gs.Fill; box.Line = gs.Line;
        if (n == Draw + "frame")
        {
            if (e.Element(Draw + "image") is not null)
            {
                if (walker.Picture(e) is not { Kind: BlockKind.Image } pic) return;
                box.Kind = ShapeKind.Image; box.Image = pic.Image;
                slide.Shapes.Add(box);
                return;
            }
            if (e.Element(Table + "table") is { } table)
            {
                var rows = new List<List<TableCell>>();
                foreach (var row in table.Descendants(Table + "table-row"))
                    rows.Add(row.Elements(Table + "table-cell").Select(c => new TableCell(string.Join("<br>", c.Elements(TextNs + "p").Select(p => walker.Inline(p, new TextStyle(), null, out _))))).ToList());
                if (rows.Count > 0) { box.Kind = ShapeKind.Table; box.Rows = rows; slide.Shapes.Add(box); }
                return;
            }
            var textBox = e.Element(Draw + "text-box");
            if (textBox is null) return;
            box.Kind = ShapeKind.Text;
            box.IsTitle = cls is "title";
            box.Paragraphs = Paragraphs(textBox, styles, walker, out var size);
            box.SizePt = size ?? (cls is "title" ? 40 : null);
            box.Geometry = "textbox";
            slide.Shapes.Add(box);
            return;
        }
        if (n == Draw + "custom-shape" || n == Draw + "rect" || n == Draw + "ellipse" || n == Draw + "circle")
        {
            box.Kind = ShapeKind.Text;
            box.Geometry = n == Draw + "rect" ? "rect" : n == Draw + "ellipse" || n == Draw + "circle" ? "ellipse"
                : (e.Element(Draw + "enhanced-geometry")?.Attribute(Draw + "type")?.Value) switch
                {
                    "rectangle" => "rect", "round-rectangle" => "roundRect", "ellipse" => "ellipse", "isosceles-triangle" => "triangle",
                    "diamond" => "diamond", "right-arrow" => "rightArrow", _ => "rect",
                };
            box.Paragraphs = Paragraphs(e, styles, walker, out var size);
            box.SizePt = size;
            if (box.Paragraphs.Count == 0 && box.Fill is null && box.Line is null) return;
            slide.Shapes.Add(box);
        }
    }

    static List<Block> Paragraphs(XElement container, Styles styles, TextWalker walker, out double? size)
    {
        var w = new TextWalker(styles, walker.ZipOf, walker.WarningsOf);
        w.Blocks(container, 0, null);
        var paragraphs = w.Out.Where(b => b.Kind is BlockKind.Paragraph or BlockKind.Heading).ToList();
        size = null;
        var first = container.Descendants(TextNs + "p").FirstOrDefault();
        if (first is not null)
        {
            var ps = styles.Para(first.Attribute(TextNs + "style-name")?.Value);
            size = ps.Text.SizePt ?? styles.Text(first.Elements(TextNs + "span").FirstOrDefault()?.Attribute(TextNs + "style-name")?.Value, ps.Text).SizePt;
        }
        return paragraphs;
    }

    // ---------- styles ----------

    public sealed record TextStyle
    {
        public bool? Bold, Italic, Underline, Strike;
        public string? Color, Highlight, Font;
        public double? SizePt;
    }

    public sealed record ParaStyle
    {
        public TextStyle Text = new();
        public string? Align, Style;
        public bool PageBreakBefore;
        public int Heading;
    }

    public sealed record CellStyle
    {
        public TextStyle Text = new();
        public string? Fill, Align, VAlign, Format;
        public bool Wrap;
        public bool IsPlain => Fill is null && Text.Bold != true && Text.Italic != true && Text.Color is null;
    }

    public sealed record GraphicStyle { public string? Fill, Line; }

    sealed class Styles
    {
        readonly Dictionary<string, XElement> _byKey = new(StringComparer.Ordinal);
        readonly Dictionary<string, string> _fontFamily = [];
        readonly Dictionary<string, XElement> _dataStyles = [];
        readonly Dictionary<string, XElement> _listStyles = [];

        public Styles(XDocument content, XDocument? styles)
        {
            foreach (var doc in new[] { styles, content })
            {
                if (doc?.Root is null) continue;
                foreach (var section in new[] { doc.Root.Element(Office + "styles"), doc.Root.Element(Office + "automatic-styles") })
                {
                    if (section is null) continue;
                    foreach (var s in section.Elements(Style + "style"))
                    {
                        var name = s.Attribute(Style + "name")?.Value;
                        var family = s.Attribute(Style + "family")?.Value ?? "";
                        if (name is not null) _byKey[family + "/" + name] = s;
                    }
                    foreach (var d in section.Elements().Where(x => x.Name.Namespace == Number))
                        if (d.Attribute(Style + "name")?.Value is { } dn) _dataStyles[dn] = d;
                    foreach (var l in section.Elements(TextNs + "list-style"))
                        if (l.Attribute(Style + "name")?.Value is { } ln) _listStyles[ln] = l;
                }
                foreach (var f in doc.Root.Element(Office + "font-face-decls")?.Elements(Style + "font-face") ?? [])
                    if (f.Attribute(Style + "name")?.Value is { } fn && f.Attribute(Svg + "font-family")?.Value is { } fam) _fontFamily[fn] = fam.Trim('\'', '"');
            }
        }

        /// <summary>The style and its ancestors, root first.</summary>
        IEnumerable<XElement> Chain(string family, string? name)
        {
            var chain = new List<XElement>();
            for (var n = name; n is not null && chain.Count < 20 && _byKey.TryGetValue(family + "/" + n, out var s); n = s.Attribute(Style + "parent-style-name")?.Value)
            {
                if (chain.Contains(s)) break;
                chain.Add(s);
            }
            chain.Reverse();
            return chain;
        }

        static string? DisplayName(XElement s) => s.Attribute(Style + "display-name")?.Value ?? s.Attribute(Style + "name")?.Value;

        public TextStyle Text(string? name, TextStyle inherited)
        {
            var t = inherited;
            foreach (var s in Chain("text", name)) t = ApplyText(t, s.Element(Style + "text-properties"));
            return t;
        }

        TextStyle ApplyText(TextStyle t, XElement? tp)
        {
            if (tp is null) return t;
            if (tp.Attribute(Fo + "font-weight")?.Value is { } fw) t = t with { Bold = fw is "bold" or "bolder" || (int.TryParse(fw, out var w) && w >= 600) };
            if (tp.Attribute(Fo + "font-style")?.Value is { } fs) t = t with { Italic = fs is "italic" or "oblique" };
            if (tp.Attribute(Style + "text-underline-style")?.Value is { } us) t = t with { Underline = us != "none" };
            if (tp.Attribute(Style + "text-line-through-style")?.Value is { } ls) t = t with { Strike = ls != "none" };
            if (tp.Attribute(Fo + "color")?.Value is { } c && Hex(c) is { } hc) t = t with { Color = hc };
            if (tp.Attribute(Fo + "background-color")?.Value is { } bg) t = t with { Highlight = Hex(bg) };
            if (tp.Attribute(Fo + "font-size")?.Value is { } size && size.EndsWith("pt") && double.TryParse(size[..^2], NumberStyles.Float, Inv, out var pt)) t = t with { SizePt = pt };
            if (tp.Attribute(Style + "font-name")?.Value is { } fn) t = t with { Font = _fontFamily.GetValueOrDefault(fn, fn) };
            else if (tp.Attribute(Fo + "font-family")?.Value is { } ff) t = t with { Font = ff.Trim('\'', '"') };
            return t;
        }

        public ParaStyle Para(string? name)
        {
            var p = new ParaStyle();
            foreach (var s in Chain("paragraph", name))
            {
                p.Text = ApplyText(p.Text, s.Element(Style + "text-properties"));
                var pp = s.Element(Style + "paragraph-properties");
                if (pp?.Attribute(Fo + "text-align")?.Value is { } al) p.Align = al switch { "center" => "center", "end" or "right" => "right", "justify" => "justify", "start" or "left" => "left", _ => p.Align };
                if (pp?.Attribute(Fo + "break-before")?.Value == "page") p.PageBreakBefore = true;
                var display = DisplayName(s) ?? "";
                if (Regex.Match(display, @"^(?:Heading|标题)[ _]?(\d)$", RegexOptions.IgnoreCase) is { Success: true } m) p.Heading = int.Parse(m.Groups[1].Value);
                else if (display is "Title" or "标题") p.Style = "Title";
                else if (display is "Subtitle" or "副标题") p.Style = "Subtitle";
                else if (display is "Quotations" or "Quote") p.Style = "Quote";
            }
            return p;
        }

        public CellStyle Cell(string? name)
        {
            var c = new CellStyle();
            foreach (var s in Chain("table-cell", name))
            {
                c.Text = ApplyText(c.Text, s.Element(Style + "text-properties"));
                var cp = s.Element(Style + "table-cell-properties");
                if (cp?.Attribute(Fo + "background-color")?.Value is { } bg) c.Fill = Hex(bg);
                if (cp?.Attribute(Fo + "wrap-option")?.Value == "wrap") c.Wrap = true;
                if (cp?.Attribute(Style + "vertical-align")?.Value is { } va) c.VAlign = va switch { "top" => "top", "middle" => "middle", "bottom" => "bottom", _ => c.VAlign };
                if (s.Element(Style + "paragraph-properties")?.Attribute(Fo + "text-align")?.Value is { } al) c.Align = al switch { "center" => "center", "end" or "right" => "right", "start" or "left" => "left", _ => c.Align };
                if (s.Attribute(Style + "data-style-name")?.Value is { } ds && _dataStyles.TryGetValue(ds, out var d)) c.Format = FormatCode(d);
            }
            return c;
        }

        public GraphicStyle Graphic(string? name)
        {
            var g = new GraphicStyle();
            foreach (var s in Chain("graphic", name).Concat(Chain("presentation", name)))
            {
                var gp = s.Element(Style + "graphic-properties");
                if (gp is null) continue;
                if (gp.Attribute(Draw + "fill")?.Value is { } fill) g.Fill = fill == "solid" ? Hex(gp.Attribute(Draw + "fill-color")?.Value) : null;
                else if (gp.Attribute(Draw + "fill-color")?.Value is { } fc && g.Fill is not null) g.Fill = Hex(fc);
                if (gp.Attribute(Draw + "stroke")?.Value is { } stroke) g.Line = stroke == "none" ? null : Hex(gp.Attribute(Svg + "stroke-color")?.Value) ?? "000000";
                else if (gp.Attribute(Svg + "stroke-color")?.Value is { } sc && g.Line is not null) g.Line = Hex(sc);
            }
            return g;
        }

        public string? ListType(string? name, int depth)
        {
            if (name is null || !_listStyles.TryGetValue(name, out var l)) return null;
            var level = l.Elements().FirstOrDefault(x => x.Attribute(TextNs + "level")?.Value == (depth + 1).ToString(Inv)) ?? l.Elements().FirstOrDefault();
            if (level is null) return null;
            return level.Name == TextNs + "list-level-style-number" ? "number" : "bullet";
        }

        public double? ColumnWidthCm(string? name) => Length(Chain("table-column", name).LastOrDefault()?.Element(Style + "table-column-properties")?.Attribute(Style + "column-width")?.Value);

        public double? RowHeightPt(string? name)
        {
            var cm = Length(Chain("table-row", name).LastOrDefault()?.Element(Style + "table-row-properties")?.Attribute(Style + "row-height")?.Value);
            return cm is > 0 ? cm * 72 / 2.54 : null;
        }

        static string FormatCode(XElement d)
        {
            var kind = d.Name.LocalName;
            if (kind == "date-style")
            {
                var sb = new StringBuilder();
                foreach (var part in d.Elements())
                {
                    var longForm = part.Attribute(Number + "style")?.Value == "long";
                    sb.Append(part.Name.LocalName switch
                    {
                        "year" => longForm ? "yyyy" : "yy", "month" => longForm ? "mm" : "m", "day" => longForm ? "dd" : "d", "hours" => "hh", "minutes" => "mm", "seconds" => "ss", "text" => part.Value, _ => "",
                    });
                }
                return sb.Length > 0 ? sb.ToString() : "yyyy-mm-dd";
            }
            var num = d.Elements().FirstOrDefault(x => x.Name.LocalName is "number" or "scientific-number" or "fraction");
            var decimals = Int(num?.Attribute(Number + "decimal-places")?.Value);
            var grouping = num?.Attribute(Number + "grouping")?.Value == "true";
            var code = (grouping ? "#,##0" : "0") + (decimals > 0 ? "." + new string('0', decimals) : "");
            return kind switch { "percentage-style" => code + "%", "currency-style" => code, "number-style" => code, "time-style" => "hh:mm:ss", _ => code };
        }
    }

    // ---------- helpers ----------

    static int Int(string? s, int fallback = 0) => int.TryParse(s, NumberStyles.Integer, Inv, out var v) ? v : fallback;

    /// <summary>A length with its unit → cm.</summary>
    public static double? Length(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        var m = Regex.Match(s.Trim(), @"^(-?[\d.]+)\s*(cm|mm|in|pt|px|pc)?$");
        if (!m.Success || !double.TryParse(m.Groups[1].Value, NumberStyles.Float, Inv, out var v)) return null;
        return m.Groups[2].Value switch { "mm" => v / 10, "in" => v * 2.54, "pt" => v * 2.54 / 72, "px" => v * 2.54 / 96, "pc" => v * 2.54 / 6, _ => v };
    }

    static string? Hex(string? s)
    {
        if (s is null) return null;
        s = s.Trim();
        if (s.StartsWith('#') && s.Length == 7) return s[1..].ToUpperInvariant();
        return null;
    }
}

file static class CellModelExtensions
{
    public static CellModel MemberwiseCloneOf(this CellModel c) => new()
    {
        Row = c.Row, Col = c.Col, Value = c.Value, Formula = c.Formula, Bold = c.Bold, Italic = c.Italic, Underline = c.Underline, Strike = c.Strike, Wrap = c.Wrap,
        Color = c.Color, Fill = c.Fill, Font = c.Font, SizePt = c.SizePt, Format = c.Format, Align = c.Align, VAlign = c.VAlign,
    };
}
