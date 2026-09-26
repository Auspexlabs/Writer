using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml;
using Writer.Core;
using Writer.Formats.Common;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

sealed class DocxTable(DocxDocument doc, W.Table table) : Node, IDocxContainer
{
    internal const int Twip = 635; // EMU per twentieth of a point
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    public override string Kind => "table";
    public override object Anchor => table;
    public OpenXmlElement Container => table;

    public IEnumerable<W.TableRow> Rows => table.Elements<W.TableRow>();

    protected override IEnumerable<Node> ProjectChildren() => Rows.Select(r => (Node)new DocxRow(doc, r));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var rows = Rows.ToList();
        var props = new Dictionary<string, string>
        {
            ["rows"] = rows.Count.ToString(Inv),
            ["cols"] = GridCols(table).ToString(Inv),
            ["data"] = NodeJson.Compact(w =>
            {
                w.WriteStartArray();
                foreach (var r in rows) DocxRow.WriteData(w, r);
                w.WriteEndArray();
            }),
        };
        var pr = table.GetFirstChild<W.TableProperties>();
        if (pr?.TableStyle?.Val?.Value is { } style) props["style"] = style;
        if (pr?.TableBorders is { } borders)
        {
            if (BordersOf(borders) is { } kind) props["borders"] = kind;
            if (BorderColorOf(borders) is { } color) props["borderColor"] = color;
        }
        if (pr?.TableWidth is { } tw && WidthOf(tw) is { } width) props["width"] = width;
        if (WidthsOf(table) is { } widths) props["widths"] = widths;
        if (AlignOf(pr?.TableJustification?.Val?.InnerText) is { } align) props["align"] = align;
        if (pr?.TableLook is { } look && FirstRowOf(look) is { } first) props["header"] = first ? "true" : "false";
        return props;
    }

    public override string GetRaw() => table.OuterXml;

    /// <summary>A new table from rows/cols or data, styled with Table Grid and spanning the page width.
    /// Style, borders, widths and alignment arrive through SetProp like on any other table.</summary>
    public static (OpenXmlElement Element, string[] Consumed) New(DocxDocument doc, IReadOnlyDictionary<string, string> props)
    {
        var data = props.TryGetValue("data", out var json) ? ParseRows(json) : null;
        var rows = props.TryGetValue("rows", out var r) ? int.Parse(r, Inv) : data?.Count ?? 0;
        var cols = props.TryGetValue("cols", out var c) ? int.Parse(c, Inv) : data?.Max(row => row.Count) ?? 0;
        if (data is not null)
        {
            rows = Math.Max(rows, data.Count);
            cols = Math.Max(cols, data.Count == 0 ? 0 : data.Max(row => row.Count));
        }
        if (rows < 1 || cols < 1)
            throw new WriterException(ErrorCode.Validation, "A table needs rows and cols, or data", "Example: --prop rows=3 --prop cols=4, or --prop data='[[\"a\",\"b\"]]'");
        var width = ContentWidthTwips(doc) / cols;
        var t = new W.Table(
            new W.TableProperties(
                new W.TableStyle { Val = doc.Styles.ResolveStyle("TableGrid", "table") },
                new W.TableWidth { Width = "5000", Type = W.TableWidthUnitValues.Pct },
                new W.TableLook { Val = "04A0" }),
            new W.TableGrid(Enumerable.Range(0, cols).Select(_ => (OpenXmlElement)new W.GridColumn { Width = width.ToString(Inv) })));
        for (var i = 0; i < rows; i++)
            t.Append(new W.TableRow(Enumerable.Range(0, cols).Select(_ => (OpenXmlElement)NewCell(width))));
        if (data is not null) Fill(doc, t, data);
        return (t, ["rows", "cols", "data"]);
    }

    /// <summary>Whether the style's first-row look applies (标题行): tblLook's firstRow, or its bit in the older val.</summary>
    static bool? FirstRowOf(W.TableLook look) =>
        look.FirstRow?.Value is { } first ? first
        : look.Val?.Value is { } hex && int.TryParse(hex, NumberStyles.HexNumber, Inv, out var bits) ? (bits & 0x20) != 0 : null;

    static void SetFirstRow(W.TableLook look, bool on)
    {
        look.FirstRow = on;
        if (look.Val?.Value is { } hex && int.TryParse(hex, NumberStyles.HexNumber, Inv, out var bits)) look.Val = (on ? bits | 0x20 : bits & ~0x20).ToString("X4", Inv);
    }

    internal static W.TableCell NewCell(int widthTwips) => new(
        new W.TableCellProperties(new W.TableCellWidth { Width = widthTwips.ToString(Inv), Type = W.TableWidthUnitValues.Dxa }),
        new W.Paragraph());

    internal static int ContentWidthTwips(DocxDocument doc)
    {
        var section = doc.Main.Document!.Body!.GetFirstChild<W.SectionProperties>();
        var page = section?.GetFirstChild<W.PageSize>()?.Width?.Value;
        var margin = section?.GetFirstChild<W.PageMargin>();
        if (page is null || margin is null) return 9026;
        return (int)page.Value - (int)(margin.Left?.Value ?? 1440) - (int)(margin.Right?.Value ?? 1440);
    }

    internal static List<List<string>> ParseRows(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Array || parsed.RootElement.EnumerateArray().Any(r => r.ValueKind != JsonValueKind.Array))
            throw new WriterException(ErrorCode.Validation, "data must be a JSON array of rows", "Example: [[\"Name\",\"Score\"],[\"Ann\",\"90\"]]");
        return parsed.RootElement.EnumerateArray().Select(row => row.EnumerateArray().Select(CellText).ToList()).ToList();
    }

    internal static List<string> ParseCells(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Array)
            throw new WriterException(ErrorCode.Validation, "data must be a JSON array", "Example: [\"Ann\",\"90\"]");
        return parsed.RootElement.EnumerateArray().Select(CellText).ToList();
    }

    static string CellText(JsonElement e) => e.ValueKind switch
    {
        JsonValueKind.String => e.GetString()!,
        JsonValueKind.Null => "",
        _ => e.GetRawText(),
    };

    public override void SetProp(string name, string value)
    {
        var pr = table.GetFirstChild<W.TableProperties>() ?? table.PrependChild(new W.TableProperties());
        switch (name)
        {
            case "rows": SetRowCount(doc, table, int.Parse(value, Inv)); break;
            case "cols": SetColumnCount(table, int.Parse(value, Inv)); break;
            case "data": Fill(doc, table, ParseRows(value)); break;
            case "style": pr.TableStyle = value.Length == 0 ? null : new W.TableStyle { Val = doc.Styles.ResolveStyle(value, "table") }; break;
            case "borders":
                pr.TableBorders = value == "style" ? null : MakeBorders<W.TableBorders>(value, pr.TableBorders is { } old ? BorderColorOf(old) : null);
                break;
            case "borderColor": Recolor(pr.TableBorders ??= StyleBorders(doc, pr) ?? MakeBorders<W.TableBorders>("all", null), value); break;
            case "width": pr.TableWidth = ParseWidth(value); break;
            case "header": SetFirstRow(pr.TableLook ??= new W.TableLook { Val = "04A0" }, value == "true"); break;
            case "widths": SetWidths(value); break;
            case "align":
                pr.TableJustification = new W.TableJustification
                {
                    Val = value switch
                    {
                        "center" => W.TableRowAlignmentValues.Center,
                        "right" => W.TableRowAlignmentValues.Right,
                        _ => W.TableRowAlignmentValues.Left,
                    },
                };
                break;
        }
    }

    static void Fill(DocxDocument doc, W.Table t, List<List<string>> data)
    {
        SetRowCount(doc, t, Math.Max(1, data.Count));
        var rows = t.Elements<W.TableRow>().ToList();
        for (var i = 0; i < data.Count; i++)
        {
            DocxRow.SetCellCount(doc, rows[i], Math.Max(1, data[i].Count));
            var cells = Visible(rows[i]).ToList();
            for (var j = 0; j < data[i].Count; j++) DocxCell.SetText(doc, cells[j], data[i][j]);
        }
        SyncGrid(t);
    }

    static void SetRowCount(DocxDocument doc, W.Table t, int count)
    {
        var rows = t.Elements<W.TableRow>().ToList();
        while (rows.Count < count)
        {
            var clone = CloneRow(doc, rows[^1]);
            t.Append(clone);
            rows.Add(clone);
        }
        while (rows.Count > count)
        {
            rows[^1].Remove();
            rows.RemoveAt(rows.Count - 1);
        }
    }

    /// <summary>A copy of the row with empty cells; vertical merges are not continued into it.</summary>
    internal static W.TableRow CloneRow(DocxDocument doc, W.TableRow row)
    {
        var clone = (W.TableRow)row.CloneNode(true);
        if (clone.TableRowProperties is { } rowPr)
        {
            rowPr.GetFirstChild<W.TableHeader>()?.Remove(); // a new row is a body row, even when added under the header row
            if (!rowPr.HasChildren) rowPr.Remove();
        }
        foreach (var cell in clone.Elements<W.TableCell>())
        {
            if (cell.TableCellProperties is { } pr)
            {
                pr.VerticalMerge = null;
                if (!pr.HasChildren) pr.Remove();
            }
            DocxCell.SetText(doc, cell, "");
        }
        return clone;
    }

    static void SetColumnCount(W.Table t, int count)
    {
        foreach (var row in t.Elements<W.TableRow>())
        {
            while (GridWidth(row) > count)
            {
                var last = row.Elements<W.TableCell>().Last();
                Release(last);
                last.Remove();
            }
            while (GridWidth(row) < count) row.Append(NewCell(GridColWidth(t, GridWidth(row)) ?? 2000));
        }
        SyncGrid(t);
    }

    /// <summary>Makes tblGrid match the widest row, counting spanned columns.</summary>
    internal static void SyncGrid(W.Table t)
    {
        var wanted = t.Elements<W.TableRow>().Select(GridWidth).DefaultIfEmpty(0).Max();
        var grid = t.GetFirstChild<W.TableGrid>();
        if (grid is null)
        {
            grid = new W.TableGrid();
            t.InsertAfter(grid, t.GetFirstChild<W.TableProperties>());
        }
        var columns = grid.Elements<W.GridColumn>().ToList();
        if (columns.Count == wanted) return;
        while (columns.Count < wanted)
        {
            var clone = (W.GridColumn?)columns.LastOrDefault()?.CloneNode(true) ?? new W.GridColumn { Width = "2000" };
            grid.Append(clone);
            columns.Add(clone);
        }
        while (columns.Count > wanted)
        {
            columns[^1].Remove();
            columns.RemoveAt(columns.Count - 1);
        }
        // a column came or went somewhere in the middle: each column takes the width of a one-column cell in it again
        var set = new bool[columns.Count];
        foreach (var cell in t.Elements<W.TableRow>().SelectMany(r => r.Elements<W.TableCell>()))
            if (Span(cell) == 1 && GridStart(cell) is var i && i < columns.Count && !set[i] && cell.TableCellProperties?.TableCellWidth is { } tw
                && tw.Type?.Value == W.TableWidthUnitValues.Dxa && int.TryParse(tw.Width?.Value, NumberStyles.Integer, Inv, out var twips) && twips > 0)
            {
                columns[i].Width = twips.ToString(Inv);
                set[i] = true;
            }
    }

    // --- grid and merges -------------------------------------------------------------------------------------

    internal static int Span(W.TableCell c) => c.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
    static bool IsRestart(W.TableCell c) => c.TableCellProperties?.VerticalMerge?.Val?.Value == W.MergedCellValues.Restart;
    /// <summary>A cell merged into the one above it; hidden from the tree.</summary>
    internal static bool IsContinue(W.TableCell c) => c.TableCellProperties?.VerticalMerge is { } vm && vm.Val?.Value != W.MergedCellValues.Restart;
    internal static IEnumerable<W.TableCell> Visible(W.TableRow row) => row.Elements<W.TableCell>().Where(c => !IsContinue(c));
    internal static int GridWidth(W.TableRow row) => row.Elements<W.TableCell>().Sum(Span);
    static int GridStart(W.TableCell c) => c.ElementsBefore().OfType<W.TableCell>().Sum(Span);
    static IEnumerable<W.TableRow> RowsBelow(W.TableCell c) => c.Parent!.ElementsAfter().OfType<W.TableRow>();
    static W.Table TableOf(W.TableCell c) => (W.Table)c.Parent!.Parent!;

    static int GridCols(W.Table t)
    {
        var grid = t.GetFirstChild<W.TableGrid>()?.Elements<W.GridColumn>().Count() ?? 0;
        return grid > 0 ? grid : t.Elements<W.TableRow>().Select(GridWidth).DefaultIfEmpty(0).Max();
    }

    static int? GridColWidth(W.Table t, int index) =>
        t.GetFirstChild<W.TableGrid>()?.Elements<W.GridColumn>().ElementAtOrDefault(index)?.Width?.Value is { } w
        && int.TryParse(w, NumberStyles.Integer, Inv, out var twips) ? twips : null;

    static W.TableCell? CellStartingAt(W.TableRow row, int col)
    {
        var start = 0;
        foreach (var c in row.Elements<W.TableCell>())
        {
            if (start == col) return c;
            if (start > col) return null;
            start += Span(c);
        }
        return null;
    }

    internal static int RowSpan(W.TableCell c)
    {
        if (!IsRestart(c)) return 1;
        var col = GridStart(c);
        var n = 1;
        foreach (var row in RowsBelow(c))
        {
            if (CellStartingAt(row, col) is not { } below || !IsContinue(below)) break;
            n++;
        }
        return n;
    }

    /// <summary>The cell without spans, for clones that become new cells.</summary>
    internal static W.TableCell Plain(W.TableCell c)
    {
        if (c.TableCellProperties is { } pr)
        {
            pr.GridSpan = null;
            pr.VerticalMerge = null;
            if (!pr.HasChildren) pr.Remove();
        }
        return c;
    }

    /// <summary>Ends a vertical merge before the cell goes away, so the cells below become ordinary cells again.</summary>
    internal static void Release(W.TableCell c)
    {
        if (IsRestart(c)) SetRowspan(c, 1);
    }

    /// <summary>Resizes the cell and, when it is vertically merged, the hidden cells below it; everything absorbed lands in the visible cell.</summary>
    internal static void SetColspan(W.TableCell cell, int n)
    {
        var col = GridStart(cell);
        var group = new List<W.TableCell> { cell };
        foreach (var row in RowsBelow(cell).Take(RowSpan(cell) - 1))
            if (CellStartingAt(row, col) is { } below) group.Add(below);
        CheckWiden(cell, group, n);
        foreach (var c in group) Resize(c, n, cell);
        SyncGrid(TableOf(cell));
    }

    /// <summary>Widens a cell over its right neighbours (their content moves into <paramref name="into"/>), or narrows it and puts empty cells in the freed columns.</summary>
    static void Resize(W.TableCell cell, int n, W.TableCell into)
    {
        var table = TableOf(cell);
        var span = Span(cell);
        while (span < n && cell.NextSibling<W.TableCell>() is { } next)
        {
            Release(next);
            Absorb(into, next);
            span += Span(next);
            next.Remove();
        }
        if (span < n) span = n; // ponytail: no neighbour left, the row simply grows; SyncGrid widens the grid
        OpenXmlElement after = cell;
        for (var i = 0; span - i > n; i++)
            after = cell.Parent!.InsertAfter(SplitOff(cell, GridColWidth(table, GridStart(cell) + n + i) ?? 2000), after);
        var pr = cell.TableCellProperties ??= new W.TableCellProperties();
        pr.GridSpan = n > 1 ? new W.GridSpan { Val = n } : null;
        FitWidth(cell);
    }

    /// <summary>An empty one-column cell with the formatting of the cell it is split from (shading, borders, alignment), as Word's Split Cells.</summary>
    static W.TableCell SplitOff(W.TableCell from, int widthTwips)
    {
        var cell = NewCell(widthTwips);
        if (from.TableCellProperties is { } own)
        {
            var pr = (W.TableCellProperties)own.CloneNode(true);
            pr.GridSpan = null;
            pr.VerticalMerge = null;
            pr.TableCellWidth = new W.TableCellWidth { Width = widthTwips.ToString(Inv), Type = W.TableWidthUnitValues.Dxa };
            cell.TableCellProperties = pr;
        }
        return cell;
    }

    /// <summary>Widening must not take a cell out of a vertical merge that starts above the widened cell: the rest of that merge
    /// would be left without its first cell. Word refuses such non-rectangular merges too.</summary>
    static void CheckWiden(W.TableCell cell, IReadOnlyList<W.TableCell> group, int n)
    {
        var rows = TableOf(cell).Elements<W.TableRow>().ToList();
        var top = rows.IndexOf((W.TableRow)cell.Parent!);
        foreach (var member in group)
        {
            var span = Span(member);
            for (var next = member.NextSibling<W.TableCell>(); span < n && next is not null; next = next.NextSibling<W.TableCell>())
            {
                if (IsContinue(next) && rows.IndexOf((W.TableRow)Owner(next).Parent!) < top)
                    throw new WriterException(ErrorCode.Validation, "The cells to the right belong to a merged cell that starts in a row above",
                        "Merge only cells that cover the same rows, or set rowspan=1 on that merged cell first.");
                span += Span(next);
            }
        }
    }

    /// <summary>The visible cell a hidden (vertically merged) cell belongs to.</summary>
    static W.TableCell Owner(W.TableCell c)
    {
        var col = GridStart(c);
        for (var row = (W.TableRow)c.Parent!; IsContinue(c) && row.PreviousSibling<W.TableRow>() is { } up && CellStartingAt(up, col) is { } above; row = up) c = above;
        return c;
    }

    internal static void SetRowspan(W.TableCell cell, int n)
    {
        var current = RowSpan(cell);
        var col = GridStart(cell);
        var below = RowsBelow(cell).ToList();
        if (n - 1 > below.Count)
            throw new WriterException(ErrorCode.Validation, $"rowspan {n} needs {n - 1} rows below this one; there are {below.Count}", "Add rows to the table first.");
        var table = TableOf(cell);
        for (var i = 1; i < n; i++)
        {
            var row = below[i - 1];
            var target = CellStartingAt(row, col);
            if (target is null)
            {
                if (GridWidth(row) > col)
                    throw new WriterException(ErrorCode.Validation, "The cell below is part of a horizontal merge", "Set colspan=1 on it first.");
                while (GridWidth(row) < col) row.Append(NewCell(GridColWidth(table, GridWidth(row)) ?? 2000));
                target = row.AppendChild(NewCell(GridColWidth(table, col) ?? 2000));
                if (Span(cell) > 1) (target.TableCellProperties ??= new W.TableCellProperties()).GridSpan = new W.GridSpan { Val = Span(cell) };
            }
            if (i < current) continue;
            Release(target); // a merge of its own ends here; the cells it covered become ordinary cells
            if (Span(target) > Span(cell))
                throw new WriterException(ErrorCode.Validation, $"The cell below spans {Span(target)} columns, this one {Span(cell)}", $"Set colspan={Span(target)} on this cell first.");
            if (Span(target) < Span(cell)) SetColspan(target, Span(cell)); // the merged region is a rectangle, as in Word
            Absorb(cell, target);
            foreach (var child in target.ChildElements.Where(e => e is not W.TableCellProperties).ToList()) child.Remove();
            target.Append(new W.Paragraph());
            var merged = cell.TableCellProperties is { } own ? (W.TableCellProperties)own.CloneNode(true) : new W.TableCellProperties();
            merged.VerticalMerge = new W.VerticalMerge(); // the merged cell looks the same in every row, as when Word merges cells
            target.TableCellProperties = merged;
        }
        for (var i = n; i < current; i++)
            if (CellStartingAt(below[i - 1], col)?.TableCellProperties is { } freed)
            {
                freed.VerticalMerge = null;
                if (!freed.HasChildren) freed.Remove();
            }
        var pr = cell.TableCellProperties ??= new W.TableCellProperties();
        pr.VerticalMerge = n > 1 ? new W.VerticalMerge { Val = W.MergedCellValues.Restart } : null;
        if (!pr.HasChildren) pr.Remove();
        SyncGrid(table);
    }

    /// <summary>Removes a visible cell together with the hidden cells of its vertical merge: the whole merged cell, as the tree shows it.
    /// Rows keep their other cells in their grid columns, so removing a column cell by cell leaves every merge intact.</summary>
    internal static void RemoveMerged(W.TableCell cell)
    {
        var col = GridStart(cell);
        var hidden = RowsBelow(cell).Take(RowSpan(cell) - 1).Select(r => CellStartingAt(r, col)!).ToList();
        if (hidden.Any(c => c.Parent!.Elements<W.TableCell>().Count() == 1))
            throw new WriterException(ErrorCode.Validation, "A row below would lose its last cell", "Remove those rows instead.");
        foreach (var c in hidden) c.Remove();
        cell.Remove();
    }

    /// <summary>Before a row with vertically merged cells goes away: the cell below each merge takes over its content and formatting, as in Word.</summary>
    internal static void Promote(W.TableRow row)
    {
        foreach (var cell in row.Elements<W.TableCell>().Where(c => IsRestart(c) && RowSpan(c) > 1).ToList())
        {
            var below = CellStartingAt(RowsBelow(cell).First(), GridStart(cell))!;
            var pr = (W.TableCellProperties)cell.TableCellProperties!.CloneNode(true);
            pr.VerticalMerge = RowSpan(cell) > 2 ? new W.VerticalMerge { Val = W.MergedCellValues.Restart } : null;
            below.TableCellProperties = pr;
            Absorb(below, cell);
        }
    }

    /// <summary>A row like <paramref name="template"/> for insertion above <paramref name="next"/>: vertical merges that run past the insertion point grow over the new row.</summary>
    internal static W.TableRow NewRow(DocxDocument doc, W.TableRow template, W.TableRow? above, W.TableRow? next)
    {
        var row = CloneRow(doc, template);
        if (above is null || next is null) return row;
        foreach (var cell in row.Elements<W.TableCell>())
        {
            var col = GridStart(cell);
            var up = CellStartingAt(above, col);
            var down = CellStartingAt(next, col);
            if (up is not null && down is not null && (IsRestart(up) || IsContinue(up)) && IsContinue(down))
                (cell.TableCellProperties ??= new W.TableCellProperties()).VerticalMerge = new W.VerticalMerge();
        }
        return row;
    }

    /// <summary>Moves the content of one cell to the end of another, dropping blank paragraphs.</summary>
    static void Absorb(W.TableCell into, W.TableCell from)
    {
        var blocks = from.ChildElements.Where(e => e is not W.TableCellProperties && !IsBlank(e)).ToList();
        if (blocks.Count == 0) return;
        var own = into.ChildElements.Where(e => e is not W.TableCellProperties).ToList();
        if (own.Count == 1 && IsBlank(own[0])) own[0].Remove();
        foreach (var block in blocks)
        {
            block.Remove();
            into.Append(block);
        }
        if (into.LastChild is not W.Paragraph) into.Append(new W.Paragraph());
    }

    static bool IsBlank(OpenXmlElement e) => e is W.Paragraph p && p.InnerText.Length == 0 && !p.Descendants<W.Drawing>().Any();

    /// <summary>tcW from the grid columns the cell covers, when the grid has widths.</summary>
    static void FitWidth(W.TableCell cell)
    {
        var start = GridStart(cell);
        var widths = Enumerable.Range(start, Span(cell)).Select(i => GridColWidth(TableOf(cell), i)).ToList();
        if (widths.Any(w => w is null)) return;
        (cell.TableCellProperties ??= new W.TableCellProperties()).TableCellWidth =
            new W.TableCellWidth { Width = widths.Sum(w => w!.Value).ToString(Inv), Type = W.TableWidthUnitValues.Dxa };
    }

    internal static void SetCellWidth(W.TableCell cell, int twips)
    {
        (cell.TableCellProperties ??= new W.TableCellProperties()).TableCellWidth =
            new W.TableCellWidth { Width = twips.ToString(Inv), Type = W.TableWidthUnitValues.Dxa };
        if (Span(cell) == 1 && TableOf(cell).GetFirstChild<W.TableGrid>()?.Elements<W.GridColumn>().ElementAtOrDefault(GridStart(cell)) is { } column)
            column.Width = twips.ToString(Inv);
    }

    // --- widths, borders, alignment ----------------------------------------------------------------------------

    internal static int TwipsOfEmu(string emu) => (int)Math.Round(long.Parse(emu, Inv) / (double)Twip);

    /// <summary>EMU of a twip count, snapped to the nearest 0.01cm when that is less than a twip away: twips cannot hold 5cm
    /// exactly (2835 twips is 5.0005cm), and a sum of columns set as 3cm and 5cm should read back as 8cm.</summary>
    internal static long EmuOfTwips(long twips)
    {
        var emu = twips * Twip;
        var snapped = (long)Math.Round(emu / 3600.0) * 3600;
        return Math.Abs(snapped - emu) < Twip ? snapped : emu;
    }

    static int TwipsOfLength(string length)
    {
        var emu = Units.ParseLength(length);
        if (emu <= 0) throw new WriterException(ErrorCode.Validation, $"'{length}' is not a positive length", "Use a length like 3cm.");
        return (int)Math.Round(emu / (double)Twip);
    }

    static string? WidthOf(W.TableWidth tw)
    {
        var w = tw.Width?.Value;
        if (tw.Type?.Value == W.TableWidthUnitValues.Pct && w is not null)
            return w.EndsWith('%') ? w : int.TryParse(w, NumberStyles.Integer, Inv, out var fiftieths) ? (fiftieths / 50.0).ToString("0.##", Inv) + "%" : null;
        if (tw.Type?.Value == W.TableWidthUnitValues.Dxa && int.TryParse(w, NumberStyles.Integer, Inv, out var twips) && twips > 0)
            return Units.FormatLength(EmuOfTwips(twips));
        return null;
    }

    static W.TableWidth ParseWidth(string value)
    {
        var v = value.Trim();
        if (v.Equals("auto", StringComparison.OrdinalIgnoreCase)) return new W.TableWidth { Width = "0", Type = W.TableWidthUnitValues.Auto };
        if (!v.EndsWith('%')) return new W.TableWidth { Width = TwipsOfLength(v).ToString(Inv), Type = W.TableWidthUnitValues.Dxa };
        if (!double.TryParse(v[..^1], NumberStyles.Float, Inv, out var pct) || pct <= 0 || pct > 100)
            throw new WriterException(ErrorCode.Validation, $"width: '{value}' is not a percentage", "Use 1% to 100%, auto, or a length like 12cm.");
        return new W.TableWidth { Width = ((int)Math.Round(pct * 50)).ToString(Inv), Type = W.TableWidthUnitValues.Pct };
    }

    static string? WidthsOf(W.Table t)
    {
        var columns = t.GetFirstChild<W.TableGrid>()?.Elements<W.GridColumn>().ToList();
        if (columns is null || columns.Count == 0 || columns.Any(c => !int.TryParse(c.Width?.Value, NumberStyles.Integer, Inv, out _))) return null;
        return NodeJson.Compact(w =>
        {
            w.WriteStartArray();
            foreach (var c in columns) w.WriteStringValue(Units.FormatLength(EmuOfTwips(long.Parse(c.Width!.Value!, Inv))));
            w.WriteEndArray();
        });
    }

    /// <summary>Column widths from the left: gridCol, every cell's tcW, the table width as their sum, fixed layout.</summary>
    void SetWidths(string json)
    {
        var widths = ParseCells(json).Select(TwipsOfLength).ToList();
        if (widths.Count == 0) throw new WriterException(ErrorCode.Validation, "widths needs at least one column", "Example: widths=[\"3cm\",\"5cm\"]");
        SyncGrid(table);
        var columns = table.GetFirstChild<W.TableGrid>()!.Elements<W.GridColumn>().ToList();
        for (var i = 0; i < columns.Count && i < widths.Count; i++) columns[i].Width = widths[i].ToString(Inv);
        foreach (var cell in table.Elements<W.TableRow>().SelectMany(r => r.Elements<W.TableCell>())) FitWidth(cell);
        var total = columns.Sum(c => int.Parse(c.Width!.Value!, Inv));
        var pr = table.GetFirstChild<W.TableProperties>()!;
        pr.TableWidth = new W.TableWidth { Width = total.ToString(Inv), Type = W.TableWidthUnitValues.Dxa };
        pr.TableLayout = new W.TableLayout { Type = W.TableLayoutValues.Fixed };
    }

    internal static string? AlignOf(string? jc) => jc switch
    {
        "left" or "start" => "left",
        "center" => "center",
        "right" or "end" => "right",
        _ => null,
    };

    /// <summary>none | all | outside | inside | horizontal from which sides are drawn; null for other mixes.
    /// horizontal is every horizontal line (top, bottom and between rows) and no vertical one.</summary>
    internal static string? BordersOf(OpenXmlElement borders)
    {
        var top = On<W.TopBorder>(borders);
        var bottom = On<W.BottomBorder>(borders);
        var sides = On<W.LeftBorder>(borders) || On<W.StartBorder>(borders);
        var right = On<W.RightBorder>(borders) || On<W.EndBorder>(borders);
        var h = On<W.InsideHorizontalBorder>(borders);
        var v = On<W.InsideVerticalBorder>(borders);
        return (top, sides, bottom, right, h, v) switch
        {
            (true, true, true, true, true, true) => "all",
            (true, true, true, true, false, false) => "outside",
            (false, false, false, false, true, true) => "inside",
            (true, false, true, false, true, false) => "horizontal",
            (false, false, false, false, false, false) => "none",
            _ => null,
        };
    }

    /// <summary>The sides drawn, for a cell's own mix (top, bottom, left, right in that order); null when none is.</summary>
    internal static string? SidesOf(OpenXmlElement borders)
    {
        var sides = new List<string>();
        if (On<W.TopBorder>(borders)) sides.Add("top");
        if (On<W.BottomBorder>(borders)) sides.Add("bottom");
        if (On<W.LeftBorder>(borders) || On<W.StartBorder>(borders)) sides.Add("left");
        if (On<W.RightBorder>(borders) || On<W.EndBorder>(borders)) sides.Add("right");
        return sides.Count == 0 ? null : string.Join(' ', sides);
    }

    static bool On<T>(OpenXmlElement borders) where T : W.BorderType =>
        borders.GetFirstChild<T>()?.Val?.Value is { } val && val != W.BorderValues.Nil && val != W.BorderValues.None;

    /// <summary>The borders the table's style (or a style it is based on) draws, as a copy to override; null when it draws none.</summary>
    static W.TableBorders? StyleBorders(DocxDocument doc, W.TableProperties pr)
    {
        var id = pr.TableStyle?.Val?.Value;
        for (var depth = 0; id is not null && depth < 10; depth++)
        {
            var style = doc.Styles.Find(id);
            if (style?.StyleTableProperties?.TableBorders is { } borders) return (W.TableBorders)borders.CloneNode(true);
            id = style?.BasedOn?.Val?.Value;
        }
        return null;
    }

    internal static string? BorderColorOf(OpenXmlElement borders) =>
        borders.Elements<W.BorderType>().Select(b => b.Color?.Value).FirstOrDefault(c => c is not null && !c.Equals("auto", StringComparison.OrdinalIgnoreCase))?.ToUpperInvariant();

    /// <summary>Borders for a kind (none, all, outside, inside, horizontal) or, for a cell, the sides to draw listed (top bottom left right).</summary>
    internal static T MakeBorders<T>(string kind, string? color) where T : OpenXmlCompositeElement, new()
    {
        var known = kind is "none" or "all" or "outside" or "inside" or "horizontal" or "box";
        var listed = known ? [] : kind.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries).Select(s => s.ToLowerInvariant()).ToHashSet();
        if (listed.Except(["top", "bottom", "left", "right"]).Any())
            throw new WriterException(ErrorCode.Validation, $"Unknown borders '{kind}'", "Use none, all, outside, inside, horizontal, or the sides to draw: top bottom left right.");
        var sides = kind is "all" or "outside" or "box";
        var topBottom = sides || kind == "horizontal";
        var insideH = kind is "all" or "inside" or "horizontal";
        var insideV = kind is "all" or "inside";
        var borders = new T();
        borders.Append(Border<W.TopBorder>(topBottom || listed.Contains("top"), color), Border<W.LeftBorder>(sides || listed.Contains("left"), color), Border<W.BottomBorder>(topBottom || listed.Contains("bottom"), color), Border<W.RightBorder>(sides || listed.Contains("right"), color),
            Border<W.InsideHorizontalBorder>(insideH, color), Border<W.InsideVerticalBorder>(insideV, color));
        return borders;
    }

    static T Border<T>(bool on, string? color) where T : W.BorderType, new() => on
        ? new T { Val = W.BorderValues.Single, Size = 4U, Space = 0U, Color = color ?? "auto" }
        : new T { Val = W.BorderValues.None, Size = 0U, Space = 0U, Color = "auto" };

    static void Recolor(OpenXmlElement borders, string color)
    {
        foreach (var b in borders.Elements<W.BorderType>())
            if (b.Val?.Value is { } val && val != W.BorderValues.Nil && val != W.BorderValues.None) b.Color = color == "none" ? "auto" : color;
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var rows = Rows.ToList();
        var at = index is { } i ? Math.Clamp(i, 1, rows.Count + 1) : rows.Count + 1; // 1-based position of the new row
        var above = at > 1 ? rows[at - 2] : null;
        var row = rows.Count == 0 ? new W.TableRow(NewCell(2000)) : NewRow(doc, above ?? rows[0], above, at <= rows.Count ? rows[at - 1] : null);
        DocxBlocks.InsertAt(this, table, row, index);
        var node = new DocxRow(doc, row);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override void Remove() => DocxBlocks.Detach(table);
    public override void MoveTo(Node newParent, int? index) => DocxBlocks.Move(newParent, table, index);
    public override void SetRaw(string raw) => DocxBlocks.ReplaceRaw(doc, table, raw);
}

sealed class DocxRow(DocxDocument doc, W.TableRow row) : Node, IDocxContainer
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    public override string Kind => "row";
    public override object Anchor => row;
    public OpenXmlElement Container => row;

    protected override IEnumerable<Node> ProjectChildren() => DocxTable.Visible(row).Select(c => (Node)new DocxCell(doc, c));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["data"] = NodeJson.Compact(w => WriteData(w, row)) };
        var pr = row.TableRowProperties;
        if (pr?.GetFirstChild<W.TableHeader>() is { } header && (header.Val is null || header.Val.Value == W.OnOffOnlyValues.On)) props["header"] = "true";
        if (pr?.GetFirstChild<W.TableRowHeight>()?.Val?.Value is { } twips && twips > 0) props["height"] = DocxTable.EmuOfTwips(twips).ToString(Inv);
        return props;
    }

    /// <summary>Visible cells only: cells merged into the one above are left out.</summary>
    internal static void WriteData(Utf8JsonWriter w, W.TableRow row)
    {
        w.WriteStartArray();
        foreach (var cell in DocxTable.Visible(row)) w.WriteStringValue(DocxCell.TextOf(cell));
        w.WriteEndArray();
    }

    public override string GetRaw() => row.OuterXml;

    /// <summary>Adds or removes cells at the end until the row shows <paramref name="count"/> cells. Removal takes the last cell whether
    /// it is visible or hidden, so the cells that stay keep their grid columns.</summary>
    internal static void SetCellCount(DocxDocument doc, W.TableRow row, int count)
    {
        var cells = DocxTable.Visible(row).ToList();
        while (cells.Count < count)
        {
            var clone = cells.Count > 0 ? DocxTable.Plain((W.TableCell)cells[^1].CloneNode(true)) : DocxTable.NewCell(2000);
            DocxCell.SetText(doc, clone, "");
            row.Append(clone);
            cells.Add(clone);
        }
        while (DocxTable.Visible(row).Count() > count)
        {
            var last = row.Elements<W.TableCell>().Last();
            DocxTable.Release(last);
            last.Remove();
        }
    }

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "data":
                var texts = DocxTable.ParseCells(value);
                SetCellCount(doc, row, Math.Max(1, texts.Count));
                var cells = DocxTable.Visible(row).ToList();
                for (var i = 0; i < texts.Count; i++) DocxCell.SetText(doc, cells[i], texts[i]);
                if (row.Parent is W.Table t) DocxTable.SyncGrid(t);
                break;
            case "header":
                Properties().GetFirstChild<W.TableHeader>()?.Remove();
                if (value == "true") Properties().Append(new W.TableHeader());
                Trim();
                break;
            case "height":
                var twips = DocxTable.TwipsOfEmu(value);
                Properties().GetFirstChild<W.TableRowHeight>()?.Remove();
                if (twips > 0) Properties().Append(new W.TableRowHeight { Val = (uint)twips, HeightType = W.HeightRuleValues.AtLeast });
                Trim();
                break;
        }
    }

    W.TableRowProperties Properties() => row.TableRowProperties ??= new W.TableRowProperties();

    void Trim()
    {
        if (row.TableRowProperties is { HasChildren: false } pr) pr.Remove();
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var visible = DocxTable.Visible(row).ToList();
        var at = index is { } i ? Math.Clamp(i, 1, visible.Count + 1) : visible.Count + 1;
        var template = visible.Count == 0 ? null : visible[Math.Max(0, at - 2)]; // the cell the new one sits after, or the first
        var cell = template is null ? DocxTable.NewCell(2000) : DocxTable.Plain((W.TableCell)template.CloneNode(true));
        DocxCell.SetText(doc, cell, "");
        DocxBlocks.InsertAt(this, row, cell, index);
        if (row.Parent is W.Table t) DocxTable.SyncGrid(t);
        var node = new DocxCell(doc, cell);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override void Remove()
    {
        if (row.Parent is W.Table t && t.Elements<W.TableRow>().Count() == 1)
            throw new WriterException(ErrorCode.Validation, "A table cannot lose its last row", "Remove the table instead.");
        DocxTable.Promote(row);
        row.Remove();
    }

    public override void MoveTo(Node newParent, int? index) => DocxBlocks.Move(newParent, row, index);
    public override void SetRaw(string raw) => DocxBlocks.ReplaceRaw(doc, row, raw);
}

sealed class DocxCell(DocxDocument doc, W.TableCell cell) : Node, IDocxContainer
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    public override string Kind => "cell";
    public override object Anchor => cell;
    public OpenXmlElement Container => cell;

    protected override IEnumerable<Node> ProjectChildren() => DocxBlocks.Project(doc, cell);

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = TextOf(cell), ["html"] = HtmlOf(cell) };
        if (DocxParagraph.AlignOf(cell.GetFirstChild<W.Paragraph>()?.ParagraphProperties) is { } align) props["align"] = align;
        var pr = cell.TableCellProperties;
        if (pr?.Shading?.Fill?.Value is { } fill && !fill.Equals("auto", StringComparison.OrdinalIgnoreCase))
            props["fill"] = fill.ToUpperInvariant();
        var colspan = DocxTable.Span(cell);
        if (colspan > 1) props["colspan"] = colspan.ToString(Inv);
        var rowspan = DocxTable.RowSpan(cell);
        if (rowspan > 1) props["rowspan"] = rowspan.ToString(Inv);
        if (pr?.TableCellBorders is { } borders && (DocxTable.BordersOf(borders) ?? DocxTable.SidesOf(borders)) is { } kind) props["borders"] = kind;
        if (pr?.TableCellVerticalAlignment?.Val?.InnerText is { } valign)
            props["valign"] = valign switch { "center" => "middle", "bottom" => "bottom", _ => "top" };
        if (pr?.TableCellWidth is { } tw && tw.Type?.Value == W.TableWidthUnitValues.Dxa
            && int.TryParse(tw.Width?.Value, NumberStyles.Integer, Inv, out var twips) && twips > 0)
            props["width"] = DocxTable.EmuOfTwips(twips).ToString(Inv);
        return props;
    }

    /// <summary>The cell as plain lines (its text, and the table's and row's data): a page break in it is a line break here.</summary>
    internal static string TextOf(W.TableCell cell) => string.Join("\n", cell.Elements<W.Paragraph>().Select(DocxRuns.ParagraphText)).Replace('\f', '\n');

    string HtmlOf(W.TableCell cell) => string.Join("<br>", cell.Elements<W.Paragraph>().Select(p => Exporter.HtmlOf(new DocxParagraph(doc, p))));

    /// <summary>Replaces the cell content with one paragraph holding the runs, keeping the first paragraph's properties.</summary>
    internal static void SetRuns(DocxDocument doc, W.TableCell cell, IEnumerable<RunSpec> specs)
    {
        var first = cell.Elements<W.Paragraph>().FirstOrDefault();
        if (first is null)
        {
            first = new W.Paragraph();
            cell.Append(first);
        }
        foreach (var child in cell.ChildElements.Where(c => c is not W.TableCellProperties && !ReferenceEquals(c, first)).ToList()) child.Remove();
        DocxRuns.ReplaceContent(first, DocxRuns.MakeRuns(doc, specs, DocxRuns.BaseRunProperties(first)));
    }

    internal static void SetText(DocxDocument doc, W.TableCell cell, string text) => SetRuns(doc, cell, [new RunSpec(text)]);

    /// <summary>Writes new content into the cell's paragraphs in place (as paragraphs are written), so fields, pictures, bookmarks
    /// and comment anchors survive: when the new text has as many lines as the paragraphs hold now, each paragraph takes its
    /// own lines; otherwise the first paragraph takes everything and the others go.</summary>
    static void Write(DocxDocument doc, W.TableCell cell, IReadOnlyList<RunSpec> specs, DocxReplace.Source source)
    {
        var paragraphs = cell.Elements<W.Paragraph>().ToList();
        if (paragraphs.Count == 0)
        {
            SetRuns(doc, cell, specs);
            return;
        }
        var lines = Lines(specs);
        var held = paragraphs.Select(p => DocxRuns.ParagraphText(p).Count(c => c == '\n') + 1).ToList();
        if (held.Sum() == lines.Count && cell.ChildElements.All(e => e is W.TableCellProperties or W.Paragraph))
        {
            for (int i = 0, at = 0; i < paragraphs.Count; at += held[i], i++)
                DocxReplace.Paragraph(doc, paragraphs[i], Join(lines.Skip(at).Take(held[i]).ToList()), source);
            return;
        }
        foreach (var extra in cell.ChildElements.Where(e => e is not W.TableCellProperties && !ReferenceEquals(e, paragraphs[0])).ToList()) extra.Remove();
        DocxReplace.Paragraph(doc, paragraphs[0], specs, source);
    }

    /// <summary>The runs cut at line breaks: one list of runs per line.</summary>
    static List<List<RunSpec>> Lines(IEnumerable<RunSpec> specs)
    {
        var lines = new List<List<RunSpec>> { new() };
        foreach (var s in specs)
        {
            var parts = s.Text.Split('\n');
            for (var i = 0; i < parts.Length; i++)
            {
                if (i > 0) lines.Add([]);
                if (parts[i].Length > 0) lines[^1].Add(s with { Text = parts[i] });
            }
        }
        return lines;
    }

    static List<RunSpec> Join(IReadOnlyList<List<RunSpec>> lines)
    {
        var joined = new List<RunSpec>();
        for (var i = 0; i < lines.Count; i++)
        {
            if (i > 0) joined.Add(new RunSpec("\n"));
            joined.AddRange(lines[i]);
        }
        return joined;
    }

    public override string GetRaw() => cell.OuterXml;

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "text": Write(doc, cell, [new RunSpec(value)], DocxReplace.Source.Text); break;
            case "md": Write(doc, cell, InlineMarkdown.Parse(value).ToList(), DocxReplace.Source.Markdown); break;
            case "html": Write(doc, cell, InlineHtml.Parse(value, pageBreaks: true).ToList(), DocxReplace.Source.Html); break;
            case "align":
                foreach (var p in cell.Elements<W.Paragraph>())
                    (p.ParagraphProperties ??= new W.ParagraphProperties()).Justification = DocxParagraph.JustificationOf(value);
                break;
            case "fill":
                Properties().Shading = value == "none"
                    ? null
                    : new W.Shading { Val = W.ShadingPatternValues.Clear, Color = "auto", Fill = value };
                break;
            case "colspan": DocxTable.SetColspan(cell, int.Parse(value, Inv)); break;
            case "rowspan": DocxTable.SetRowspan(cell, int.Parse(value, Inv)); break;
            case "borders": Properties().TableCellBorders = DocxTable.MakeBorders<W.TableCellBorders>(value, null); break;
            case "valign":
                Properties().TableCellVerticalAlignment = new W.TableCellVerticalAlignment
                {
                    Val = value switch
                    {
                        "middle" => W.TableVerticalAlignmentValues.Center,
                        "bottom" => W.TableVerticalAlignmentValues.Bottom,
                        _ => W.TableVerticalAlignmentValues.Top,
                    },
                };
                break;
            case "width": DocxTable.SetCellWidth(cell, DocxTable.TwipsOfEmu(value)); break;
        }
    }

    W.TableCellProperties Properties() => cell.TableCellProperties ??= new W.TableCellProperties();

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index) => DocxBlocks.Add(doc, this, cell, kind, props, index);

    public override void Remove()
    {
        if (cell.Parent is W.TableRow r && DocxTable.Visible(r).Count() == 1)
            throw new WriterException(ErrorCode.Validation, "A row cannot lose its last cell", "Remove the row instead.");
        var table = cell.Parent?.Parent as W.Table;
        DocxTable.RemoveMerged(cell);
        if (table is not null) DocxTable.SyncGrid(table);
    }

    public override void MoveTo(Node newParent, int? index) => DocxBlocks.Move(newParent, cell, index);
    public override void SetRaw(string raw) => DocxBlocks.ReplaceRaw(doc, cell, raw);
}
