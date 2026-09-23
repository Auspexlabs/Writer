using System.Globalization;
using System.Text.Json;
using Writer.Core;
using Writer.Formats.Common;

namespace Writer.Formats.Markdown;

sealed class MdRoot(MdDocument doc) : Node
{
    public override string Kind => "document";
    public override string Format => "md";
    public override object Anchor => doc;
    protected override IEnumerable<Node> ProjectChildren() => [new MdBody(doc)];
    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string> { ["format"] = "md" };
}

/// <summary>A node whose runs can be edited in place.</summary>
interface IMdRunHost
{
    List<MdRun> Runs { get; }
    MdChunk Chunk { get; }
}

sealed class MdBody(MdDocument doc) : Node
{
    public override string Kind => "body";
    public override object Anchor => doc.Chunks;

    protected override IEnumerable<Node> ProjectChildren()
    {
        foreach (var (chunk, item) in doc.Flatten())
            yield return item is not null ? new MdParagraphNode(doc, item) : Wrap(doc, chunk.Block);
    }

    internal static Node Wrap(MdDocument doc, MdBlock block) => block switch
    {
        MdHeading h => new MdHeadingNode(doc, h),
        MdCode c => new MdCodeNode(doc, c),
        MdTable t => new MdTableNode(doc, t),
        MdImage i => new MdImageNode(doc, i),
        MdItems items => new MdParagraphNode(doc, items.Items[0]),
        _ => throw new InvalidOperationException("Raw chunks are not projected."),
    };

    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string>();

    public override string GetRaw() => doc.Render();

    public override void SetRaw(string raw) => doc.Reset(raw);

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        MdBlock block = kind switch
        {
            "heading" => new MdHeading { Level = props.TryGetValue("level", out var l) ? int.Parse(l, CultureInfo.InvariantCulture) : 1 },
            "paragraph" => new MdItems { Items = [new MdParagraph()] },
            "code" => new MdCode(),
            "table" => MdTableNode.New(props),
            "image" => new MdImage
            {
                Src = props.GetValueOrDefault("src") ?? throw new WriterException(ErrorCode.Validation, "An image needs src", "Example: --prop src=chart.png"),
            },
            _ => throw new WriterException(ErrorCode.UnsupportedKind, $"Cannot add {kind} to a markdown body", "Blocks: paragraph, heading, code, table, image."),
        };
        var chunk = doc.InsertBlock(block, index);
        var node = Wrap(doc, chunk.Block);
        foreach (var (name, value) in props)
            if (name is not ("src" or "rows" or "cols" or "data")) node.SetProp(name, value);
        doc.MergeAdjacentLists(doc.ChunkOf(node.Anchor));
        return node;
    }
}

sealed class MdHeadingNode(MdDocument doc, MdHeading heading) : Node, IMdRunHost
{
    public override string Kind => "heading";
    public override object Anchor => heading;
    public List<MdRun> Runs => heading.Runs;
    public MdChunk Chunk => doc.ChunkOf(heading);

    protected override IEnumerable<Node> ProjectChildren() => heading.Runs.Select(r => (Node)new MdRunNode(this, r));

    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string>
    {
        ["text"] = MdText.Of(heading.Runs),
        ["html"] = InlineHtml.Render(heading.Runs.Select(r => r.ToSpec())),
        ["level"] = heading.Level.ToString(CultureInfo.InvariantCulture),
    };

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "text": heading.Runs = [new MdRun { Text = value }]; break;
            case "md": heading.Runs = MdRun.FromSpecs(InlineMarkdown.Parse(value)); break;
            case "html": heading.Runs = MdRun.FromSpecs(InlineHtml.Parse(value)); break;
            case "level": heading.Level = int.Parse(value, CultureInfo.InvariantCulture); break;
        }
        Chunk.Dirty = true;
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index) => MdText.AddRun(doc, this, props, index);
    public override string GetRaw() => Chunk.Text;
    public override void SetRaw(string raw) => doc.ReplaceChunkRaw(Chunk, raw);
    public override void Remove() => doc.RemoveChunk(Chunk);
    public override void MoveTo(Node newParent, int? index) => MdText.MoveBlock(doc, heading, newParent, index);
}

sealed class MdParagraphNode(MdDocument doc, MdParagraph item) : Node, IMdRunHost
{
    public override string Kind => "paragraph";
    public override object Anchor => item;
    public List<MdRun> Runs => item.Runs;
    public MdChunk Chunk => doc.ChunkOf(item);

    protected override IEnumerable<Node> ProjectChildren() => item.Runs.Select(r => (Node)new MdRunNode(this, r));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = MdText.Of(item.Runs), ["html"] = InlineHtml.Render(item.Runs.Select(r => r.ToSpec())) };
        if (item.List != "none")
        {
            props["list"] = item.List;
            props["level"] = item.Level.ToString(CultureInfo.InvariantCulture);
        }
        return props;
    }

    public override void SetProp(string name, string value)
    {
        var chunk = Chunk;
        MdDocument.EnsureEditable(chunk);
        switch (name)
        {
            case "text": item.Runs = [new MdRun { Text = value }]; break;
            case "md": item.Runs = MdRun.FromSpecs(InlineMarkdown.Parse(value)); break;
            case "html": item.Runs = MdRun.FromSpecs(InlineHtml.Parse(value)); break;
            case "list":
                item.List = value;
                if (value == "none") item.Level = 0;
                break;
            case "level":
                if (item.List == "none")
                    throw new WriterException(ErrorCode.Validation, "level applies to list paragraphs", "Set list=bullet or list=number first.");
                item.Level = int.Parse(value, CultureInfo.InvariantCulture);
                break;
        }
        chunk.Dirty = true;
        if (name == "list") doc.MergeAdjacentLists(chunk);
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        MdDocument.EnsureEditable(Chunk);
        return MdText.AddRun(doc, this, props, index);
    }

    public override string GetRaw() => Chunk.Text;
    public override void SetRaw(string raw) => doc.ReplaceChunkRaw(Chunk, raw);

    public override void Remove()
    {
        MdDocument.EnsureEditable(Chunk);
        doc.RemoveItem(item);
    }

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent is not MdBody)
            throw new WriterException(ErrorCode.Validation, "Markdown paragraphs can only move within the body", "Use --to /body.");
        MdDocument.EnsureEditable(Chunk);
        doc.RemoveItem(item);
        var chunk = doc.InsertBlock(new MdItems { Items = [item] }, index);
        doc.MergeAdjacentLists(chunk);
    }
}

sealed class MdCodeNode(MdDocument doc, MdCode code) : Node
{
    public override string Kind => "code";
    public override object Anchor => code;
    MdChunk Chunk => doc.ChunkOf(code);

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = code.Text };
        if (code.Lang.Length > 0) props["lang"] = code.Lang;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        if (name == "text") code.Text = value.ReplaceLineEndings("\n");
        if (name == "lang") code.Lang = value.Trim();
        Chunk.Dirty = true;
    }

    public override string GetRaw() => Chunk.Text;
    public override void SetRaw(string raw) => doc.ReplaceChunkRaw(Chunk, raw);
    public override void Remove() => doc.RemoveChunk(Chunk);
    public override void MoveTo(Node newParent, int? index) => MdText.MoveBlock(doc, code, newParent, index);
}

sealed class MdImageNode(MdDocument doc, MdImage image) : Node
{
    public override string Kind => "image";
    public override object Anchor => image;
    MdChunk Chunk => doc.ChunkOf(image);

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["src"] = image.Src };
        if (image.Alt.Length > 0) props["alt"] = image.Alt;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        if (name == "src") image.Src = value;
        if (name == "alt") image.Alt = value;
        Chunk.Dirty = true;
    }

    public override (string ContentType, byte[] Data)? GetBinary()
    {
        if (image.Src.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || image.Src.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) return null;
        var path = System.IO.Path.IsPathRooted(image.Src) ? image.Src : System.IO.Path.Combine(doc.BaseDirectory, image.Src);
        if (!File.Exists(path)) return null;
        var bytes = File.ReadAllBytes(path);
        try
        {
            return (ImageInfo.Read(bytes).ContentType, bytes);
        }
        catch (WriterException)
        {
            return null;
        }
    }

    public override string GetRaw() => Chunk.Text;
    public override void SetRaw(string raw) => doc.ReplaceChunkRaw(Chunk, raw);
    public override void Remove() => doc.RemoveChunk(Chunk);
    public override void MoveTo(Node newParent, int? index) => MdText.MoveBlock(doc, image, newParent, index);
}

sealed class MdTableNode(MdDocument doc, MdTable table) : Node
{
    public override string Kind => "table";
    public override object Anchor => table;
    MdChunk Chunk => doc.ChunkOf(table);

    protected override IEnumerable<Node> ProjectChildren() => table.Rows.Select(r => (Node)new MdRowNode(doc, table, r));

    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string>
    {
        ["rows"] = table.Rows.Count.ToString(CultureInfo.InvariantCulture),
        ["cols"] = Cols(table).ToString(CultureInfo.InvariantCulture),
        ["data"] = NodeJson.Compact(w =>
        {
            w.WriteStartArray();
            foreach (var row in table.Rows) MdRowNode.WriteData(w, row);
            w.WriteEndArray();
        }),
    };

    internal static int Cols(MdTable table) => table.Rows.Count == 0 ? 0 : table.Rows.Max(r => r.Cells.Count);

    public static MdTable New(IReadOnlyDictionary<string, string> props)
    {
        var table = new MdTable();
        var data = props.TryGetValue("data", out var json) ? ParseRows(json) : null;
        var rows = props.TryGetValue("rows", out var r) ? int.Parse(r, CultureInfo.InvariantCulture) : data?.Count ?? 0;
        var cols = props.TryGetValue("cols", out var c) ? int.Parse(c, CultureInfo.InvariantCulture) : data?.Max(x => x.Count) ?? 0;
        if (data is not null)
        {
            rows = Math.Max(rows, data.Count);
            cols = Math.Max(cols, data.Count == 0 ? 0 : data.Max(x => x.Count));
        }
        if (rows < 1 || cols < 1)
            throw new WriterException(ErrorCode.Validation, "A table needs rows and cols, or data", "Example: --prop rows=3 --prop cols=4, or --prop data='[[\"a\",\"b\"]]'");
        Resize(table, rows, cols);
        if (data is not null) Fill(table, data);
        return table;
    }

    internal static List<List<string>> ParseRows(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Array || parsed.RootElement.EnumerateArray().Any(x => x.ValueKind != JsonValueKind.Array))
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

    static void Resize(MdTable table, int rows, int cols)
    {
        while (table.Rows.Count < rows) table.Rows.Add(new MdRow());
        while (table.Rows.Count > rows) table.Rows.RemoveAt(table.Rows.Count - 1);
        foreach (var row in table.Rows)
        {
            while (row.Cells.Count < cols) row.Cells.Add(new MdCell());
            while (row.Cells.Count > cols) row.Cells.RemoveAt(row.Cells.Count - 1);
        }
        while (table.Aligns.Count < cols) table.Aligns.Add(null);
        while (table.Aligns.Count > cols) table.Aligns.RemoveAt(table.Aligns.Count - 1);
    }

    static void Fill(MdTable table, List<List<string>> data)
    {
        Resize(table, Math.Max(1, data.Count), Math.Max(1, data.Count == 0 ? 1 : data.Max(x => x.Count)));
        for (var i = 0; i < data.Count; i++)
            for (var j = 0; j < data[i].Count; j++)
                table.Rows[i].Cells[j].Runs = [new MdRun { Text = data[i][j] }];
    }

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "rows": Resize(table, int.Parse(value, CultureInfo.InvariantCulture), Math.Max(1, Cols(table))); break;
            case "cols": Resize(table, Math.Max(1, table.Rows.Count), int.Parse(value, CultureInfo.InvariantCulture)); break;
            case "data": Fill(table, ParseRows(value)); break;
        }
        Chunk.Dirty = true;
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var row = new MdRow { Cells = Enumerable.Range(0, Math.Max(1, Cols(table))).Select(_ => new MdCell()).ToList() };
        var at = index is { } i ? Math.Clamp(i - 1, 0, table.Rows.Count) : table.Rows.Count;
        table.Rows.Insert(at, row);
        Chunk.Dirty = true;
        var node = new MdRowNode(doc, table, row);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override string GetRaw() => Chunk.Text;
    public override void SetRaw(string raw) => doc.ReplaceChunkRaw(Chunk, raw);
    public override void Remove() => doc.RemoveChunk(Chunk);
    public override void MoveTo(Node newParent, int? index) => MdText.MoveBlock(doc, table, newParent, index);
}

sealed class MdRowNode(MdDocument doc, MdTable table, MdRow row) : Node
{
    public override string Kind => "row";
    public override object Anchor => row;
    MdChunk Chunk => doc.ChunkOf(row);

    protected override IEnumerable<Node> ProjectChildren() => row.Cells.Select((c, i) => (Node)new MdCellNode(doc, table, c, i));

    public override IReadOnlyDictionary<string, string> GetProps() =>
        new Dictionary<string, string> { ["data"] = NodeJson.Compact(w => WriteData(w, row)) };

    internal static void WriteData(Utf8JsonWriter w, MdRow row)
    {
        w.WriteStartArray();
        foreach (var cell in row.Cells) w.WriteStringValue(MdText.Of(cell.Runs));
        w.WriteEndArray();
    }

    public override void SetProp(string name, string value)
    {
        if (name != "data") return;
        var texts = MdTableNode.ParseCells(value);
        while (row.Cells.Count < texts.Count) row.Cells.Add(new MdCell());
        for (var i = 0; i < texts.Count; i++) row.Cells[i].Runs = [new MdRun { Text = texts[i] }];
        while (table.Aligns.Count < row.Cells.Count) table.Aligns.Add(null);
        Chunk.Dirty = true;
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var cell = new MdCell();
        var at = index is { } i ? Math.Clamp(i - 1, 0, row.Cells.Count) : row.Cells.Count;
        row.Cells.Insert(at, cell);
        while (table.Aligns.Count < row.Cells.Count) table.Aligns.Add(null);
        Chunk.Dirty = true;
        var node = new MdCellNode(doc, table, cell, at);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override void Remove()
    {
        if (table.Rows.Count == 1)
            throw new WriterException(ErrorCode.Validation, "A table cannot lose its last row", "Remove the table instead.");
        var chunk = Chunk;
        table.Rows.Remove(row);
        chunk.Dirty = true;
    }

    public override string GetRaw() => Chunk.Text;
    public override void SetRaw(string raw) => doc.ReplaceChunkRaw(Chunk, raw);
}

sealed class MdCellNode(MdDocument doc, MdTable table, MdCell cell, int column) : Node, IMdRunHost
{
    public override string Kind => "cell";
    public override object Anchor => cell;
    public List<MdRun> Runs => cell.Runs;
    public MdChunk Chunk => doc.ChunkOf(cell);

    protected override IEnumerable<Node> ProjectChildren() => cell.Runs.Select(r => (Node)new MdRunNode(this, r));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = MdText.Of(cell.Runs), ["html"] = InlineHtml.Render(cell.Runs.Select(r => r.ToSpec())) };
        if (table.Aligns.ElementAtOrDefault(column) is { } align) props["align"] = align;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "text": cell.Runs = [new MdRun { Text = value }]; break;
            case "md": cell.Runs = MdRun.FromSpecs(InlineMarkdown.Parse(value)); break;
            case "html": cell.Runs = MdRun.FromSpecs(InlineHtml.Parse(value)); break;
            case "align":
                while (table.Aligns.Count <= column) table.Aligns.Add(null);
                table.Aligns[column] = value == "justify" ? null : value;
                break;
        }
        Chunk.Dirty = true;
    }

    public override void Remove()
    {
        var row = table.Rows.First(r => r.Cells.Contains(cell));
        if (row.Cells.Count == 1)
            throw new WriterException(ErrorCode.Validation, "A row cannot lose its last cell", "Remove the row instead.");
        var chunk = Chunk;
        row.Cells.Remove(cell);
        chunk.Dirty = true;
    }

    public override string GetRaw() => Chunk.Text;
    public override void SetRaw(string raw) => doc.ReplaceChunkRaw(Chunk, raw);
}

sealed class MdRunNode(IMdRunHost host, MdRun run) : Node
{
    public override string Kind => "run";
    public override object Anchor => run;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = run.Text };
        if (run.Bold) props["bold"] = "true";
        if (run.Italic) props["italic"] = "true";
        if (run.Strike) props["strike"] = "true";
        if (run.Code) props["code"] = "true";
        if (run.Link is not null) props["link"] = run.Link;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        var chunk = host.Chunk;
        MdDocument.EnsureEditable(chunk);
        switch (name)
        {
            case "text": run.Text = value; break;
            case "md": run.Text = string.Concat(InlineMarkdown.Parse(value).Select(s => s.Text)); break;
            case "bold": run.Bold = value == "true"; break;
            case "italic": run.Italic = value == "true"; break;
            case "strike": run.Strike = value == "true"; break;
            case "code": run.Code = value == "true"; break;
            case "link": run.Link = value.Length == 0 ? null : value; break;
        }
        chunk.Dirty = true;
    }

    public override void Remove()
    {
        var chunk = host.Chunk;
        MdDocument.EnsureEditable(chunk);
        host.Runs.Remove(run);
        chunk.Dirty = true;
    }

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent is not IMdRunHost target)
            throw new WriterException(ErrorCode.Validation, $"Runs cannot move under {newParent.Kind}", "Move runs into a paragraph or heading.");
        var from = host.Chunk;
        MdDocument.EnsureEditable(from);
        MdDocument.EnsureEditable(target.Chunk);
        host.Runs.Remove(run);
        from.Dirty = true;
        var at = index is { } i ? Math.Clamp(i - 1, 0, target.Runs.Count) : target.Runs.Count;
        target.Runs.Insert(at, run);
        target.Chunk.Dirty = true;
    }

    public override string GetRaw() => MdWriter.Inline([run]);
}

/// <summary>Helpers shared by the markdown nodes.</summary>
static class MdText
{
    public static string Of(IEnumerable<MdRun> runs) => string.Concat(runs.Select(r => r.Text));

    public static Node AddRun(MdDocument doc, IMdRunHost host, IReadOnlyDictionary<string, string> props, int? index)
    {
        if (!props.ContainsKey("text") && !props.ContainsKey("md"))
            throw new WriterException(ErrorCode.Validation, "A run needs text", "Add --prop text=\"...\" or --prop md=\"...\".");
        var run = new MdRun();
        var at = index is { } i ? Math.Clamp(i - 1, 0, host.Runs.Count) : host.Runs.Count;
        host.Runs.Insert(at, run);
        host.Chunk.Dirty = true;
        var node = new MdRunNode(host, run);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public static void MoveBlock(MdDocument doc, MdBlock block, Node newParent, int? index)
    {
        if (newParent is not MdBody)
            throw new WriterException(ErrorCode.Validation, "Markdown blocks can only move within the body", "Use --to /body.");
        doc.RemoveChunk(doc.ChunkOf(block));
        doc.InsertBlock(block, index);
    }
}
