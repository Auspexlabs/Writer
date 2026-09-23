using System.Text;
using Writer.Core;

namespace Writer.Formats.Markdown;

public sealed class MdAdapter : IFormatAdapter
{
    public string Format => "md";
    public IReadOnlyList<string> Extensions { get; } = [".md", ".markdown"];
    public bool CanWrite => true;

    public Document Create() => new MdDocument("", hadBom: false);

    public Document Open(Stream stream)
    {
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        var bytes = ms.ToArray();
        var hadBom = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF;
        var text = Encoding.UTF8.GetString(bytes, hadBom ? 3 : 0, bytes.Length - (hadBom ? 3 : 0));
        return new MdDocument(text, hadBom);
    }
}

/// <summary>A markdown file as a list of top-level chunks. Untouched chunks are written back byte for byte.</summary>
public sealed class MdDocument : Document
{
    internal string Prefix;
    internal List<MdChunk> Chunks;
    internal readonly string NewLine;
    readonly bool _hadBom;

    internal MdDocument(string text, bool hadBom)
    {
        (Prefix, Chunks) = MdParser.Parse(text);
        NewLine = text.Contains("\r\n", StringComparison.Ordinal) ? "\r\n" : "\n";
        _hadBom = hadBom;
    }

    public override string Format => "md";
    public override Node Root => new MdRoot(this);

    internal string BaseDirectory => SourcePath is { } p
        ? System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(p)) ?? Directory.GetCurrentDirectory()
        : Directory.GetCurrentDirectory();

    public string Render()
    {
        var sb = new StringBuilder(Prefix);
        foreach (var chunk in Chunks)
        {
            var regenerated = chunk.Dirty || chunk.Original is null;
            sb.Append(regenerated ? chunk.Text.ReplaceLineEndings(NewLine) : chunk.Text);
            sb.Append(chunk.Original is null ? chunk.Gap.ReplaceLineEndings(NewLine) : chunk.Gap);
        }
        return sb.ToString();
    }

    public override void Save(Stream stream)
    {
        if (_hadBom) stream.Write(Encoding.UTF8.GetPreamble());
        stream.Write(Encoding.UTF8.GetBytes(Render()));
    }

    /// <summary>Replaces the whole document text and re-parses it.</summary>
    internal void Reset(string text) => (Prefix, Chunks) = MdParser.Parse(text);

    /// <summary>Block-level children in order: every non-raw chunk, list chunks expanded to one entry per item.</summary>
    internal List<(MdChunk Chunk, MdParagraph? Item)> Flatten()
    {
        var result = new List<(MdChunk, MdParagraph?)>();
        foreach (var chunk in Chunks)
        {
            switch (chunk.Block)
            {
                case MdItems items:
                    foreach (var item in items.Items) result.Add((chunk, item));
                    break;
                case MdRaw:
                    break;
                default:
                    result.Add((chunk, null));
                    break;
            }
        }
        return result;
    }

    /// <summary>The chunk that currently holds a model object (block, item, row, cell or run).</summary>
    internal MdChunk ChunkOf(object anchor) =>
        Chunks.FirstOrDefault(c => ReferenceEquals(c.Block, anchor) || Holds(c.Block, anchor))
        ?? throw new InvalidOperationException("The element is no longer in the document.");

    static bool Holds(MdBlock block, object anchor) => block switch
    {
        MdHeading h => h.Runs.Contains(anchor),
        MdItems items => items.Items.Contains(anchor) || items.Items.Any(i => i.Runs.Contains(anchor)),
        MdTable t => t.Rows.Contains(anchor) || t.Rows.Any(r => r.Cells.Contains(anchor) || r.Cells.Any(c => c.Runs.Contains(anchor))),
        _ => false,
    };

    /// <summary>Inserts a new block before the 1-based position among the flattened children; null appends.</summary>
    internal MdChunk InsertBlock(MdBlock block, int? index)
    {
        var chunk = new MdChunk { Block = block, Dirty = true };
        var flat = Flatten();
        if (index is { } i && i >= 1 && i <= flat.Count)
        {
            var (target, item) = flat[i - 1];
            var at = Chunks.IndexOf(target);
            if (item is not null && target.Block is MdItems items && items.Items.IndexOf(item) is var k and > 0)
            {
                var tail = new MdChunk
                {
                    Block = new MdItems { Items = items.Items.Skip(k).ToList(), Complex = items.Complex },
                    Dirty = true,
                    Gap = target.Gap,
                };
                items.Items.RemoveRange(k, items.Items.Count - k);
                target.Dirty = true;
                target.Gap = "\n\n";
                Chunks.Insert(at + 1, tail);
                at++;
            }
            Chunks.Insert(at, chunk);
            return chunk;
        }
        if (Chunks.Count > 0)
        {
            chunk.Gap = Chunks[^1].Gap;
            Chunks[^1].Gap = "\n\n";
        }
        else chunk.Gap = "\n";
        Chunks.Add(chunk);
        return chunk;
    }

    internal void RemoveChunk(MdChunk chunk)
    {
        var at = Chunks.IndexOf(chunk);
        if (at < 0) return;
        Chunks.RemoveAt(at);
        if (at == Chunks.Count && Chunks.Count > 0) Chunks[^1].Gap = chunk.Gap;
    }

    internal void RemoveItem(MdParagraph item)
    {
        var chunk = ChunkOf(item);
        var items = (MdItems)chunk.Block;
        items.Items.Remove(item);
        if (items.Items.Count == 0) RemoveChunk(chunk);
        else chunk.Dirty = true;
    }

    /// <summary>Joins a list chunk with neighbouring list chunks so nesting and numbering render as one list.</summary>
    internal void MergeAdjacentLists(MdChunk chunk)
    {
        if (chunk.Block is not MdItems mine || mine.Complex) return;
        var at = Chunks.IndexOf(chunk);
        if (at < 0) return;
        if (at > 0 && Chunks[at - 1].Block is MdItems previous && !previous.Complex && IsItem(previous.Items[^1]) && IsItem(mine.Items[0]))
        {
            previous.Items.AddRange(mine.Items);
            Chunks[at - 1].Dirty = true;
            Chunks[at - 1].Gap = chunk.Gap;
            Chunks.RemoveAt(at);
            at--;
            chunk = Chunks[at];
            mine = previous;
        }
        if (at + 1 < Chunks.Count && Chunks[at + 1].Block is MdItems next && !next.Complex && IsItem(mine.Items[^1]) && IsItem(next.Items[0]))
        {
            mine.Items.AddRange(next.Items);
            chunk.Dirty = true;
            chunk.Gap = Chunks[at + 1].Gap;
            Chunks.RemoveAt(at + 1);
        }
    }

    static bool IsItem(MdParagraph p) => p.List != "none";

    /// <summary>Replaces a chunk with the blocks parsed from raw markdown.</summary>
    internal void ReplaceChunkRaw(MdChunk chunk, string raw)
    {
        var parsed = MdParser.ParseFragment(raw.Trim());
        if (parsed.Count == 0)
            throw new WriterException(ErrorCode.Validation, "Raw markdown is empty", "Pass the markdown of one or more blocks.");
        var at = Chunks.IndexOf(chunk);
        parsed[^1].Gap = chunk.Gap;
        Chunks.RemoveAt(at);
        Chunks.InsertRange(at, parsed);
    }

    internal static void EnsureEditable(MdChunk chunk)
    {
        if (chunk.Block is MdItems { Complex: true })
            throw new WriterException(ErrorCode.Validation, "This list holds nested content (code, quotes or extra paragraphs) that writer cannot rewrite",
                "Read it with 'get --raw' and replace it with 'set --raw', or edit the file directly.");
    }
}
