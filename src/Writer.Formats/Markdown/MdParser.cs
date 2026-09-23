using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Writer.Formats.Common;

namespace Writer.Formats.Markdown;

/// <summary>Markdig does the parsing; this maps its blocks to the chunk model and keeps every block's source slice.</summary>
static class MdParser
{
    public static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UsePipeTables().UseEmphasisExtras().UseAutoLinks().UseYamlFrontMatter().Build();

    public static (string Prefix, List<MdChunk> Chunks) Parse(string text)
    {
        var blocks = Markdig.Markdown.Parse(text, Pipeline).ToList();
        var chunks = new List<MdChunk>();
        if (blocks.Count == 0) return (text, chunks);
        var prefix = text[..blocks[0].Span.Start];
        for (var i = 0; i < blocks.Count; i++)
        {
            var block = blocks[i];
            var start = block.Span.Start;
            var end = Math.Min(block.Span.End, text.Length - 1);
            while (end >= start && text[end] is '\n' or '\r') end--;
            var nextStart = i + 1 < blocks.Count ? Math.Max(blocks[i + 1].Span.Start, end + 1) : text.Length;
            chunks.Add(new MdChunk
            {
                Original = end >= start ? text[start..(end + 1)] : "",
                Gap = text[(end + 1)..nextStart],
                Block = Convert(block),
            });
        }
        return (prefix, chunks);
    }

    /// <summary>Parses a fragment on its own, e.g. text passed to set --raw.</summary>
    public static List<MdChunk> ParseFragment(string text) => Parse(text).Chunks;

    static MdBlock Convert(Block block) => block switch
    {
        HeadingBlock h => new MdHeading { Level = h.Level, Runs = Runs(h.Inline) },
        ParagraphBlock p when ImageOnly(p) is { } image => image,
        ParagraphBlock p => new MdItems { Items = [new MdParagraph { Runs = Runs(p.Inline) }] },
        FencedCodeBlock f => new MdCode { Lang = f.Info ?? "", Text = f.Lines.ToString() },
        CodeBlock c => new MdCode { Text = c.Lines.ToString() },
        ListBlock l => List(l),
        Table t => Table(t),
        _ => new MdRaw(),
    };

    static List<MdRun> Runs(ContainerInline? inline) => MdRun.FromSpecs(InlineMarkdown.FromInline(inline));

    static MdImage? ImageOnly(ParagraphBlock p)
    {
        if (p.Inline is null) return null;
        var meaningful = p.Inline.Where(i => i is not LiteralInline literal || !literal.Content.IsEmptyOrWhitespace()).ToList();
        if (meaningful.Count != 1 || meaningful[0] is not LinkInline { IsImage: true } link) return null;
        return new MdImage { Src = link.Url ?? "", Alt = string.Concat(InlineMarkdown.FromInline(link).Select(r => r.Text)) };
    }

    static MdItems List(ListBlock list)
    {
        var items = new MdItems();
        Flatten(list, 0, items);
        return items;
    }

    static void Flatten(ListBlock list, int level, MdItems into)
    {
        foreach (var item in list.OfType<ListItemBlock>())
        {
            var blocks = item.ToList();
            var first = blocks.FirstOrDefault() as ParagraphBlock;
            into.Items.Add(new MdParagraph
            {
                List = list.IsOrdered ? "number" : "bullet",
                Level = level,
                Runs = first is null ? [] : Runs(first.Inline),
            });
            foreach (var rest in blocks.Skip(first is null ? 0 : 1))
            {
                if (rest is ListBlock nested) Flatten(nested, level + 1, into);
                else into.Complex = true;
            }
        }
    }

    static MdTable Table(Table table)
    {
        var result = new MdTable();
        foreach (var column in table.ColumnDefinitions)
            result.Aligns.Add(column.Alignment switch
            {
                TableColumnAlign.Center => "center",
                TableColumnAlign.Right => "right",
                TableColumnAlign.Left => "left",
                _ => null,
            });
        foreach (var row in table.OfType<TableRow>())
        {
            var r = new MdRow();
            foreach (var cell in row.OfType<TableCell>())
            {
                var runs = new List<MdRun>();
                var firstBlock = true;
                foreach (var leaf in cell.Descendants<LeafBlock>())
                {
                    if (!firstBlock) runs.Add(new MdRun { Text = "\n" });
                    firstBlock = false;
                    runs.AddRange(Runs(leaf.Inline));
                }
                r.Cells.Add(new MdCell { Runs = runs });
            }
            result.Rows.Add(r);
        }
        return result;
    }
}
