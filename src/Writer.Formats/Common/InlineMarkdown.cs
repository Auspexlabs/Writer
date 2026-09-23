using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Writer.Core;

namespace Writer.Formats.Common;

/// <summary>Turns inline markdown into runs. Block structure is flattened: paragraphs are joined with line breaks.
/// Written markdown may mark tracked changes with inline &lt;ins&gt; / &lt;del&gt; tags, as html does.</summary>
public static class InlineMarkdown
{
    static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder().UseEmphasisExtras().UseAutoLinks().Build();

    public static List<RunSpec> Parse(string markdown)
    {
        var runs = new List<RunSpec>();
        var first = true;
        foreach (var block in Markdig.Markdown.Parse(markdown, Pipeline).Descendants<LeafBlock>())
        {
            if (!first) Add(runs, new RunSpec("\n"));
            first = false;
            if (block is CodeBlock code) Add(runs, new RunSpec(code.Lines.ToString(), Code: true));
            else if (block.Inline is { } inline) Walk(inline, runs, new RunSpec(""), revisions: true);
        }
        return runs;
    }

    /// <summary>Runs for an already parsed inline tree.</summary>
    public static List<RunSpec> FromInline(ContainerInline? inline)
    {
        var runs = new List<RunSpec>();
        if (inline is not null) Walk(inline, runs, new RunSpec(""));
        return runs;
    }

    /// <summary>With revisions, ins / del tags change the style of the siblings that follow them; otherwise (markdown documents,
    /// which keep their source) every raw tag stays literal text.</summary>
    static void Walk(ContainerInline container, List<RunSpec> runs, RunSpec style, bool revisions = false)
    {
        foreach (var inline in container)
        {
            switch (inline)
            {
                case HtmlInline tag when revisions && InlineHtml.RevisionTag(tag.Tag) is { } change:
                    style = style with { Change = change.Change, Author = change.Author, Date = change.Date };
                    break;
                case LiteralInline literal:
                    Add(runs, style with { Text = literal.Content.ToString() });
                    break;
                case CodeInline code:
                    Add(runs, style with { Text = code.Content.ToString(), Code = true });
                    break;
                case LineBreakInline lineBreak:
                    Add(runs, style with { Text = lineBreak.IsHard ? "\n" : " " });
                    break;
                case EmphasisInline emphasis:
                    Walk(emphasis, runs, emphasis.DelimiterChar switch
                    {
                        '*' or '_' when emphasis.DelimiterCount >= 2 => style with { Bold = true },
                        '*' or '_' => style with { Italic = true },
                        '~' when emphasis.DelimiterCount >= 2 => style with { Strike = true },
                        _ => style,
                    }, revisions);
                    break;
                case LinkInline link:
                    Walk(link, runs, link.IsImage ? style : style with { Link = link.Url }, revisions);
                    break;
                case AutolinkInline auto:
                    Add(runs, style with { Text = auto.Url, Link = (auto.IsEmail ? "mailto:" : "") + auto.Url });
                    break;
                case HtmlInline html:
                    Add(runs, style with { Text = html.Tag });
                    break;
                case ContainerInline other:
                    Walk(other, runs, style, revisions);
                    break;
            }
        }
    }

    static void Add(List<RunSpec> runs, RunSpec run)
    {
        if (run.Text.Length == 0) return;
        if (runs.Count > 0 && runs[^1] with { Text = "" } == run with { Text = "" })
            runs[^1] = runs[^1] with { Text = runs[^1].Text + run.Text };
        else runs.Add(run);
    }
}
