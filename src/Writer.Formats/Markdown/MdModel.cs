using Writer.Core;

namespace Writer.Formats.Markdown;

/// <summary>Mutable run so nodes can anchor to it across edits.</summary>
sealed class MdRun
{
    public string Text = "";
    public bool Bold, Italic, Strike, Code;
    public string? Link;

    public RunSpec ToSpec() => new(Text, Bold, Italic, Strike, Code, Link);

    public static MdRun From(RunSpec s) => new() { Text = s.Text, Bold = s.Bold, Italic = s.Italic, Strike = s.Flattened().Strike, Code = s.Code, Link = s.Link };

    public static List<MdRun> FromSpecs(IEnumerable<RunSpec> specs) => specs.Select(From).ToList();
}

abstract class MdBlock;

sealed class MdHeading : MdBlock
{
    public int Level = 1;
    public List<MdRun> Runs = [];
}

/// <summary>One paragraph, or one list item when List is not "none".</summary>
sealed class MdParagraph
{
    public List<MdRun> Runs = [];
    public string List = "none";
    public int Level;
}

/// <summary>A paragraph or a whole list: one or more items rendered together.</summary>
sealed class MdItems : MdBlock
{
    public List<MdParagraph> Items = [];
    /// <summary>Some item holds content writer cannot rewrite (nested code, quotes, extra paragraphs). Read-only unless replaced raw.</summary>
    public bool Complex;
}

sealed class MdCode : MdBlock
{
    public string Lang = "";
    public string Text = "";
}

sealed class MdImage : MdBlock
{
    public string Src = "";
    public string Alt = "";
}

sealed class MdCell
{
    public List<MdRun> Runs = [];
}

sealed class MdRow
{
    public List<MdCell> Cells = [];
}

sealed class MdTable : MdBlock
{
    public List<MdRow> Rows = [];
    public List<string?> Aligns = [];
}

/// <summary>A block writer does not model (HTML, quotes, rules, front matter). Kept as source text only.</summary>
sealed class MdRaw : MdBlock;

/// <summary>One top-level block of the file with its original text, so unchanged blocks are written back byte for byte.</summary>
sealed class MdChunk
{
    public string? Original;
    /// <summary>Text between this block and the next one, or after the last block.</summary>
    public string Gap = "\n\n";
    public MdBlock Block = new MdRaw();
    public bool Dirty;

    public string Text => Dirty || Original is null ? MdWriter.Render(Block) : Original;
}
