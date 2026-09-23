using System.Globalization;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using Writer.Core;

namespace Writer.Formats.Pdf;

/// <summary>Read-only PDF access: pages, text blocks and images. Editing means exporting first.</summary>
public sealed class PdfAdapter : IFormatAdapter
{
    public string Format => "pdf";
    public IReadOnlyList<string> Extensions { get; } = [".pdf"];
    public bool CanWrite => false;

    public Document Create() =>
        throw new WriterException(ErrorCode.FormatReadonly, "PDF files cannot be created here", "Create a .docx or .md instead and print it to PDF with another tool.");

    public Document Open(Stream stream)
    {
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        try
        {
            return new PdfDocumentModel(PdfDocument.Open(ms.ToArray()));
        }
        catch (Exception ex) when (ex is not WriterException)
        {
            throw new WriterException(ErrorCode.FormatError, $"Not a valid PDF: {ex.Message}", "Check that the file opens in a PDF viewer.");
        }
    }
}

public sealed class PdfDocumentModel(PdfDocument pdf) : Document
{
    List<Page>? _pages;
    readonly Dictionary<int, List<TextBlock>> _blocks = new();

    internal PdfDocument Pdf => pdf;
    public override string Format => "pdf";
    public override Node Root => new PdfRoot(this);

    public override void Save(Stream stream) =>
        throw new WriterException(ErrorCode.FormatReadonly, "PDF files are read-only", "Export first: writer export file.pdf --to file.md");

    public override void Dispose() => pdf.Dispose();

    internal List<Page> Pages => _pages ??= pdf.GetPages().ToList();

    internal List<TextBlock> Blocks(Page page)
    {
        if (_blocks.TryGetValue(page.Number, out var cached)) return cached;
        var words = page.GetWords().ToList();
        var blocks = words.Count == 0 ? [] : DocstrumBoundingBoxes.Instance.GetBlocks(words)
            .OrderByDescending(b => Math.Round(b.BoundingBox.Top)).ThenBy(b => b.BoundingBox.Left).ToList();
        _blocks[page.Number] = blocks;
        return blocks;
    }

    internal static string Emu(double points) => ((long)Math.Round(points * Units.EmuPerPt)).ToString(CultureInfo.InvariantCulture);
}

sealed class PdfRoot(PdfDocumentModel doc) : Node
{
    public override string Kind => "document";
    public override string Format => "pdf";
    public override object Anchor => doc.Pdf;

    protected override IEnumerable<Node> ProjectChildren() => doc.Pages.Select(p => (Node)new PdfPage(doc, p));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["format"] = "pdf", ["pages"] = doc.Pages.Count.ToString(CultureInfo.InvariantCulture) };
        if (doc.Pdf.Information.Title is { Length: > 0 } title) props["title"] = title;
        return props;
    }
}

sealed class PdfPage(PdfDocumentModel doc, Page page) : Node
{
    public override string Kind => "page";
    public override object Anchor => page;

    protected override IEnumerable<Node> ProjectChildren()
    {
        var blocks = doc.Blocks(page).Select(b => (Top: b.BoundingBox.Top, Left: b.BoundingBox.Left, Node: (Node)new PdfText(page, b)));
        var images = page.GetImages().Select(i => (Top: i.BoundingBox.Top, Left: i.BoundingBox.Left, Node: (Node)new PdfImage(page, i)));
        return blocks.Concat(images).OrderByDescending(x => Math.Round(x.Top)).ThenBy(x => x.Left).Select(x => x.Node);
    }

    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string>
    {
        ["width"] = PdfDocumentModel.Emu(page.Width),
        ["height"] = PdfDocumentModel.Emu(page.Height),
        ["text"] = string.Join("\n\n", doc.Blocks(page).Select(b => b.Text)),
    };
}

sealed class PdfText(Page page, TextBlock block) : Node
{
    public override string Kind => "text";
    public override object Anchor => block;

    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string>
    {
        ["text"] = block.Text,
        ["x"] = PdfDocumentModel.Emu(block.BoundingBox.Left),
        ["y"] = PdfDocumentModel.Emu(page.Height - block.BoundingBox.Top),
        ["w"] = PdfDocumentModel.Emu(block.BoundingBox.Width),
        ["h"] = PdfDocumentModel.Emu(block.BoundingBox.Height),
    };
}

sealed class PdfImage(Page page, IPdfImage image) : Node
{
    public override string Kind => "image";
    public override object Anchor => image;

    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string>
    {
        ["x"] = PdfDocumentModel.Emu(image.BoundingBox.Left),
        ["y"] = PdfDocumentModel.Emu(page.Height - image.BoundingBox.Top),
        ["w"] = PdfDocumentModel.Emu(image.BoundingBox.Width),
        ["h"] = PdfDocumentModel.Emu(image.BoundingBox.Height),
    };

    public override (string ContentType, byte[] Data)? GetBinary()
    {
        try
        {
            if (image.TryGetPng(out var png)) return ("image/png", png);
            var raw = image.RawBytes.ToArray();
            if (raw.Length > 2 && raw[0] == 0xFF && raw[1] == 0xD8) return ("image/jpeg", raw);
        }
        catch (Exception)
        {
            // unsupported image encodings are simply not exported
        }
        return null;
    }
}
