using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;

namespace Writer.Formats.Docx;

public sealed class DocxAdapter : IFormatAdapter
{
    public string Format => "docx";
    public IReadOnlyList<string> Extensions { get; } = [".docx"];
    public bool CanWrite => true;

    public Document Create()
    {
        var ms = new MemoryStream();
        DocxTemplate.WriteBlank(ms);
        ms.Position = 0;
        return Open(ms);
    }

    public Document Open(Stream stream)
    {
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        ms.Position = 0;
        WordprocessingDocument package;
        try
        {
            package = WordprocessingDocument.Open(ms, true);
        }
        catch (Exception ex) when (ex is not WriterException)
        {
            throw new WriterException(ErrorCode.FormatError, $"Not a valid .docx file: {ex.Message}", "Check that the file opens in Word.");
        }
        if (package.MainDocumentPart?.Document?.Body is null)
        {
            package.Dispose();
            throw new WriterException(ErrorCode.FormatError, "The .docx has no document body", "Check that the file opens in Word.");
        }
        return new DocxDocument(package);
    }
}

/// <summary>An open Word package. The Open XML SDK tree is the native tree; Save clones it, so the document stays usable.</summary>
public sealed class DocxDocument(WordprocessingDocument package) : Document
{
    public WordprocessingDocument Package { get; } = package;
    public MainDocumentPart Main => Package.MainDocumentPart!;
    internal DocxStyles Styles { get; } = new(package);
    internal IEnumerable<KeyValuePair<string, string>> Namespaces => Main.Document!.NamespaceDeclarations;

    public override string Format => "docx";
    public override Node Root => new DocxRoot(this);

    int? _nextId;

    /// <summary>A w:id no revision, bookmark or comment in the document uses yet.</summary>
    internal string NextId()
    {
        _nextId ??= 1 + Main.Document!.Descendants().SelectMany(e => e.GetAttributes())
            .Where(a => a.LocalName == "id" && a.NamespaceUri == DocxTemplate.Ns)
            .Select(a => int.TryParse(a.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0)
            .DefaultIfEmpty(0).Max();
        return (_nextId++).Value.ToString(CultureInfo.InvariantCulture);
    }

    public override void Save(Stream stream)
    {
        using var clone = Package.Clone(stream);
    }

    public override void Dispose() => Package.Dispose();
}
