using System.IO.Compression;
using System.Text;
using Writer.Core;

namespace Writer.Formats.Compat;

/// <summary>A format Writer opens but never writes back: Office 97-2003 and WPS binaries, macro-enabled and template
/// OOXML, OpenDocument, RTF, HTML, CSV. The file is converted in memory into a Word, Excel or PowerPoint document
/// (<see cref="Target"/>), which the tree, the editors and <c>export</c> then see; the original stays as it was,
/// like Office's compatibility mode. The bytes decide the reader, not the extension: a .wps that is a Word binary,
/// a .xls that is really an HTML table, a .doc that is RTF all open.</summary>
public sealed class CompatAdapter(string format, string target, params string[] extensions) : IFormatAdapter
{
    public string Format => format;
    /// <summary>docx, xlsx or pptx: what the file becomes when it is opened, and what 另存为 offers.</summary>
    public string Target => target;
    public IReadOnlyList<string> Extensions { get; } = extensions;
    public bool CanWrite => false;

    public Document Create() =>
        throw new WriterException(ErrorCode.FormatReadonly, $"{format} files cannot be created", $"Create a .{target} instead: writer create <file>.{target}");

    public Document Open(Stream stream)
    {
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        var warnings = new List<string>();
        var inner = Compat.Convert(ms.ToArray(), target, format, warnings);
        return new CompatDocument(format, inner, warnings);
    }
}

/// <summary>The converted document: the inner docx/xlsx/pptx's tree under the legacy format's name, so the registry
/// keeps it read-only. <c>export</c> to the inner format hands the inner document out unchanged.</summary>
public sealed class CompatDocument(string format, Document inner, List<string> warnings) : Document
{
    public Document Inner { get; } = inner;
    /// <summary>What the reader had to drop, one line each.</summary>
    public IReadOnlyList<string> Warnings => warnings;
    public override string Format => format;
    public override Node Root => Inner.Root;
    public override void Save(Stream stream) => Inner.Save(stream);
    bool _disposed;
    public override void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Inner.Dispose();
    }
}

/// <summary>Sniffs the bytes and runs the matching reader.</summary>
public static class Compat
{
    static Compat() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    public static Document Convert(byte[] bytes, string target, string format, List<string> warnings)
    {
        if (Cfb.IsCfb(bytes))
        {
            var cfb = new Cfb(bytes);
            if (cfb.Has("WordDocument")) return Writers.Docx(DocReader.Read(bytes, warnings), warnings);
            if (cfb.Has("Workbook") || cfb.Has("Book")) return Writers.Xlsx(XlsReader.Read(bytes, warnings), warnings);
            if (cfb.Has("PowerPoint Document")) return Writers.Pptx(PptReader.Read(bytes, warnings), warnings);
            if (cfb.Has("EncryptedPackage")) throw new WriterException(ErrorCode.FormatError, "This file is password-protected", "Remove the password in Office, then open it again.");
            throw new WriterException(ErrorCode.FormatError, "Not a Word, Excel or PowerPoint 97-2003 file", $"Streams found: {string.Join(", ", cfb.StreamNames)}.");
        }
        if (bytes.Length > 4 && bytes[0] == 'P' && bytes[1] == 'K' && bytes[2] == 3 && bytes[3] == 4)
        {
            using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            if (zip.GetEntry("mimetype") is { } mimetype)
            {
                using var reader = new StreamReader(mimetype.Open());
                var mime = reader.ReadToEnd().Trim();
                if (mime.Contains("opendocument.text")) return Writers.Docx(OdfReader.Text(zip, warnings), warnings);
                if (mime.Contains("opendocument.spreadsheet")) return Writers.Xlsx(OdfReader.Sheets(zip, warnings), warnings);
                if (mime.Contains("opendocument.presentation")) return Writers.Pptx(OdfReader.Deck(zip, warnings), warnings);
                throw new WriterException(ErrorCode.FormatError, $"Unsupported OpenDocument type {mime}", "Text, spreadsheet and presentation documents open.");
            }
            if (zip.GetEntry("[Content_Types].xml") is not null) return OoxmlVariants.Open(bytes, zip);
            throw new WriterException(ErrorCode.FormatError, "Not an Office or OpenDocument package", "The zip has neither [Content_Types].xml nor a mimetype entry.");
        }
        var text = TextFiles.Decode(bytes);
        var head = text.AsSpan().TrimStart();
        if (head.StartsWith("{\\rtf")) return Writers.Docx(RtfReader.Read(text, warnings), warnings);
        if (HtmlReader.LooksLikeHtml(head))
            return target == "xlsx" ? Writers.Xlsx(HtmlReader.Sheets(text, warnings), warnings) : Writers.Docx(HtmlReader.Read(text, warnings), warnings);
        if (target == "xlsx") return Writers.Xlsx([TextFiles.Csv(text, format == "tsv" ? '\t' : null)], warnings);
        return Writers.Docx(TextFiles.Paragraphs(text), warnings);
    }
}
