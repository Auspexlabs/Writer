using Writer.Core;
using Writer.Formats.Docx;
using Writer.Formats.Markdown;
using Writer.Formats.MindMap;
using Writer.Formats.Pdf;
using Writer.Formats.Pptx;
using Writer.Formats.Xlsx;

namespace Writer.Formats;

/// <summary>The formats this build knows, looked up by file extension or by name.</summary>
public static class Adapters
{
    public static IReadOnlyList<IFormatAdapter> All { get; } = [new DocxAdapter(), new MdAdapter(), new PptxAdapter(), new XlsxAdapter(), new PdfAdapter(), new MmAdapter()];

    static readonly Dictionary<string, string> Aliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["word"] = "docx", ["excel"] = "xlsx", ["ppt"] = "pptx", ["powerpoint"] = "pptx", ["markdown"] = "md", ["mindmap"] = "mm", ["freemind"] = "mm",
    };

    public static string? CanonicalFormat(string name)
    {
        var n = name.Trim().TrimStart('.');
        n = Aliases.GetValueOrDefault(n, n).ToLowerInvariant();
        return All.Any(a => a.Format == n) ? n : null;
    }

    public static IFormatAdapter ForName(string name) =>
        CanonicalFormat(name) is { } format
            ? All.First(a => a.Format == format)
            : throw new WriterException(ErrorCode.UnknownFormat, $"Unknown format '{name}'",
                $"Formats: {string.Join(", ", All.Select(a => a.Format))}.");

    public static IFormatAdapter ForPath(string path)
    {
        var ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
        return All.FirstOrDefault(a => a.Extensions.Contains(ext))
            ?? throw new WriterException(ErrorCode.UnknownFormat, $"Unsupported file type '{ext}'",
                $"Supported: {string.Join(", ", All.SelectMany(a => a.Extensions))}.");
    }
}
