using Writer.Core;
using Writer.Formats.Compat;
using Writer.Formats.Docx;
using Writer.Formats.Markdown;
using Writer.Formats.MindMap;
using Writer.Formats.Pdf;
using Writer.Formats.Pptx;
using Writer.Formats.Xlsx;

namespace Writer.Formats;

/// <summary>The formats this build knows, looked up by file extension or by name. The native ones first; then the
/// compatibility formats, which open as a Word, Excel or PowerPoint document and are never written back.</summary>
public static class Adapters
{
    public static IReadOnlyList<IFormatAdapter> All { get; } =
    [
        new DocxAdapter(), new MdAdapter(), new PptxAdapter(), new XlsxAdapter(), new PdfAdapter(), new MmAdapter(), new XmindAdapter(),
        new CompatAdapter("doc", "docx", ".doc", ".dot", ".wps", ".wpt"),
        new CompatAdapter("docm", "docx", ".docm", ".dotx", ".dotm"),
        new CompatAdapter("odt", "docx", ".odt", ".ott"),
        new CompatAdapter("rtf", "docx", ".rtf"),
        new CompatAdapter("html", "docx", ".html", ".htm"),
        new CompatAdapter("xls", "xlsx", ".xls", ".xlt", ".et", ".ett"),
        new CompatAdapter("xlsm", "xlsx", ".xlsm", ".xltx", ".xltm"),
        new CompatAdapter("ods", "xlsx", ".ods", ".ots"),
        new CompatAdapter("csv", "xlsx", ".csv"),
        new CompatAdapter("tsv", "xlsx", ".tsv"),
        new CompatAdapter("ppt", "pptx", ".ppt", ".pot", ".pps", ".dps", ".dpt"),
        new CompatAdapter("pptm", "pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"),
        new CompatAdapter("odp", "pptx", ".odp", ".otp"),
    ];

    static readonly Dictionary<string, string> Aliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["word"] = "docx", ["excel"] = "xlsx", ["powerpoint"] = "pptx", ["markdown"] = "md", ["mindmap"] = "mm", ["freemind"] = "mm",
    };

    /// <summary>Every readable extension (without the dot) and the editor it opens in: its own format for the native
    /// ones, docx/xlsx/pptx for the compatibility ones. The one list the app's open dialogs and drop targets follow.</summary>
    public static IReadOnlyDictionary<string, string> OpensAs { get; } =
        All.SelectMany(a => a.Extensions.Select(e => (Ext: e.TrimStart('.'), Editor: a is CompatAdapter c ? c.Target : a.Format)))
            .ToDictionary(x => x.Ext, x => x.Editor, StringComparer.OrdinalIgnoreCase);

    public static string? CanonicalFormat(string name)
    {
        var n = name.Trim().TrimStart('.');
        if (All.FirstOrDefault(a => a.Format.Equals(n, StringComparison.OrdinalIgnoreCase) || a.Extensions.Contains("." + n.ToLowerInvariant())) is { } byName) return byName.Format;
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
