using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Formats.Pptx;
using A = DocumentFormat.OpenXml.Drawing;

namespace Writer.Formats.Xlsx;

/// <summary>Excel's colours as RRGGBB for drawing: rgb as written, theme slots through the workbook's theme with Excel's tint, indexed
/// ones through the legacy palette (or the file's own). Automatic colours are null. Writes go the other way through <see cref="Like{T}"/>.</summary>
sealed class XlsxColors
{
    /// <summary>Excel's 64-entry legacy palette (indexed 0–63); 64 and 65 are the system's automatic colours.</summary>
    static readonly string[] Legacy =
    [
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
        "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
        "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
    ];

    /// <summary>Theme indexes as SpreadsheetML counts them (light before dark), and the Office theme for a workbook without one.</summary>
    static readonly string[] Slots = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    static readonly string[] Office = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"];

    readonly WorkbookPart _workbook;
    readonly string[] _theme;
    readonly string[] _indexed;

    public XlsxColors(WorkbookPart workbook)
    {
        _workbook = workbook;
        var styles = workbook.WorkbookStylesPart?.Stylesheet;
        var scheme = workbook.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        _theme = Slots.Select((slot, i) => Scheme(scheme, slot) ?? Office[i]).ToArray();
        var own = styles?.Colors?.IndexedColors?.Elements<RgbColor>().Select(c => XlsxStyles.Rgb(c.Rgb?.Value)).ToList();
        _indexed = own is { Count: > 0 } ? own.Select((c, i) => c ?? Legacy.ElementAtOrDefault(i) ?? "000000").ToArray() : Legacy;
    }

    static string? Scheme(A.ColorScheme? scheme, string slot) =>
        scheme?.ChildElements.FirstOrDefault(e => e.LocalName == slot)?.FirstChild switch
        {
            A.RgbColorModelHex rgb => rgb.Val?.Value?.ToUpperInvariant(),
            A.SystemColor sys => sys.LastColor?.Value?.ToUpperInvariant() ?? (sys.Val?.InnerText == "window" ? "FFFFFF" : "000000"),
            _ => null,
        };

    /// <summary>The theme's accent colours in turn, as Excel colours chart series.</summary>
    public string Accent(int i) => _theme[4 + i % 6];

    /// <summary>RRGGBB, or null for an automatic colour.</summary>
    public string? Resolve(ColorType? color)
    {
        if (color is null || color.Auto?.Value == true) return null;
        var rgb = color.Rgb?.Value is { } argb ? XlsxStyles.Rgb(argb)
            : color.Theme?.Value is { } theme ? _theme.ElementAtOrDefault((int)theme)
            : color.Indexed?.Value is { } index ? _indexed.ElementAtOrDefault((int)index)
            : null;
        return rgb is null || color.Tint?.Value is not { } tint || tint == 0 ? rgb : Tint(rgb, tint);
    }

    /// <summary>Excel's tint: towards black below 0, towards white above, on the HLS lightness.</summary>
    static string Tint(string rgb, double tint)
    {
        var (h, s, l) = PptxLook.ToHsl(Convert.ToInt32(rgb[..2], 16) / 255.0, Convert.ToInt32(rgb[2..4], 16) / 255.0, Convert.ToInt32(rgb[4..], 16) / 255.0);
        l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
        var (r, g, b) = PptxLook.HslToRgb(h, s, Math.Clamp(l, 0, 1));
        return PptxLook.Hex(r, g, b);
    }

    /// <summary>A colour element for RRGGBB as the file would write it: the theme reference (with its tint) of a colour the stylesheet
    /// already uses that shows as this, else rgb. So a theme colour the editor read as RRGGBB goes back as the theme colour.</summary>
    public T Like<T>(string rrggbb) where T : ColorType, new()
    {
        var wanted = rrggbb.ToUpperInvariant();
        var theme = _workbook.WorkbookStylesPart?.Stylesheet?.Descendants<ColorType>().FirstOrDefault(c => c.Theme is not null && c.Auto is null && Resolve(c) == wanted);
        return theme is null ? new T { Rgb = "FF" + wanted } : new T { Theme = theme.Theme!.Value, Tint = theme.Tint is null ? null : new DoubleValue(theme.Tint.Value) };
    }
}
