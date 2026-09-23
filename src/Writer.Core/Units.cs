using System.Globalization;

namespace Writer.Core;

/// <summary>Parses and formats the value types shared by every format: lengths, font sizes and colors.</summary>
public static class Units
{
    public const long EmuPerInch = 914400;
    public const long EmuPerCm = 360000;
    public const long EmuPerMm = 36000;
    public const long EmuPerPt = 12700;
    public const long EmuPerPx = 9525;

    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    /// <summary>"2cm", "20mm", "1in", "12pt", "96px", "720000emu" or a bare EMU integer → EMU.</summary>
    public static long ParseLength(string input)
    {
        var s = input.Trim().ToLowerInvariant();
        var (number, factor) =
            s.EndsWith("emu") ? (s[..^3], 1.0) :
            s.EndsWith("cm") ? (s[..^2], (double)EmuPerCm) :
            s.EndsWith("mm") ? (s[..^2], (double)EmuPerMm) :
            s.EndsWith("in") ? (s[..^2], (double)EmuPerInch) :
            s.EndsWith("pt") ? (s[..^2], (double)EmuPerPt) :
            s.EndsWith("px") ? (s[..^2], (double)EmuPerPx) :
            (s, 1.0);
        if (!double.TryParse(number.Trim(), NumberStyles.Float, Inv, out var value) || !double.IsFinite(value))
            throw new WriterException(ErrorCode.Validation, $"'{input}' is not a length",
                "Use a number with a unit: 2cm, 20mm, 1in, 12pt, 96px, or a bare EMU integer.");
        return (long)Math.Round(value * factor);
    }

    /// <summary>EMU → centimetres with at most three decimals, e.g. "2cm", "2.54cm".</summary>
    public static string FormatLength(long emu) => (emu / (double)EmuPerCm).ToString("0.###", Inv) + "cm";

    /// <summary>"12", "12pt", "10.5" → canonical points string "12", "10.5".</summary>
    public static string ParsePoints(string input)
    {
        var s = input.Trim().ToLowerInvariant();
        if (s.EndsWith("pt")) s = s[..^2].Trim();
        if (!double.TryParse(s, NumberStyles.Float, Inv, out var value) || !(value > 0) || value > 4000)
            throw new WriterException(ErrorCode.Validation, $"'{input}' is not a font size", "Use points, e.g. 12 or 10.5pt.");
        return Math.Round(value, 2).ToString("0.##", Inv);
    }

    /// <summary>Canonical points string → "12pt".</summary>
    public static string FormatPoints(string canonical) => canonical + "pt";

    static readonly Dictionary<string, string> NamedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000", ["white"] = "FFFFFF", ["red"] = "FF0000", ["green"] = "008000",
        ["blue"] = "0000FF", ["yellow"] = "FFFF00", ["orange"] = "FFA500", ["purple"] = "800080",
        ["gray"] = "808080", ["grey"] = "808080", ["silver"] = "C0C0C0", ["navy"] = "000080",
        ["teal"] = "008080", ["maroon"] = "800000", ["olive"] = "808000", ["lime"] = "00FF00",
        ["aqua"] = "00FFFF", ["cyan"] = "00FFFF", ["fuchsia"] = "FF00FF", ["magenta"] = "FF00FF",
    };

    /// <summary>"FF0000", "#ff0000" or a basic color name → "FF0000". "none" passes through for properties that can be cleared.</summary>
    public static string ParseColor(string input)
    {
        var s = input.Trim();
        if (s.Equals("none", StringComparison.OrdinalIgnoreCase)) return "none";
        if (NamedColors.TryGetValue(s, out var named)) return named;
        if (s.StartsWith('#')) s = s[1..];
        if (s.Length == 6 && s.All(Uri.IsHexDigit)) return s.ToUpperInvariant();
        throw new WriterException(ErrorCode.Validation, $"'{input}' is not a color",
            "Use RRGGBB hex (FF0000 or #FF0000) or a basic name: " + string.Join(", ", NamedColors.Keys.Take(8)) + ", ...");
    }
}
