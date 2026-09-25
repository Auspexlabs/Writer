using System.Globalization;
using System.Text;

namespace Writer.Formats.Compat;

// The intermediate model the compatibility readers fill. A reader parses one legacy format into blocks (Word-like),
// sheets (Excel-like) or a deck (PowerPoint-like); Writers.cs turns the model into a real docx, xlsx or pptx through
// the engine's own node API, so every reader gets tables, lists, pictures and styles the same way.

public enum BlockKind { Heading, Paragraph, Table, Image, PageBreak }

/// <summary>One block of flowing content. Text is inline html (see <see cref="Inline"/>).</summary>
public sealed class Block
{
    public BlockKind Kind;
    /// <summary>Heading level 1-6, or list nesting 0-8 for list paragraphs.</summary>
    public int Level;
    public string Html = "";
    /// <summary>bullet, number, or null for a plain paragraph.</summary>
    public string? List;
    /// <summary>left, center, right, justify, or null.</summary>
    public string? Align;
    /// <summary>A paragraph style hint the target may know: Quote, Title, Subtitle.</summary>
    public string? Style;
    /// <summary>Table rows; every row is a list of cells.</summary>
    public List<List<TableCell>>? Rows;
    /// <summary>Picture bytes (PNG, JPEG, GIF or BMP) with its display size in cm when known.</summary>
    public byte[]? Image;
    public double? WidthCm, HeightCm;

    public static Block Heading(int level, string html) => new() { Kind = BlockKind.Heading, Level = Math.Clamp(level, 1, 6), Html = html };
    public static Block Paragraph(string html) => new() { Kind = BlockKind.Paragraph, Html = html };
    public static Block Table(List<List<TableCell>> rows) => new() { Kind = BlockKind.Table, Rows = rows };
    public static Block Picture(byte[] bytes, double? widthCm = null, double? heightCm = null) => new() { Kind = BlockKind.Image, Image = bytes, WidthCm = widthCm, HeightCm = heightCm };
    public static Block PageBreak() => new() { Kind = BlockKind.PageBreak };
}

public sealed class TableCell
{
    public string Html = "";
    public int ColSpan = 1, RowSpan = 1;
    /// <summary>Background as RRGGBB, or null.</summary>
    public string? Fill;
    public TableCell() { }
    public TableCell(string html) { Html = html; }
}

/// <summary>One worksheet: typed cells with their formatting, merges and custom sizes.</summary>
public sealed class SheetModel
{
    public string Name = "Sheet1";
    public List<CellModel> Cells = [];
    /// <summary>Merged ranges as A1:C3.</summary>
    public List<string> Merges = [];
    /// <summary>Column index (1-based) to width in characters.</summary>
    public Dictionary<int, double> ColWidths = [];
    /// <summary>Row index (1-based) to height in points.</summary>
    public Dictionary<int, double> RowHeights = [];
    public List<SheetImage> Images = [];
}

public sealed class SheetImage
{
    public byte[] Bytes = [];
    /// <summary>Top-left anchor cell, 1-based.</summary>
    public int Col = 1, Row = 1;
    public double WidthCm = 8, HeightCm = 6;
}

public sealed class CellModel
{
    /// <summary>1-based row and column.</summary>
    public int Row, Col;
    /// <summary>string, double, bool or DateTime; null for a styled empty cell.</summary>
    public object? Value;
    /// <summary>Formula text without the leading =, or null.</summary>
    public string? Formula;
    public bool Bold, Italic, Underline, Strike, Wrap;
    /// <summary>RRGGBB colors, or null.</summary>
    public string? Color, Fill;
    public string? Font;
    public double? SizePt;
    /// <summary>Number format code such as 0.00 or yyyy-mm-dd, or null for General.</summary>
    public string? Format;
    /// <summary>left, center, right, or null.</summary>
    public string? Align;
    /// <summary>top, middle, bottom, or null.</summary>
    public string? VAlign;
}

public sealed class DeckModel
{
    public double WidthCm = 25.4, HeightCm = 19.05;
    public List<SlideModel> Slides = [];
}

public sealed class SlideModel
{
    public List<ShapeModel> Shapes = [];
    public string? Notes;
    /// <summary>Solid background as RRGGBB, or null.</summary>
    public string? Background;
}

public enum ShapeKind { Text, Image, Table }

/// <summary>A shape on a slide; positions and sizes in cm from the slide's top-left corner.</summary>
public sealed class ShapeModel
{
    public ShapeKind Kind;
    public double X, Y, W, H;
    /// <summary>Text shapes: paragraphs (Html, List, Level, Align are used).</summary>
    public List<Block> Paragraphs = [];
    /// <summary>The slide's title placeholder.</summary>
    public bool IsTitle;
    public byte[]? Image;
    public List<List<TableCell>>? Rows;
    /// <summary>RRGGBB fill and outline, or null.</summary>
    public string? Fill, Line;
    /// <summary>Default font size in points for the shape's text, when the runs do not say.</summary>
    public double? SizePt;
    /// <summary>textbox, rect, roundRect, ellipse, triangle, diamond, rightArrow.</summary>
    public string? Geometry;
}

/// <summary>Inline html the engine's html property understands: b, i, u, s, a href, br, and span style with color,
/// background-color, font-size and font-family.</summary>
public static class Inline
{
    public static string Esc(string text) =>
        text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    /// <summary>One formatted run. Colors are RRGGBB; size in points; text is escaped here.</summary>
    public static string Run(string text, bool bold = false, bool italic = false, bool underline = false, bool strike = false,
        string? color = null, double? sizePt = null, string? font = null, string? highlight = null, string? link = null)
    {
        if (text.Length == 0) return "";
        var sb = new StringBuilder(Esc(text).Replace("\n", "<br>").Replace("\v", "<br>"));
        if (bold) sb.Insert(0, "<b>").Append("</b>");
        if (italic) sb.Insert(0, "<i>").Append("</i>");
        if (underline) sb.Insert(0, "<u>").Append("</u>");
        if (strike) sb.Insert(0, "<s>").Append("</s>");
        var style = new StringBuilder();
        if (color is not null) style.Append("color:#").Append(color).Append(';');
        if (highlight is not null) style.Append("background-color:#").Append(highlight).Append(';');
        if (sizePt is > 0) style.Append("font-size:").Append(sizePt.Value.ToString("0.#", CultureInfo.InvariantCulture)).Append("pt;");
        if (!string.IsNullOrEmpty(font)) style.Append("font-family:").Append(font.Replace(';', ' ')).Append(';');
        if (style.Length > 0) sb.Insert(0, $"<span style=\"{style}\">").Append("</span>");
        if (!string.IsNullOrEmpty(link)) sb.Insert(0, $"<a href=\"{Esc(link)}\">").Append("</a>");
        return sb.ToString();
    }

    /// <summary>RRGGBB from a 24-bit BGR value as Office's binary formats store colors.</summary>
    public static string Bgr(uint bgr) => $"{bgr & 0xFF:X2}{(bgr >> 8) & 0xFF:X2}{(bgr >> 16) & 0xFF:X2}";

    /// <summary>RRGGBB from r, g, b.</summary>
    public static string Rgb(int r, int g, int b) => $"{r & 0xFF:X2}{g & 0xFF:X2}{b & 0xFF:X2}";
}
