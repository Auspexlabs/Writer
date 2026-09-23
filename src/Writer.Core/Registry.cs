using System.Globalization;
using System.Text.Json;

namespace Writer.Core;

/// <summary>Every element kind and property the engine knows. Pure data plus validation; adapters implement the behaviour.</summary>
public static class Registry
{
    static readonly string[] Docx = ["docx"];
    static readonly string[] Md = ["md"];
    static readonly string[] Pptx = ["pptx"];
    static readonly string[] Xlsx = ["xlsx"];
    static readonly string[] Pdf = ["pdf"];
    static readonly string[] Mm = ["mm"];
    static readonly string[] DocxMd = ["docx", "md"];
    static readonly string[] DocxPptx = ["docx", "pptx"];
    static readonly string[] PptxPdf = ["pptx", "pdf"];
    static readonly string[] Office = ["docx", "pptx", "xlsx"];
    static readonly string[] Documents = ["docx", "md", "pptx"];
    static readonly string[] Visual = ["docx", "md", "pptx", "pdf"];
    static readonly string[] Tables = ["docx", "md", "pptx", "xlsx", "pdf"];
    static readonly string[] Pictures = Tables;
    static readonly string[] All = ["docx", "md", "pptx", "xlsx", "pdf", "mm"];
    /// <summary>Preset shapes a picture can be cropped to.</summary>
    public static readonly string[] PictureShapes = ["rect", "roundRect", "ellipse", "triangle", "diamond", "hexagon", "star5"];

    static Prop Text => new("text", PropType.String, "Plain text. Writing it replaces the content and keeps the node's own style.") { Example = "Revenue grew 25%", Formats = Documents };
    static Prop MdProp => new("md", PropType.Md, "Inline markdown: **bold**, *italic*, ~~strike~~, `code`, [link](url); <ins> / <del> tags mark tracked changes. Write-only; replaces the content.") { WriteOnly = true, Example = "Revenue **grew** 25%", Formats = Documents };
    static Prop HtmlProp => new("html", PropType.Html, "Inline HTML: b, i, u, s, code, a href, br (in Word, br style='page-break-before:always' is a page break), span style color / font-size / background-color / font-family, and ins / del (data-author, data-date) for tracked changes. Reading gives the content as HTML; writing replaces it.") { Example = "Revenue <b>grew</b> 25%", Formats = Documents };
    static Prop Accept => new("accept", PropType.Enum, "Write all to accept every tracked change here: insertions stay, deletions go.") { Values = ["all"], WriteOnly = true, Example = "all", Formats = Docx };
    static Prop Reject => new("reject", PropType.Enum, "Write all to reject every tracked change here: insertions go, deletions come back.") { Values = ["all"], WriteOnly = true, Example = "all", Formats = Docx };
    static Prop Highlight(string[] formats) => new("highlight", PropType.Color, "Background highlight color, or none.") { Example = "FFFF00", Formats = formats };
    static Prop PageBreakBefore => new("pageBreakBefore", PropType.Bool, "Starts on a new page. Reading gives it as computed when the paragraph's style sets it.") { Example = "true", Formats = Docx };
    static Prop ParagraphFill => new("fill", PropType.Color, "Paragraph shading, or none. Reading gives the style's shading as computed.") { Example = "D9D9D9", Formats = Docx };
    static Prop Align(string[]? formats) => new("align", PropType.Enum, "Horizontal alignment.") { Values = ["left", "center", "right", "justify"], Example = "center", Formats = formats };
    static Prop Id(string[] formats) => new("id", PropType.String, "Stable id stored in the file.") { ReadOnly = true, Formats = formats };
    static Prop Data(string description, string example) => new("data", PropType.Json, description) { Example = example };
    static Prop Len(string name, string description, string example) => new(name, PropType.Length, description) { Example = example, Formats = PptxPdf };
    static Prop[] Box => [Len("x", "Left edge.", "2cm"), Len("y", "Top edge.", "3cm"), Len("w", "Width.", "10cm"), Len("h", "Height.", "2cm")];

    /// <summary>Worksheet cell formatting, shared by the xlsx cell and range kinds.</summary>
    static Prop[] XlsxStyle =>
    [
        new("bold", PropType.Bool, "Bold.") { Example = "true", Formats = Xlsx },
        new("italic", PropType.Bool, "Italic.") { Example = "true", Formats = Xlsx },
        new("underline", PropType.Bool, "Single underline.") { Example = "true", Formats = Xlsx },
        new("strike", PropType.Bool, "Strikethrough.") { Example = "true", Formats = Xlsx },
        new("size", PropType.Points, "Font size in points.") { Example = "12pt", Formats = Xlsx },
        new("font", PropType.String, "Font name.") { Example = "Arial", Formats = Xlsx },
        new("color", PropType.Color, "Text color.") { Example = "C00000", Formats = Xlsx },
        new("fill", PropType.Color, "Background color, or none.") { Example = "D9E2F3", Formats = Office },
        new("format", PropType.String, "Number format code, e.g. 0.00, #,##0, 0%, yyyy-mm-dd.") { Example = "0.00", Formats = Xlsx },
        Align(Xlsx),
        new("valign", PropType.Enum, "Vertical alignment.") { Values = ["top", "middle", "bottom"], Example = "middle", Formats = Xlsx },
        new("wrap", PropType.Bool, "Wrap text inside the cell.") { Example = "true", Formats = Xlsx },
        new("indent", PropType.Int, "Indent level.") { Min = 0, Max = 15, Example = "2", Formats = Xlsx },
        new("border", PropType.Enum, "Border style on all four sides. Reading gives the shared style, or the most common one plus borders.")
            { Values = ["none", "thin", "medium", "thick", "dashed", "dotted", "double", "hair"], Example = "thin", Formats = Xlsx },
        new("borders", PropType.Json, "Border style per side: an object with top, right, bottom and/or left. Writing sets only the given sides.")
            { Example = "{\"top\":\"thin\",\"right\":\"thin\",\"bottom\":\"medium\",\"left\":\"thin\"}", Formats = Xlsx },
        new("borderColor", PropType.Color, "Border color, applied to every bordered side.") { Example = "FF0000", Formats = Xlsx },
    ];

    public static IReadOnlyList<Kind> Kinds { get; } =
    [
        new("document", "The whole file.")
        {
            Formats = All,
            Props =
            [
                new("format", PropType.String, "File format.") { ReadOnly = true },
                new("title", PropType.String, "Title in the file properties.") { Example = "Q4 Report", Formats = Office },
                new("page", PropType.String, "Paper: A4, Letter, Legal, A3, A5, B5, or a size like 21cm x 29.7cm.") { Example = "A4", Formats = Docx },
                new("orientation", PropType.Enum, "Page orientation.") { Values = ["portrait", "landscape"], Example = "landscape", Formats = Docx },
                new("margin", PropType.String, "Margins: narrow, normal, moderate, wide, one length, or top right bottom left.") { Example = "narrow", Formats = Docx },
                new("columns", PropType.Int, "Text columns.") { Min = 1, Max = 4, Example = "2", Formats = Docx },
                new("header", PropType.Html, "Default header as inline HTML, lines split by br or paragraphs as <p style=\"text-align:center\">; {page} and {pages} become page number fields; pictures and other content it cannot express read as <img data-keep> / <span data-keep> placeholders that keep them when written back; empty removes it.") { Example = "Q3 report", Formats = Docx },
                new("footer", PropType.Html, "Default footer, as header.") { Example = "Page {page} of {pages}", Formats = Docx },
                new("titlePg", PropType.Bool, "The first page has its own header and footer (firstHeader, firstFooter; empty unless set).") { Example = "true", Formats = Docx },
                new("firstHeader", PropType.Html, "First-page header, shown when titlePg is true; as header.") { Example = "Cover", Formats = Docx },
                new("firstFooter", PropType.Html, "First-page footer, shown when titlePg is true; as header.") { Example = "Draft", Formats = Docx },
                new("track", PropType.Bool, "Track changes: when true, text written to paragraphs is recorded as insertions and deletions.") { Example = "true", Formats = Docx },
                new("author", PropType.String, "Default author of tracked changes and comments written from here (else Writer); kept in the file as the document variable WriterAuthor, not the file's creator.") { Example = "Ann", Formats = Docx },
                new("revisions", PropType.Int, "Tracked insertions and deletions still pending.") { ReadOnly = true, Formats = Docx },
                new("comments", PropType.Int, "Comment count.") { ReadOnly = true, Formats = Docx },
                Accept, Reject,
                new("slides", PropType.Int, "Slide count.") { ReadOnly = true, Formats = Pptx },
                new("width", PropType.Length, "Slide width.") { ReadOnly = true, Formats = Pptx },
                new("height", PropType.Length, "Slide height.") { ReadOnly = true, Formats = Pptx },
                new("sheets", PropType.Int, "Sheet count.") { ReadOnly = true, Formats = Xlsx },
                new("pages", PropType.Int, "Page count.") { ReadOnly = true, Formats = Pdf },
            ],
        },
        new("body", "Main flow of block content.") { Formats = DocxMd, Singleton = true },
        new("page", "One PDF page. Children are its text blocks and images, top to bottom. Read-only.")
        {
            Formats = Pdf,
            Props =
            [
                new("width", PropType.Length, "Page width.") { ReadOnly = true },
                new("height", PropType.Length, "Page height.") { ReadOnly = true },
                new("text", PropType.String, "All text on the page.") { ReadOnly = true },
            ],
        },
        new("text", "A block of text on a PDF page. Read-only.")
        {
            Formats = Pdf,
            Props =
            [
                new("text", PropType.String, "The block's text.") { ReadOnly = true },
                new("x", PropType.Length, "Left edge.") { ReadOnly = true },
                new("y", PropType.Length, "Top edge.") { ReadOnly = true },
                new("w", PropType.Length, "Width.") { ReadOnly = true },
                new("h", PropType.Length, "Height.") { ReadOnly = true },
            ],
        },
        new("slide", "One slide. Children are its shapes, pictures and tables in z-order.")
        {
            Formats = Pptx, Parents = ["document"],
            Props =
            [
                new("title", PropType.String, "Text of the title placeholder. Writing it creates one when missing.") { Example = "Q4 Results" },
                new("layout", PropType.String, "Layout name: Title, Content, Blank, or any layout in the file.") { Example = "Content" },
                new("background", PropType.Color, "Solid background color; reading falls back to the layout's, then the master's.") { Example = "1A1A2E" },
                new("backgroundImage", PropType.Bool, "True when the effective background is a picture; its bytes are on the decor child with background=true.") { ReadOnly = true },
                new("notes", PropType.String, "Speaker notes. Writing creates the notes slide when missing; empty clears the text.") { Example = "Mention the Q3 numbers" },
                new("hidden", PropType.Bool, "Skipped in the slide show.") { Example = "true" },
                new("transition", PropType.Enum, "Transition into the slide; other = an effect the engine does not model (read-only value).") { Values = ["none", "fade", "push", "wipe", "split", "cover", "cut", "dissolve", "zoom", "random", "other"], Example = "fade" },
                new("duration", PropType.Int, "Transition length in milliseconds.") { Min = 0, Example = "700" },
                Id(Pptx),
            ],
        },
        new("sheet", "One worksheet. Address cells as /sheet[1]/cell[B3] and blocks as /sheet[1]/range[A1:C10]; rows are /sheet[1]/row[3], charts /sheet[1]/chart[1].")
        {
            Formats = Xlsx, Parents = ["document"], Summary = true,
            Props =
            [
                new("name", PropType.String, "Sheet name.") { Example = "Sales" },
                new("range", PropType.String, "Used range, e.g. A1:F20.") { ReadOnly = true },
                Id(Xlsx),
                new("merges", PropType.Json, "Merged ranges as a JSON array. Writing replaces the whole list; ranges may not overlap.") { Example = "[\"A1:C1\"]" },
                new("widths", PropType.Json, "Custom column widths in characters, column letter to width. Writing sets those columns and keeps the rest; null or \"default\" resets one.") { Example = "{\"A\":12.5,\"C\":30}" },
                new("heights", PropType.Json, "Custom row heights in points, row number to height. Writing sets those rows and keeps the rest; null or \"default\" resets one.") { Example = "{\"3\":24}" },
                new("freeze", PropType.String, "Top-left cell of the scrolling area: A2 freezes row 1, B1 column A, B2 both. none unfreezes.") { Example = "A2" },
                new("filter", PropType.String, "AutoFilter range, or none.") { Example = "A1:D20" },
            ],
        },
        new("chart", "A chart on the worksheet.")
        {
            Formats = Xlsx, Parents = ["sheet"],
            Props =
            [
                new("type", PropType.Enum, "Chart type. Changing it keeps the series.") { Values = ["column", "bar", "line", "pie", "area", "scatter", "doughnut"], Example = "column" },
                new("title", PropType.String, "Title above the chart; empty removes it.") { Example = "Sales" },
                new("series", PropType.Json, "JSON array of {name, values}: name is a cell (B1) or a text, values a range. Scatter series may add x, their X range.") { Example = "[{\"name\":\"B1\",\"values\":\"B2:B6\"}]" },
                new("categories", PropType.String, "Category labels range, on this sheet (A2:A6) or another (Sheet2!A2:A6). Scatter charts use it as the shared X range.") { Example = "A2:A6" },
                new("legend", PropType.Enum, "Legend position.") { Values = ["none", "right", "bottom", "top", "left"], Example = "bottom" },
                new("stacked", PropType.Bool, "Stack the series (column, bar, line and area charts).") { Example = "true" },
                new("x", PropType.Length, "Left edge on the sheet.") { Example = "10cm" },
                new("y", PropType.Length, "Top edge on the sheet.") { Example = "1cm" },
                new("w", PropType.Length, "Width.") { Example = "15cm" },
                new("h", PropType.Length, "Height.") { Example = "7.5cm" },
                Id(Xlsx),
            ],
        },
        new("heading", "Heading paragraph, level 1-6.")
        {
            Formats = DocxMd, Parents = ["body", "cell"],
            Props = [Text, MdProp, HtmlProp, new("level", PropType.Int, "Heading level.") { Min = 1, Max = 6, Example = "2" }, Align(Docx), PageBreakBefore, ParagraphFill, Id(Docx), Accept, Reject],
        },
        new("paragraph", "Paragraph of text. List items are paragraphs with a list property.")
        {
            Formats = Documents, Parents = ["body", "cell", "shape"],
            Props =
            [
                Text, MdProp, HtmlProp,
                new("style", PropType.String, "Paragraph style id, e.g. Normal, Quote, Title.") { Example = "Quote", Formats = Docx },
                new("list", PropType.Enum, "List marker.") { Values = ["none", "bullet", "number"], Example = "bullet" },
                new("level", PropType.Int, "List nesting level, 0 = top.") { Min = 0, Max = 8, Example = "1" },
                Align(DocxPptx), PageBreakBefore, ParagraphFill, Id(Docx), Accept, Reject,
            ],
        },
        new("code", "Code block.")
        {
            Formats = DocxMd, Parents = ["body", "cell"],
            Props = [Text with { Example = "print(1)" }, new("lang", PropType.String, "Language tag of the fence.") { Example = "python", Formats = Md }],
        },
        new("pagebreak", "Page break.") { Formats = Docx, Parents = ["body", "cell"] },
        new("toc", "Table of contents field, updated by Word on open.")
        {
            Formats = Docx, Parents = ["body"],
            Props =
            [
                new("levels", PropType.Int, "Heading levels listed, 1 up to this. Writing it regenerates the entries.") { Min = 1, Max = 9, Example = "3" },
                new("title", PropType.String, "Heading shown above the entries, e.g. Contents or 目录; empty for none.") { Example = "Contents" },
                new("text", PropType.String, "The entries as last generated, one per line.") { ReadOnly = true },
            ],
        },
        new("run", "Span of text with one set of character formatting.")
        {
            Formats = Documents, Parents = ["heading", "paragraph"], Inline = true,
            Props =
            [
                Text,
                new("bold", PropType.Bool, "Bold.") { Aliases = ["b"], Example = "true" },
                new("italic", PropType.Bool, "Italic.") { Aliases = ["i"], Example = "true" },
                new("underline", PropType.Bool, "Underline.") { Aliases = ["u"], Example = "true", Formats = DocxPptx },
                new("strike", PropType.Bool, "Strikethrough.") { Example = "true" },
                new("code", PropType.Bool, "Inline code.") { Example = "true", Formats = Md },
                new("color", PropType.Color, "Text color, or none.") { Example = "C00000", Formats = DocxPptx },
                new("size", PropType.Points, "Font size in points.") { Example = "14pt", Formats = DocxPptx },
                new("font", PropType.String, "Font family.") { Example = "Arial", Formats = DocxPptx },
                Highlight(DocxPptx),
                new("link", PropType.String, "Hyperlink target: a URL, or #bookmark.") { Example = "https://example.com" },
                new("change", PropType.Enum, "Set when the run is a tracked change. Deleted runs are left out of the paragraph text.") { Values = ["inserted", "deleted"], ReadOnly = true, Formats = Docx },
                new("author", PropType.String, "Who made the tracked change.") { ReadOnly = true, Formats = Docx },
                new("date", PropType.String, "When the tracked change was made (ISO 8601).") { ReadOnly = true, Formats = Docx },
            ],
        },
        new("comment", "A comment anchored to text in a paragraph. Address one as //comment[@id=3] or by position under its paragraph.")
        {
            Formats = Docx, Parents = ["paragraph", "heading"], Inline = true,
            Props =
            [
                Id(Docx),
                new("author", PropType.String, "Comment author.") { Example = "Ann" },
                new("initials", PropType.String, "The author's initials, as Word shows them; empty removes them.") { Example = "AL" },
                new("date", PropType.String, "When it was written (ISO 8601).") { Example = "2026-01-15T09:30:00Z" },
                new("text", PropType.String, "The comment; paragraphs joined by newlines.") { Example = "Please check this figure" },
                new("quote", PropType.String, "Text of the paragraph the comment is attached to; the whole paragraph when omitted on add. Writing it moves the anchor.") { Example = "paragraph" },
                new("resolved", PropType.Bool, "Marked as done.") { Example = "true" },
            ],
        },
        new("shape", "Text box or drawn shape on a slide. Children are its paragraphs.")
        {
            Formats = Pptx, Parents = ["slide"],
            Props =
            [
                Text, MdProp, HtmlProp,
                new("geometry", PropType.Enum, "Preset shape.") { Values = ["textbox", "rect", "roundRect", "ellipse", "triangle", "diamond", "rightArrow"], Example = "rect" },
                .. Box,
                new("fill", PropType.Color, "Fill color, or none.") { Example = "4472C4" },
                new("line", PropType.Color, "Outline color, or none.") { Example = "1F2937" },
                new("font", PropType.String, "Font for all text in the shape.") { Example = "Arial" },
                new("size", PropType.Points, "Font size for all text in the shape.") { Example = "24pt" },
                new("color", PropType.Color, "Text color for all text in the shape.") { Example = "FFFFFF" },
                new("placeholder", PropType.String, "Placeholder role: title, body, subtitle...") { ReadOnly = true },
                new("name", PropType.String, "Shape name.") { ReadOnly = true },
                Id(Pptx),
            ],
        },
        new("table", "Table.")
        {
            Formats = Documents, Parents = ["body", "cell", "slide"], Summary = true,
            Props =
            [
                new("rows", PropType.Int, "Row count. Writing it adds or removes rows at the end.") { Min = 1, Example = "3" },
                new("cols", PropType.Int, "Column count. Writing it adds or removes columns at the end.") { Min = 1, Example = "4" },
                Data("Cell texts as a JSON array of rows (visible cells; merged cells count once). Writing it resizes the table and keeps merges that still fit.", "[[\"Name\",\"Score\"],[\"Ann\",\"90\"]]"),
                .. Box,
                new("style", PropType.String, "Table style id or name, e.g. TableGrid, PlainTable1, GridTable4Accent1, ThreeLineTable (三线表); empty removes it.") { Example = "TableGrid", Formats = Docx },
                new("borders", PropType.Enum, "Which lines are drawn, overriding the style: horizontal is every horizontal line and no vertical one; style removes the override so the style's lines show.")
                    { Values = ["none", "all", "outside", "inside", "horizontal", "style"], Example = "all", Formats = Docx },
                new("borderColor", PropType.Color, "Color of the drawn lines.") { Example = "808080", Formats = Docx },
                new("width", PropType.String, "Table width: a percentage of the text width, a length, or auto.") { Example = "100%", Formats = Docx },
                new("widths", PropType.Json, "Column widths from the left as a JSON array of lengths; the table width becomes their sum.") { Example = "[\"3cm\",\"5cm\"]", Formats = Docx },
                new("align", PropType.Enum, "Table position between the margins.") { Values = ["left", "center", "right"], Example = "center", Formats = Docx },
            ],
        },
        new("row", "Table row, or a worksheet row addressed by its number.")
        {
            Formats = Tables, Parents = ["table", "sheet"],
            Props =
            [
                Data("Cell values as a JSON array, from the first column.", "[\"Ann\",\"90\"]"),
                new("header", PropType.Bool, "Repeat the row at the top of every page.") { Example = "true", Formats = Docx },
                new("height", PropType.Length, "Minimum row height; 0 removes it.") { Example = "1cm", Formats = Docx },
            ],
        },
        new("cell", "Table cell, or a worksheet cell addressed by its reference.")
        {
            Formats = Tables, Parents = ["row", "sheet"],
            Props =
            [
                Text, MdProp, HtmlProp, Align(Documents),
                new("text", PropType.String, "Displayed value.") { ReadOnly = true, Formats = Xlsx },
                new("value", PropType.String, "Value to store. Numbers, true/false and ISO dates are typed automatically.") { Example = "42", Formats = Xlsx },
                new("type", PropType.Enum, "Stored type; forces a conversion when written.") { Values = ["number", "string", "bool", "date"], Example = "number", Formats = Xlsx },
                new("formula", PropType.String, "Formula without the leading =. Results appear when the file is next recalculated.") { Example = "SUM(A1:A3)", Formats = Xlsx },
                .. XlsxStyle,
                new("link", PropType.String, "Hyperlink: a URL, or #Sheet2!A1 for a place in the workbook. Empty removes it.") { Example = "https://example.com", Formats = Xlsx },
                new("note", PropType.String, "Comment shown when the cell is hovered. Empty removes it.") { Example = "Reviewed by Ann", Formats = Xlsx },
                new("colspan", PropType.Int, "Columns the cell spans. Raising it absorbs the cells to its right (their text moves in); lowering it splits empty cells off.") { Min = 1, Example = "2", Formats = Docx },
                new("rowspan", PropType.Int, "Rows the cell spans. Raising it absorbs the cells below; lowering it splits empty cells off.") { Min = 1, Example = "2", Formats = Docx },
                new("borders", PropType.Enum, "The cell's own lines, overriding the table's.") { Values = ["none", "all", "outside"], Example = "all", Formats = Docx },
                new("valign", PropType.Enum, "Vertical alignment of the content.") { Values = ["top", "middle", "bottom"], Example = "middle", Formats = Docx },
                new("width", PropType.Length, "Width of the cell and its column.") { Example = "3cm", Formats = Docx },
            ],
        },
        new("range", "A rectangle of worksheet cells, e.g. /sheet[1]/range[A1:C10]. Formatting written here applies to every cell; reading gives what all cells share.")
        {
            Formats = Xlsx,
            Props = [new("values", PropType.Json, "Cell values as a JSON array of rows. Writing accepts JSON rows or CSV text.") { Example = "[[\"Name\",\"Score\"],[\"Ann\",90]]" }, .. XlsxStyle],
        },
        new("image", "Picture. In Word, PowerPoint and Excel files the picture tools write Office's own picture markup, so the file looks the same in Office.")
        {
            Formats = Pictures, Parents = ["body", "slide", "sheet", "cell"],
            Props =
            [
                new("src", PropType.String, "Image file to embed when writing (replaces the picture: the width stays, the height follows the new picture, the crop goes); the stored part name when reading.") { Example = "chart.png" },
                new("width", PropType.Length, "Display width.") { Example = "8cm", Formats = Docx },
                new("height", PropType.Length, "Display height.") { Example = "6cm", Formats = Docx },
                .. Box.Select(p => p with { Formats = ["pptx", "xlsx", "pdf"] }),
                new("alt", PropType.String, "Alternative text.") { Example = "Revenue chart" },
                Id(["pptx", "xlsx"]),
                new("bytes", PropType.Int, "Size of the stored picture in bytes; compare it before and after compress.") { ReadOnly = true, Formats = Office },
                new("crop", PropType.String, "Cropped off each edge, in percent of the picture: left,top,right,bottom (a:srcRect). The frame keeps the picture's scale, so it shrinks or grows with the crop; 0,0,0,0 removes it.") { Example = "10,0,10,0", Formats = Office },
                new("rotation", PropType.Int, "Clockwise rotation in degrees.") { Min = -360, Max = 360, Example = "90", Formats = Office },
                new("flipH", PropType.Bool, "Mirrored left to right.") { Example = "true", Formats = Office },
                new("flipV", PropType.Bool, "Mirrored top to bottom.") { Example = "true", Formats = Office },
                new("brightness", PropType.Int, "Brightness, -100 to 100 (a:lum bright); 0 is the picture as it is.") { Min = -100, Max = 100, Example = "20", Formats = Office },
                new("contrast", PropType.Int, "Contrast, -100 to 100 (a:lum contrast); 0 is the picture as it is.") { Min = -100, Max = 100, Example = "-10", Formats = Office },
                new("grayscale", PropType.Bool, "Shown in shades of gray (a:grayscl).") { Example = "true", Formats = Office },
                new("transparency", PropType.Int, "Transparency in percent, 0 opaque to 100 invisible (a:alphaModFix).") { Min = 0, Max = 100, Example = "30", Formats = Office },
                new("line", PropType.Color, "Border color, or none.") { Example = "1F2937", Formats = Office },
                new("lineWidth", PropType.Points, "Border width in points; a picture without a border gets a black one.") { Example = "2pt", Formats = Office },
                new("shadow", PropType.Bool, "Office's outer shadow, offset to the bottom right.") { Example = "true", Formats = Office },
                new("geometry", PropType.Enum, "Shape the picture is cropped to: roundRect gives rounded corners.") { Values = PictureShapes, Example = "roundRect", Formats = Office },
                new("compress", PropType.Enum, "Re-encodes the picture at the resolution its displayed size needs (print 220 ppi, web 150, email 96) and deletes cropped areas; a picture already small enough stays. Needs the writer-vision helper.") { Values = ["print", "web", "email"], WriteOnly = true, Example = "print", Formats = Office },
                new("background", PropType.Enum, "remove cuts the subject out with Apple Vision (macOS 14 or later, through the writer-vision helper); the original stays in the file for reset.") { Values = ["remove"], WriteOnly = true, Example = "remove", Formats = Office },
                new("reset", PropType.Bool, "true drops every adjustment (crop, rotation, colour, border, shadow, shape) and brings back the original picture; the width stays.") { WriteOnly = true, Example = "true", Formats = Office },
            ],
        },
        new("topic", "Mind map node. The first topic is the map's centre; children branch from it.")
        {
            Formats = Mm, Parents = ["document", "topic"],
            Props =
            [
                new("text", PropType.String, "Plain text of the node.") { Example = "Marketing plan" },
                new("md", PropType.Md, "Inline markdown; stored as plain text. Write-only.") { WriteOnly = true, Example = "**Q4** plan" },
                new("html", PropType.Html, "Inline HTML; stored as plain text. Write-only.") { WriteOnly = true, Example = "<b>Q4</b> plan" },
                new("note", PropType.String, "Note attached to the node; empty removes it.") { Example = "Discuss with sales" },
                new("collapsed", PropType.Bool, "Children hidden (folded).") { Example = "true" },
                new("side", PropType.Enum, "Side of the centre the branch grows on; meaningful on children of the root.") { Values = ["left", "right"], Example = "right" },
                new("link", PropType.String, "Hyperlink target; empty removes it.") { Example = "https://example.com" },
                new("color", PropType.Color, "Text color, or none.") { Example = "C00000" },
                new("fill", PropType.Color, "Background color, or none.") { Example = "FFF2CC" },
                new("icon", PropType.String, "Built-in icon name (idea, flag, button_ok, ...); none removes every icon.") { Example = "idea" },
                Id(Mm),
            ],
        },
        new("decor", "Shape, picture or background inherited from the slide's layout or master; read-only. Listed before the slide's own shapes.")
        {
            Formats = Pptx,
            Props =
            [
                new("source", PropType.Enum, "Part it comes from.") { Values = ["layout", "master", "slide"], ReadOnly = true },
                new("type", PropType.Enum, "A shape (with text) or a picture.") { Values = ["shape", "image"], ReadOnly = true },
                new("background", PropType.Bool, "True for the picture that fills the slide background.") { ReadOnly = true },
                Text with { ReadOnly = true, Formats = Pptx },
                HtmlProp with { ReadOnly = true, Formats = Pptx },
                new("geometry", PropType.String, "Preset shape.") { ReadOnly = true },
                .. Box.Select(p => p with { ReadOnly = true }),
                new("fill", PropType.Color, "Fill color, or none.") { ReadOnly = true },
                new("line", PropType.Color, "Outline color, or none.") { ReadOnly = true },
                new("font", PropType.String, "Font of the text.") { ReadOnly = true },
                new("size", PropType.Points, "Font size of the text.") { ReadOnly = true },
                new("color", PropType.Color, "Text color.") { ReadOnly = true },
                new("src", PropType.String, "Stored part name of a picture.") { ReadOnly = true },
                new("alt", PropType.String, "Alternative text of a picture.") { ReadOnly = true },
                new("name", PropType.String, "Shape name.") { ReadOnly = true },
                Id(Pptx),
            ],
        },
    ];

    public static Kind? Find(string name) => Kinds.FirstOrDefault(k => k.Name == name);

    public static IEnumerable<Kind> ForFormat(string format) => Kinds.Where(k => k.Formats.Contains(format));

    public static Kind Get(string format, string kind) =>
        Find(kind) is { } k && k.Formats.Contains(format)
            ? k
            : throw new WriterException(ErrorCode.UnsupportedKind, $"'{kind}' is not a {format} element",
                $"Elements for {format}: {string.Join(", ", ForFormat(format).Select(x => x.Name))}.");

    public static IEnumerable<Prop> PropsFor(Kind kind, string format) =>
        kind.Props.Where(p => p.Formats is null || p.Formats.Contains(format));

    public static Prop FindProp(string format, Kind kind, string name)
    {
        var props = PropsFor(kind, format).ToList();
        return props.FirstOrDefault(p => p.Name == name || p.Aliases.Contains(name))
            ?? throw new WriterException(ErrorCode.Validation, $"{kind.Name} has no property '{name}' in {format}",
                $"Writable: {string.Join(", ", props.Where(p => !p.ReadOnly).Select(p => p.Name))}. See 'writer help {format} {kind.Name}'.");
    }

    /// <summary>Validates user input and returns canonical values keyed by canonical names.</summary>
    public static Dictionary<string, string> Normalize(string format, string kind, IEnumerable<KeyValuePair<string, string>> input)
    {
        var k = Get(format, kind);
        var result = new Dictionary<string, string>();
        foreach (var (name, raw) in input)
        {
            var p = FindProp(format, k, name);
            if (p.ReadOnly)
                throw new WriterException(ErrorCode.Validation, $"{k.Name}.{p.Name} is read-only", $"See 'writer help {format} {k.Name}'.");
            result[p.Name] = NormalizeValue(p, raw);
        }
        return result;
    }

    public static string NormalizeValue(Prop p, string raw) => p.Type switch
    {
        PropType.Bool => raw.Trim().ToLowerInvariant() switch
        {
            "true" or "1" or "yes" or "on" => "true",
            "false" or "0" or "no" or "off" => "false",
            _ => throw Invalid(p, raw, "true or false"),
        },
        PropType.Int => int.TryParse(raw.Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var i) && i >= p.Min && i <= p.Max
            ? i.ToString(CultureInfo.InvariantCulture)
            : throw Invalid(p, raw, IntRange(p)),
        PropType.Length => Wrap(p, () => Units.ParseLength(raw).ToString(CultureInfo.InvariantCulture)),
        PropType.Points => Wrap(p, () => Units.ParsePoints(raw)),
        PropType.Color => Wrap(p, () => Units.ParseColor(raw)),
        PropType.Enum => p.Values!.FirstOrDefault(v => v.Equals(raw.Trim(), StringComparison.OrdinalIgnoreCase))
            ?? throw Invalid(p, raw, string.Join(" | ", p.Values!)),
        PropType.Json => ValidJson(p, raw),
        _ => raw,
    };

    static string IntRange(Prop p) => (p.Min, p.Max) switch
    {
        (int.MinValue, int.MaxValue) => "an integer",
        (_, int.MaxValue) => $"an integer of at least {p.Min}",
        _ => $"an integer from {p.Min} to {p.Max}",
    };

    static string Wrap(Prop p, Func<string> parse)
    {
        try { return parse(); }
        catch (WriterException ex) { throw new WriterException(ex.Code, $"{p.Name}: {ex.Message}", ex.Hint); }
    }

    static string ValidJson(Prop p, string raw)
    {
        if (p.Name == "values" && !raw.TrimStart().StartsWith('[')) return raw;
        try { using var _ = JsonDocument.Parse(raw); return raw; }
        catch (JsonException) { throw Invalid(p, raw, "valid JSON"); }
    }

    static WriterException Invalid(Prop p, string raw, string expected) =>
        new(ErrorCode.Validation, $"{p.Name}: '{raw}' is not valid", $"Expected {expected}. Example: {p.Name}={p.Example}");

    /// <summary>Canonical values → what users see: lengths in cm, sizes in pt. Unknown kinds pass through.</summary>
    public static Dictionary<string, string> ToDisplay(string kind, IReadOnlyDictionary<string, string> canonical) =>
        ToDisplay(null, kind, canonical);

    public static Dictionary<string, string> ToDisplay(string? format, string kind, IReadOnlyDictionary<string, string> canonical)
    {
        var definition = Find(kind);
        var props = definition is null ? null : format is null ? definition.Props : PropsFor(definition, format);
        var result = new Dictionary<string, string>();
        foreach (var (name, value) in canonical)
        {
            var type = props?.FirstOrDefault(p => p.Name == name)?.Type;
            result[name] = type switch
            {
                PropType.Length when long.TryParse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var emu) => Units.FormatLength(emu),
                PropType.Points => Units.FormatPoints(value),
                _ => value,
            };
        }
        return result;
    }

    public static bool IsJson(string kind, string prop) => IsJson(null, kind, prop);

    public static bool IsJson(string? format, string kind, string prop)
    {
        var definition = Find(kind);
        return definition is not null && (format is null ? definition.Props : PropsFor(definition, format))
            .FirstOrDefault(p => p.Name == prop)?.Type == PropType.Json;
    }

    public static void CheckParent(string format, string kind, string parentKind)
    {
        var k = Get(format, kind);
        if (!k.Parents.Contains(parentKind))
            throw new WriterException(ErrorCode.Validation, $"{kind} cannot go under {parentKind}",
                k.Parents.Length == 0 ? $"{kind} cannot be added." : $"{kind} goes under: {string.Join(", ", k.Parents)}.");
    }
}
