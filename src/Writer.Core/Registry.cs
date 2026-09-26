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
    static Prop HtmlProp => new("html", PropType.Html, "Inline HTML: b, i, u, s, code, a href, br (in Word, br style='page-break-before:always' is a page break), span style color / font-size / background-color / font-family / letter-spacing, sup / sub, and ins / del (data-author, data-date) for tracked changes. Reading gives the content as HTML; writing replaces it.") { Example = "Revenue <b>grew</b> 25%", Formats = Documents };
    static Prop Accept => new("accept", PropType.Enum, "Write all to accept every tracked change here: insertions stay, deletions go.") { Values = ["all"], WriteOnly = true, Example = "all", Formats = Docx };
    static Prop Reject => new("reject", PropType.Enum, "Write all to reject every tracked change here: insertions go, deletions come back.") { Values = ["all"], WriteOnly = true, Example = "all", Formats = Docx };
    static Prop Highlight(string[] formats) => new("highlight", PropType.Color, "Background highlight color, or none.") { Example = "FFFF00", Formats = formats };
    static Prop PageBreakBefore => new("pageBreakBefore", PropType.Bool, "Starts on a new page. Reading gives it as computed when the paragraph's style sets it.") { Example = "true", Formats = Docx };
    static Prop ParagraphFill => new("fill", PropType.Color, "Paragraph shading, or none. Reading gives the style's shading as computed.") { Example = "D9D9D9", Formats = Docx };
    static Prop Align(string[]? formats) => new("align", PropType.Enum, "Horizontal alignment.") { Values = ["left", "center", "right", "justify"], Example = "center", Formats = formats };
    /// <summary>Word's 段落 dialog, on paragraphs and headings. Lengths are written as they read: points (6pt), a length (0.5cm), or
    /// characters (2ch) for indents; none removes a setting.</summary>
    static Prop[] ParagraphFormat =>
    [
        new("align", PropType.Enum, "Horizontal alignment; distribute spreads every line, the last one too.") { Values = ["left", "center", "right", "justify", "distribute"], Example = "center", Formats = DocxPptx },
        new("lineSpacing", PropType.String, "Line spacing: a multiple (1, 1.15, 1.5, 2), an exact height (18pt) or a minimum (min 18pt); none removes it.") { Example = "1.5", Formats = Docx },
        new("spaceBefore", PropType.String, "Space before the paragraph in points (6pt), a length, or lines (0.5lines, Word's 段前 0.5 行); none removes it.") { Example = "6pt", Formats = Docx },
        new("spaceAfter", PropType.String, "Space after the paragraph in points (8pt), a length, or lines (0.5lines); none removes it.") { Example = "8pt", Formats = Docx },
        new("indentLeft", PropType.String, "Left indent: a length (1cm) or characters (2ch); none removes it.") { Example = "1cm", Formats = Docx },
        new("indentRight", PropType.String, "Right indent: a length or characters; none removes it.") { Example = "0.5cm", Formats = Docx },
        new("indentFirst", PropType.String, "First-line indent: a length or characters (2ch, Word's 首行缩进 2 字符); negative for a hanging indent (-0.74cm); none removes it.") { Example = "2ch", Formats = Docx },
        new("border", PropType.String, "Paragraph borders: the sides drawn (top bottom left right), box for all four, none for no border.") { Example = "bottom", Formats = Docx },
        new("keepNext", PropType.Bool, "Keep with the next paragraph on one page.") { Example = "true", Formats = Docx },
        new("keepLines", PropType.Bool, "Keep the paragraph's lines on one page.") { Example = "true", Formats = Docx },
        new("widowControl", PropType.String, "Widow/orphan control (孤行控制): false lets a single line sit alone at a page's top or bottom, true keeps it off (Word's default); none leaves it to the style.") { Example = "false", Formats = Docx },
        new("tabs", PropType.String, "Tab stops from the left margin, comma-separated, each a kind and a position: left 2cm, center 8cm, right 15cm, decimal 10cm; none removes them.") { Example = "left 2cm, right 15cm", Formats = Docx },
        new("bookmark", PropType.String, "A bookmark around the paragraph: its name (a letter, then letters, digits or _); cross-references and #name links go to it. none removes it.") { Example = "Results", Formats = Docx },
        new("caption", PropType.String, "A numbered caption: the label (图, 表, Figure, Table…) and a SEQ field counting the captions with that label, in front of the paragraph's text, which takes the Caption style; none takes the numbering out.") { Example = "Figure", Formats = Docx },
        new("dropCap", PropType.Enum, "The paragraph is a drop cap three lines high: drop sits in the text, margin beside it; the next paragraph flows around it. none makes it an ordinary paragraph.") { Values = ["none", "drop", "margin"], Example = "drop", Formats = Docx },
        new("sectionBreak", PropType.Enum, "The paragraph ends a section: the next one starts on the next page, right after it (continuous), or on an even or odd page; none removes the break (the text joins the following section). A new section copies the document's page setup; page, orientation, margin and columns then set this section's own.") { Values = ["none", "nextPage", "continuous", "evenPage", "oddPage"], Example = "nextPage", Formats = Docx },
        new("page", PropType.String, "Paper of the section this paragraph ends (see sectionBreak): A4, Letter, or a size like 21cm x 29.7cm.") { Example = "A4", Formats = Docx },
        new("orientation", PropType.Enum, "Orientation of the section this paragraph ends.") { Values = ["portrait", "landscape"], Example = "landscape", Formats = Docx },
        new("margin", PropType.String, "Margins of the section this paragraph ends: narrow, normal, moderate, wide, one length, or top right bottom left.") { Example = "narrow", Formats = Docx },
        new("columns", PropType.Int, "Text columns of the section this paragraph ends.") { Min = 1, Max = 4, Example = "2", Formats = Docx },
    ];
    static Prop Id(string[] formats) => new("id", PropType.String, "Stable id stored in the file.") { ReadOnly = true, Formats = formats };
    static Prop Data(string description, string example) => new("data", PropType.Json, description) { Example = example };
    static Prop Len(string name, string description, string example) => new(name, PropType.Length, description) { Example = example, Formats = PptxPdf };
    static Prop[] Box => [Len("x", "Left edge.", "2cm"), Len("y", "Top edge.", "3cm"), Len("w", "Width.", "10cm"), Len("h", "Height.", "2cm")];
    static readonly string[] Dashes = ["solid", "dash", "dot", "dashDot", "lgDash", "lgDashDot", "sysDash", "sysDot", "sysDashDot"];
    static readonly string[] LineEnds = ["none", "triangle", "arrow", "stealth", "diamond", "oval"];
    /// <summary>The outline of a shape or connector beyond its colour.</summary>
    static Prop[] Outline =>
    [
        new("lineWidth", PropType.Points, "Outline width in points.") { Example = "2pt", Formats = Pptx },
        new("dash", PropType.Enum, "Outline dash pattern.") { Values = Dashes, Example = "dash", Formats = Pptx },
    ];

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
        new("rotate", PropType.Int, "Text rotation in degrees, counter-clockwise; negative turns clockwise, 0 is level.") { Min = -90, Max = 90, Example = "45", Formats = Xlsx },
        new("border", PropType.Enum, "Border style on all four sides. Reading gives the shared style, or the most common one plus borders.")
            { Values = ["none", "thin", "medium", "thick", "dashed", "dotted", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"], Example = "thin", Formats = Xlsx },
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
                new("lineNumbers", PropType.Bool, "Line numbers in the margin, counted through the document.") { Example = "true", Formats = Docx },
                new("noteFormat", PropType.String, "How footnotes and endnotes are numbered: decimal (1, 2, 3), lowerRoman, upperRoman, lowerLetter, upperLetter, decimalEnclosedCircleChinese (①②③) or chineseCounting (一二三); none leaves Word's own (1, 2, 3 and i, ii, iii).") { Example = "lowerRoman", Formats = Docx },
                new("hyphenation", PropType.Bool, "Automatic hyphenation.") { Example = "true", Formats = Docx },
                new("track", PropType.Bool, "Track changes: when true, text written to paragraphs is recorded as insertions and deletions.") { Example = "true", Formats = Docx },
                new("author", PropType.String, "Default author of tracked changes and comments written from here (else Writer); kept in the file as the document variable WriterAuthor, not the file's creator.") { Example = "Ann", Formats = Docx },
                new("revisions", PropType.Int, "Tracked insertions and deletions still pending.") { ReadOnly = true, Formats = Docx },
                new("comments", PropType.Int, "Comment count.") { ReadOnly = true, Formats = Docx },
                new("styles", PropType.Json, "The document's paragraph and character styles as a JSON array of {id, name, type, basedOn, heading, look}: look holds what the style sets on its own (font, size, bold, italic, color, align, spaceBefore, spaceAfter, lineSpacing, indentFirst), inherited along basedOn.") { ReadOnly = true, Formats = Docx },
                new("style", PropType.Json, "Write-only. Defines or changes a style: a JSON object with id (or name) and any of name, type (paragraph, the default, or character), basedOn, font, size (points), bold, italic, color, align, spaceBefore, spaceAfter, lineSpacing, indentFirst; a value of none clears that setting. A new style is added to styles.xml; paragraphs then take it with style=<id>, runs with style=<id>.") { WriteOnly = true, Example = "{\"id\":\"Note\",\"name\":\"Note\",\"italic\":true,\"color\":\"595959\"}", Formats = Docx },
                Accept, Reject,
                new("slides", PropType.Int, "Slide count.") { ReadOnly = true, Formats = Pptx },
                new("palette", PropType.Enum, "Colour palette and fonts for the whole deck, written into its theme: paper (white, black accent), ink (dark, gold), sea (deep blue, cyan), clay (terracotta on cream), mist (cool grey, blue), sand (warm beige, amber), rose (white, red accent), night (navy, sky blue). Text, backgrounds and shapes that take theme colours follow; colours set on a shape stay. Reading gives the palette when the deck wears one.") { Values = ["paper", "ink", "sea", "clay", "mist", "sand", "rose", "night"], Example = "sea", Formats = Pptx },
                new("fonts", PropType.String, "Heading font, body font: the theme fonts. Writing it also takes fonts set on the text itself off every slide, so the whole deck uses these two.") { Example = "Noto Serif SC,Noto Sans SC", Formats = Pptx },
                new("width", PropType.Length, "Slide width.") { ReadOnly = true, Formats = Pptx },
                new("height", PropType.Length, "Slide height.") { ReadOnly = true, Formats = Pptx },
                new("sheets", PropType.Int, "Sheet count.") { ReadOnly = true, Formats = Xlsx },
                new("font", PropType.String, "The workbook's default font, which cells without a font of their own show.") { ReadOnly = true, Formats = Xlsx },
                new("size", PropType.Points, "The default font's size.") { ReadOnly = true, Formats = Xlsx },
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
                new("layout", PropType.String, "Layout: title, content, section, two, comparison, titleOnly, blank, caption (content with caption), picture (picture with caption), quote, or any layout name in the file; a standard one the file lacks is added. Changing it moves the slide's placeholders as PowerPoint does: each takes the place of the new layout's placeholder of the same kind (title, text, picture), keeping its text; an empty one left over goes; the new layout's other placeholders are added empty. Reading gives the layout's name.") { Example = "Two Content" },
                new("align", PropType.String, "Write-only. Lines shapes up: left, center, right, top, middle or bottom, a colon, then the slide's shapes, pictures or tables by path, comma-separated. One lines up with the slide, several with the box around them.") { WriteOnly = true, Example = "left:/slide[2]/shape[3],/slide[2]/shape[4]" },
                new("distribute", PropType.String, "Write-only. Spaces shapes evenly: horizontal or vertical, a colon, then the paths. Three or more keep the outer two in place and even out the gaps; one or two are spread across the slide.") { WriteOnly = true, Example = "horizontal:/slide[2]/shape[2],/slide[2]/shape[3],/slide[2]/shape[4]" },
                new("background", PropType.Color, "Solid background color; reading falls back to the layout's, then the master's.") { Example = "1A1A2E" },
                new("backgroundImage", PropType.Bool, "True when the effective background is a picture; its bytes are on the decor child with background=true.") { ReadOnly = true },
                new("notes", PropType.String, "Speaker notes. Writing creates the notes slide when missing; empty clears the text.") { Example = "Mention the Q3 numbers" },
                new("hidden", PropType.Bool, "Skipped in the slide show.") { Example = "true" },
                new("section", PropType.String, "Name of the section (节) this slide starts. Writing a name starts one here, or renames it; empty removes it, its slides joining the section before. A section runs to the next one's first slide; the first section made after slide 1 gets a Default Section in front.") { Example = "Results" },
                new("transition", PropType.Enum, "Transition into the slide; morph moves the shapes the slide shares with the previous one into place (PowerPoint 2019+, a fade elsewhere); other = an effect the engine does not model (read-only value).") { Values = ["none", "fade", "push", "wipe", "split", "cover", "cut", "dissolve", "zoom", "random", "morph", "other"], Example = "fade" },
                new("duration", PropType.Int, "Transition length in milliseconds.") { Min = 0, Example = "700" },
                new("animations", PropType.Json, "The slide's animations in the order they play, as a JSON array; writing replaces them all. Each: shape (its id), effect (entrance appear, fade, fly, float, zoom, wipe; emphasis grow, spin, transparency; exit disappear, fadeOut, flyOut, zoomOut, wipeOut), start (click, with or after the previous), duration and delay in ms. An effect the engine does not model reads as effect other with its class and xml; pass it back unchanged to keep it. Removing a shape removes its effects.") { Example = "[{\"shape\":\"4\",\"effect\":\"fade\",\"start\":\"click\",\"duration\":500,\"delay\":0}]" },
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
                new("gridlines", PropType.Bool, "Grid lines shown on screen; reading gives false only when the sheet hides them.") { Example = "false" },
                new("filter", PropType.String, "AutoFilter range, or none.") { Example = "A1:D20" },
                new("filters", PropType.Json, "AutoFilter criteria by column letter: {values:[…]} keeps only those texts (\"\" for blanks); {operator, value, value2} keeps what the comparison admits (equal, notEqual, greaterThan, greaterThanOrEqual, lessThan, lessThanOrEqual, between, contains, notContains, beginsWith, endsWith). Writing replaces the whole set; {} clears it. Excel applies the criteria when asked to reapply.")
                    { Example = "{\"A\":{\"values\":[\"East\",\"West\"]},\"C\":{\"operator\":\"greaterThan\",\"value\":\"100\"}}" },
                new("cf", PropType.Json, "Conditional formatting rules, first wins, as a JSON array of {range, type, …}: cellIs (operator equal, notEqual, greaterThan, greaterThanOrEqual, lessThan, lessThanOrEqual, between, notBetween; value, value2), containsText / notContainsText / beginsWith / endsWith (text), duplicateValues, uniqueValues, top10 (rank, percent, bottom), aboveAverage (above:false for below, stdDev), dataBar (color), colorScale (colors: 2 or 3), iconSet (iconSet, reverse), timePeriod (period), expression (value = formula). A rule's look is its fill, color, bold and italic. Writing replaces the whole list.")
                    { Example = "[{\"range\":\"A2:A11\",\"type\":\"cellIs\",\"operator\":\"greaterThan\",\"value\":\"80\",\"fill\":\"C6EFCE\",\"color\":\"006100\"}]" },
                new("validations", PropType.Json, "Data validation rules as a JSON array of {range, type, …}: list (values, or source: a range), whole / decimal / date / time / textLength (operator as cf's cellIs, value, value2; dates and times as serial numbers), custom (value = formula); optional error, errorTitle, prompt, promptTitle, allowBlank:false. Writing replaces the whole list.")
                    { Example = "[{\"range\":\"A2:A20\",\"type\":\"list\",\"values\":[\"Yes\",\"No\"],\"error\":\"Pick Yes or No\"}]" },
                new("hidden", PropType.Json, "Hidden rows and columns: {rows:[3,4], cols:[\"B\"]}. Writing replaces the whole set.") { Example = "{\"rows\":[3],\"cols\":[\"B\"]}" },
                new("color", PropType.Color, "Tab color, or none.") { Example = "C0392B" },
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
            Props = [Text, MdProp, HtmlProp, new("level", PropType.Int, "Heading level.") { Min = 1, Max = 6, Example = "2" }, .. ParagraphFormat.Select(p => p.Name == "align" ? p with { Formats = Docx } : p), PageBreakBefore, ParagraphFill, Id(Docx), Accept, Reject],
        },
        new("paragraph", "Paragraph of text. List items are paragraphs with a list property.")
        {
            Formats = Documents, Parents = ["body", "cell", "shape"],
            Props =
            [
                Text, MdProp, HtmlProp,
                new("style", PropType.String, "Paragraph style id, e.g. Normal, Quote, Title.") { Example = "Quote", Formats = Docx },
                new("list", PropType.Enum, "List marker: bullet (• ◦ ▪ by level), number (1. a. i.), outline (1. 1.1 1.1.1) or chinese (一、（一）1.); a numbered item continues the list of the item before it, else starts at 1.") { Values = ["none", "bullet", "number", "outline", "chinese"], Example = "bullet" },
                new("level", PropType.Int, "List nesting level, 0 = top.") { Min = 0, Max = 8, Example = "1" },
                new("restart", PropType.Bool, "Numbering starts again at 1 here (true), or continues the list of the item before (false); this item and the ones after it in its list. Reading gives true where the numbering starts again after another list of the same kind.") { Example = "true", Formats = Docx },
                .. ParagraphFormat, PageBreakBefore, ParagraphFill, Id(Docx), Accept, Reject,
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
                new("style", PropType.Enum, "How the entries end: classic draws dots to the page number, simple leaves a gap before it, plain has no page numbers. Writing it regenerates the entries.") { Values = ["classic", "simple", "plain"], Example = "simple" },
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
                new("style", PropType.String, "Character style id or name (Emphasis, Strong, Intense Emphasis, Subtle Emphasis, or one the document defines); none removes it. In html a run carries it as <span data-style=\"Emphasis\">.") { Example = "Emphasis", Formats = Docx },
                new("vertAlign", PropType.Enum, "Superscript or subscript (sup / sub in html).") { Values = ["baseline", "superscript", "subscript"], Example = "superscript", Formats = DocxPptx },
                new("spacing", PropType.String, "Character spacing in points, expanded (2pt) or condensed (-0.5pt); none removes it. In html: letter-spacing.") { Example = "1pt", Formats = DocxPptx },
                new("caps", PropType.Enum, "All capitals or small capitals (in html text-transform:uppercase / font-variant:small-caps).") { Values = ["none", "all", "small"], Example = "all", Formats = Pptx },
                new("underlineStyle", PropType.Enum, "The underline's line; writing it underlines the run (in html text-decoration-style on u).") { Values = ["single", "double", "dotted", "dashed", "wavy"], Example = "double", Formats = Pptx },
                new("outline", PropType.Bool, "Outlined letters (Word's text effect; in html -webkit-text-stroke).") { Example = "true", Formats = Docx },
                new("shadow", PropType.Bool, "Shadowed letters (Word's text effect; in html text-shadow).") { Example = "true", Formats = Docx },
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
                new("parent", PropType.String, "The id of the comment this one replies to (a reply added with it shares that comment's place); none makes it a comment of its own. Removing a comment removes its replies.") { Example = "0" },
            ],
        },
        new("footnote", "A footnote or endnote: its mark sits in the paragraph at a character offset, its text in footnotes.xml or endnotes.xml. Word numbers them by order. Address one as //footnote[@id=3] or by position under its paragraph.")
        {
            Formats = Docx, Parents = ["paragraph", "heading"], Inline = true,
            Props =
            [
                Id(Docx),
                new("kind", PropType.Enum, "footnote (at the foot of the page) or endnote (at the end of the document); set when added.") { Values = ["footnote", "endnote"], Example = "footnote" },
                new("text", PropType.String, "The note's text; paragraphs joined by newlines.") { Example = "See the appendix." },
                new("at", PropType.Int, "Character offset in the paragraph's text where the mark sits; the end when omitted on add. Writing it moves the mark.") { Min = 0, Example = "12" },
            ],
        },
        new("equation", "An equation in a paragraph (Word's Office Math), written and read as LaTeX: fractions, scripts, roots, sums and integrals with limits, \\left…\\right, matrices and cases, accents, Greek letters and symbols. It sits at a character offset of the paragraph's text.")
        {
            Formats = Docx, Parents = ["paragraph", "heading"], Inline = true,
            Props =
            [
                new("latex", PropType.String, "The equation as LaTeX.") { Example = "\\frac{a}{b}+x^{2}" },
                new("display", PropType.Bool, "On a line of its own (display math), else in the line of text.") { Example = "true" },
                new("at", PropType.Int, "Character offset in the paragraph's text where it sits; the end when omitted on add. Writing it moves the equation.") { Min = 0, Example = "4" },
            ],
        },
        new("shape", "Text box or drawn shape: on a slide, where its children are its paragraphs; in Word, floating in a paragraph (Word's wps shapes), placed like a floating picture, with plain text. Address a Word shape as //shape[@id=3].")
        {
            Formats = ["pptx", "docx"], Parents = ["slide", "paragraph", "heading"],
            Props =
            [
                Text, MdProp with { Formats = Pptx }, HtmlProp with { Formats = Pptx },
                new("geometry", PropType.String, "Preset shape: textbox, rect, roundRect, ellipse, triangle, rtTriangle, diamond, parallelogram, trapezoid, pentagon, hexagon, octagon, star4, star5, star6, rightArrow, leftArrow, upArrow, downArrow, leftRightArrow, chevron, wedgeRectCallout, wedgeRoundRectCallout, wedgeEllipseCallout, or any other DrawingML preset.") { Example = "rect" },
                .. Box,
                new("rotation", PropType.Int, "Clockwise rotation in degrees.") { Min = -360, Max = 360, Example = "90", Formats = Pptx },
                new("fill", PropType.Color, "Fill color, or none.") { Example = "4472C4" },
                new("gradient", PropType.String, "Two-colour gradient fill: start colour, end colour, angle in degrees (90 = top to bottom); empty removes it.") { Example = "4472C4,ED7D31,90", Formats = Pptx },
                new("line", PropType.Color, "Outline color, or none.") { Example = "1F2937" },
                .. Outline,
                new("shadow", PropType.Bool, "Office's outer shadow, offset downwards.") { Example = "true", Formats = Pptx },
                new("lockAspect", PropType.Bool, "Resizing keeps the proportions.") { Example = "true", Formats = Pptx },
                new("font", PropType.String, "Font for all text in the shape.") { Example = "Arial", Formats = Pptx },
                new("size", PropType.Points, "Font size for all text in the shape.") { Example = "24pt", Formats = Pptx },
                new("color", PropType.Color, "Text color for all text in the shape.") { Example = "FFFFFF", Formats = Pptx },
                new("lineSpacing", PropType.String, "Line spacing of every paragraph as a multiple of single; 1 removes it.") { Example = "1.5", Formats = Pptx },
                new("spaceBefore", PropType.String, "Space before every paragraph, in points; 0 removes it.") { Example = "6pt", Formats = Pptx },
                new("spaceAfter", PropType.String, "Space after every paragraph, in points; 0 removes it.") { Example = "6pt", Formats = Pptx },
                new("charSpacing", PropType.String, "Character spacing of all text, in points; negative condenses, 0 removes it.") { Example = "1.5", Formats = Pptx },
                new("columns", PropType.Int, "Text columns in the box.") { Min = 1, Max = 16, Example = "2", Formats = Pptx },
                new("autofit", PropType.String, "How text and box meet: none, shrink (text shrinks to fit; shrink:85 records the scale PowerPoint shows it at) or resize (the box grows with the text).") { Example = "shrink", Formats = Pptx },
                new("direction", PropType.Enum, "Text direction: horz, vert (rotated 90°), vert270, eaVert (upright East Asian vertical), wordArtVert (stacked).") { Values = ["horz", "vert", "vert270", "eaVert", "wordArtVert"], Example = "eaVert", Formats = Pptx },
                new("textOutline", PropType.Color, "Outline around the letters, or none.") { Example = "1F2937", Formats = Pptx },
                new("textShadow", PropType.Bool, "Shadow behind the letters.") { Example = "true", Formats = Pptx },
                new("textGradient", PropType.String, "Gradient fill of the letters: start colour, end colour, angle; empty removes it and the letters take the first colour.") { Example = "4472C4,ED7D31,90", Formats = Pptx },
                new("placeholder", PropType.String, "Placeholder role: title, body, subtitle, pic... An empty placeholder shows a prompt in the editor and in PowerPoint; fill it rather than adding a text box.") { ReadOnly = true, Formats = Pptx },
                new("fit", PropType.Enum, "Write-only. shrink: lowers every text size in the shape by the same factor (down to half) until the text fits the box.") { Values = ["shrink"], WriteOnly = true, Example = "shrink", Formats = Pptx },
                new("overflow", PropType.Bool, "True when the text is estimated to need more room than the box has: shorten it, enlarge the box, or fit=shrink.") { ReadOnly = true, Formats = Pptx },
                new("name", PropType.String, "Shape name.") { ReadOnly = true, Formats = Pptx },
                Id(Pptx),
                new("width", PropType.Length, "Width.") { Example = "5cm", Formats = Docx },
                new("height", PropType.Length, "Height.") { Example = "2.5cm", Formats = Docx },
                new("wrap", PropType.Enum, "How the text flows around it: square, tight, through and topBottom beside or around it; front floats over the text, behind under it.") { Values = ["square", "tight", "through", "topBottom", "front", "behind"], Example = "front", Formats = Docx },
                new("x", PropType.Length, "Left edge from xFrom. Writing it replaces xAlign.") { Example = "2cm", Formats = Docx },
                new("y", PropType.Length, "Top edge from yFrom. Writing it replaces yAlign.") { Example = "1cm", Formats = Docx },
                new("xFrom", PropType.Enum, "What x and xAlign measure from.") { Values = ["page", "margin", "column", "character", "leftMargin", "rightMargin", "insideMargin", "outsideMargin"], Example = "column", Formats = Docx },
                new("yFrom", PropType.Enum, "What y and yAlign measure from; paragraph is the paragraph the shape is in.") { Values = ["page", "margin", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin", "outsideMargin"], Example = "paragraph", Formats = Docx },
                new("xAlign", PropType.Enum, "Aligned in xFrom instead of placed at x.") { Values = ["left", "center", "right", "inside", "outside"], Example = "center", Formats = Docx },
                new("yAlign", PropType.Enum, "Aligned in yFrom instead of placed at y.") { Values = ["top", "center", "bottom", "inside", "outside"], Example = "top", Formats = Docx },
                Id(Docx),
            ],
        },
        new("connector", "Line or connector on a slide: straight or bent, with arrowheads, sticking to shapes when start / end name them. The box runs from the start point to the end point; flipH / flipV turn it for lines that go up or left.")
        {
            Formats = Pptx, Parents = ["slide"],
            Props =
            [
                new("geometry", PropType.Enum, "Connector kind.") { Values = ["line", "straightConnector1", "bentConnector3", "curvedConnector3"], Example = "straightConnector1" },
                .. Box,
                new("rotation", PropType.Int, "Clockwise rotation in degrees.") { Min = -360, Max = 360, Example = "90" },
                new("flipH", PropType.Bool, "Runs right to left.") { Example = "true" },
                new("flipV", PropType.Bool, "Runs bottom to top.") { Example = "true" },
                new("line", PropType.Color, "Line color, or none.") { Example = "1F2937" },
                .. Outline,
                new("head", PropType.Enum, "Arrowhead at the start.") { Values = LineEnds, Example = "triangle" },
                new("tail", PropType.Enum, "Arrowhead at the end.") { Values = LineEnds, Example = "triangle" },
                new("start", PropType.String, "Shape the start sticks to: its id and connection site (0 top, 1 left, 2 bottom, 3 right); empty lets go.") { Example = "2,3" },
                new("end", PropType.String, "Shape the end sticks to, as start.") { Example = "2,1" },
                new("shadow", PropType.Bool, "Office's outer shadow.") { Example = "true" },
                new("name", PropType.String, "Shape name.") { ReadOnly = true },
                Id(Pptx),
            ],
        },
        new("group", "Grouped shapes on a slide. Children are its members, whose boxes read in slide units; moving or resizing the group moves them along. Add one with members (the slide's shapes by path); ungroup=true puts them back on the slide.")
        {
            Formats = Pptx, Parents = ["slide"],
            Props =
            [
                new("members", PropType.String, "Write-only, when adding: the shapes to group, comma-separated paths (shape[2] relative to the slide, or from the outline).") { WriteOnly = true, Example = "shape[2],shape[3]" },
                .. Box,
                new("rotation", PropType.Int, "Clockwise rotation in degrees.") { Min = -360, Max = 360, Example = "90" },
                new("ungroup", PropType.Bool, "Write-only. true dissolves the group, leaving its members on the slide where they were shown.") { WriteOnly = true, Example = "true" },
                new("name", PropType.String, "Group name.") { ReadOnly = true },
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
                new("style", PropType.String, "Table style id or name: in Word e.g. TableGrid, PlainTable1, GridTable4Accent1, ThreeLineTable (三线表); in PowerPoint MediumStyle2Accent1, LightStyle1, LightStyle2Accent1, DarkStyle1, TableGrid, NoStyle or a style GUID; empty removes it.") { Example = "TableGrid", Formats = DocxPptx },
                new("header", PropType.Bool, "The first row is a header row (the style draws it so; Word's 标题行).") { Example = "true", Formats = DocxPptx },
                new("banded", PropType.Bool, "Banded rows.") { Example = "true", Formats = Pptx },
                new("firstCol", PropType.Bool, "The first column stands out.") { Example = "true", Formats = Pptx },
                new("total", PropType.Bool, "The last row is a total row.") { Example = "true", Formats = Pptx },
                new("borders", PropType.Enum, "Which lines are drawn, overriding the style: horizontal is every horizontal line and no vertical one; style removes the override so the style's lines show.")
                    { Values = ["none", "all", "outside", "inside", "horizontal", "style"], Example = "all", Formats = Docx },
                new("borderColor", PropType.Color, "Color of the drawn lines.") { Example = "808080", Formats = Docx },
                new("width", PropType.String, "Table width: a percentage of the text width, a length, or auto.") { Example = "100%", Formats = Docx },
                new("widths", PropType.Json, "Column widths from the left as a JSON array of lengths; the table width becomes their sum.") { Example = "[\"3cm\",\"5cm\"]", Formats = DocxPptx },
                new("align", PropType.Enum, "Table position between the margins.") { Values = ["left", "center", "right"], Example = "center", Formats = Docx },
                new("name", PropType.String, "Frame name.") { ReadOnly = true, Formats = Pptx },
                Id(Pptx),
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
                new("colspan", PropType.Int, "Columns the cell spans. Raising it absorbs the cells to its right (in Word their text moves in); lowering it splits empty cells off.") { Min = 1, Example = "2", Formats = DocxPptx },
                new("rowspan", PropType.Int, "Rows the cell spans. Raising it absorbs the cells below; lowering it splits empty cells off.") { Min = 1, Example = "2", Formats = DocxPptx },
                new("covered", PropType.Bool, "True for a cell hidden under a merge (PowerPoint keeps it in the grid).") { ReadOnly = true, Formats = Pptx },
                new("line", PropType.Color, "The cell's own line colour on all four sides, or none.") { Example = "1F2937", Formats = Pptx },
                new("borders", PropType.String, "The cell's own lines, overriding the table's: none, all, or the sides to draw (top bottom left right, space-separated).") { Example = "all", Formats = Docx },
                new("valign", PropType.Enum, "Vertical alignment of the content.") { Values = ["top", "middle", "bottom"], Example = "middle", Formats = Docx },
                new("width", PropType.Length, "Width of the cell and its column.") { Example = "3cm", Formats = Docx },
            ],
        },
        new("range", "A rectangle of worksheet cells, e.g. /sheet[1]/range[A1:C10]. Formatting written here applies to every cell; reading gives what all cells share.")
        {
            Formats = Xlsx,
            Props =
            [
                new("values", PropType.Json, "Cell values as a JSON array of rows. Writing accepts JSON rows or CSV text.") { Example = "[[\"Name\",\"Score\"],[\"Ann\",90]]" },
                new("formula", PropType.String, "Write-only. Fills the formula into every cell of the range as Excel's fill handle does: written for the top-left cell, relative references shift with each row and column, $ pins one.") { WriteOnly = true, Example = "C2*D2", Formats = Xlsx },
                new("sort", PropType.String, "Write-only. Sorts the range's rows by one of its columns: the column letter, optionally :desc. Cells move with their formatting and formulas; blanks go last. Give the data rows only, without the header row.") { WriteOnly = true, Example = "B:desc", Formats = Xlsx },
                .. XlsxStyle,
            ],
        },
        new("image", "Picture. In Word, PowerPoint and Excel files the picture tools write Office's own picture markup, so the file looks the same in Office.")
        {
            Formats = Pictures, Parents = ["body", "slide", "sheet", "cell", "paragraph", "heading"],
            Props =
            [
                new("src", PropType.String, "Image file to embed when writing (replaces the picture: the width stays, the height follows the new picture, the crop goes); the stored part name when reading.") { Example = "chart.png" },
                new("width", PropType.Length, "Display width.") { Example = "8cm", Formats = Docx },
                new("height", PropType.Length, "Display height.") { Example = "6cm", Formats = Docx },
                new("wrap", PropType.Enum, "How text flows around it: inline sits in the line of text (a picture of its own under the body); square, tight, through and topBottom float with the text around them; front floats over the text, behind under it. A picture added or moved under a paragraph floats there (square unless given); moved to the body it is inline again.")
                    { Values = ["inline", "square", "tight", "through", "topBottom", "front", "behind"], Example = "square", Formats = Docx },
                new("x", PropType.Length, "Floating: left edge from xFrom (Word's positionH offset). Writing it replaces xAlign.") { Example = "2cm", Formats = Docx },
                new("y", PropType.Length, "Floating: top edge from yFrom (Word's positionV offset). Writing it replaces yAlign.") { Example = "1cm", Formats = Docx },
                new("xFrom", PropType.Enum, "Floating: what x and xAlign measure from.") { Values = ["page", "margin", "column", "character", "leftMargin", "rightMargin", "insideMargin", "outsideMargin"], Example = "column", Formats = Docx },
                new("yFrom", PropType.Enum, "Floating: what y and yAlign measure from; paragraph is the paragraph the picture is in.") { Values = ["page", "margin", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin", "outsideMargin"], Example = "paragraph", Formats = Docx },
                new("xAlign", PropType.Enum, "Floating: aligned in xFrom instead of placed at x.") { Values = ["left", "center", "right", "inside", "outside"], Example = "center", Formats = Docx },
                new("yAlign", PropType.Enum, "Floating: aligned in yFrom instead of placed at y.") { Values = ["top", "center", "bottom", "inside", "outside"], Example = "top", Formats = Docx },
                .. Box.Select(p => p with { Formats = ["pptx", "xlsx", "pdf"] }),
                new("alt", PropType.String, "Alternative text.") { Example = "Revenue chart" },
                Id(["docx", "pptx", "xlsx"]),
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
                new("icon", PropType.String, "Icons and markers as a comma-separated list of built-in names (idea, flag, full-1…full-9 priorities, 0%…100% progress, flag-blue, star, ...); none removes every icon.") { Example = "full-1,flag" },
                new("labels", PropType.String, "Labels shown under the text, comma-separated; empty removes them. Kept as FreeMind attribute rows called label.") { Example = "urgent,Q4" },
                new("image", PropType.String, "Picture shown above the text: a data: URI or a file URI (Freeplane's external-object hook); empty removes it.") { Example = "data:image/png;base64,..." },
                new("imageSize", PropType.String, "Natural size of the picture as w,h in pixels; the topic scales it to its width.") { Example = "320,240" },
                new("bold", PropType.Bool, "Bold text.") { Example = "true" },
                new("italic", PropType.Bool, "Italic text.") { Example = "true" },
                new("strike", PropType.Bool, "Struck-through text.") { Example = "true" },
                new("size", PropType.Int, "Font size in points; 0 goes back to the level's default.") { Min = 0, Max = 200, Example = "16" },
                new("font", PropType.String, "Font family name; empty goes back to the default.") { Example = "Georgia" },
                new("free", PropType.String, "Floating topic (a child of the root drawn where it was dropped): its top-left as x,y in pixels from the centre; empty attaches it again. Kept as a FreeMind attribute row.") { Example = "240,-160" },
                new("cloud", PropType.Color, "Boundary drawn around this topic and its branch (FreeMind's cloud) in this color, or none.") { Example = "CFE2F3" },
                new("summary", PropType.String, "Makes this topic the summary of a run of its siblings, firstId:lastId; it is drawn beside a brace spanning them. Empty makes it an ordinary child. Kept as a FreeMind attribute row.") { Example = "ID_3:ID_5" },
                new("rels", PropType.Json, "Relationship lines from this topic: a JSON array of {to: target id, label, color, arrows: end | both | start | none}. Replaces all of them; [] removes them. Stored as FreeMind arrowlinks.") { Example = "[{\"to\":\"ID_7\",\"label\":\"leads to\",\"color\":\"\",\"arrows\":\"end\"}]" },
                new("structure", PropType.Enum, "How this topic's branch is drawn (on the centre: the whole map): map two-sided, logic / logic-left one-sided, org an org chart downwards, tree an indented tree, timeline milestones along an axis (centre only); empty inherits the parent's.") { Values = ["map", "logic", "logic-left", "org", "tree", "timeline", ""], Example = "org" },
                new("theme", PropType.Enum, "Colour theme of the map, on the centre: classic, ocean, sunset, forest, candy or ink; empty is classic.") { Values = ["classic", "ocean", "sunset", "forest", "candy", "ink", ""], Example = "ocean" },
                new("lines", PropType.Enum, "Branch line style of the map, on the centre: curve (default), straight or elbow.") { Values = ["curve", "straight", "elbow", ""], Example = "elbow" },
                new("mono", PropType.Bool, "true draws every branch in the theme's first colour instead of the rainbow; on the centre.") { Example = "true" },
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
