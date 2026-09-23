using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>The blank document `create` starts from: A4, 2.54cm margins, a small set of named styles.</summary>
public static class DocxTemplate
{
    public const string Ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static void WriteBlank(Stream stream)
    {
        using var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = package.AddMainDocumentPart();
        main.Document = new W.Document(new W.Body(new W.SectionProperties(
            new W.PageSize { Width = 11906U, Height = 16838U },
            new W.PageMargin { Top = 1440, Right = 1440U, Bottom = 1440, Left = 1440U, Header = 720U, Footer = 720U, Gutter = 0U })));
        main.AddNewPart<StyleDefinitionsPart>().Styles = new W.Styles(StylesXml());
        main.AddNewPart<DocumentSettingsPart>().Settings = new W.Settings(
            $"""<w:settings xmlns:w="{Ns}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>""");
    }

    /// <summary>The XML of one built-in style, with its namespace declared, so it can be added to documents that lack it. Null for unknown ids.</summary>
    public static string? StyleXml(string styleId) =>
        Styles.FirstOrDefault(s => s.Id == styleId).Xml?.Replace("<w:style ", $"<w:style xmlns:w=\"{Ns}\" ", StringComparison.Ordinal);

    public static IEnumerable<string> StyleIds => Styles.Select(s => s.Id);

    static string StylesXml() =>
        $"""
        <w:styles xmlns:w="{Ns}">
        <w:docDefaults>
        <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault>
        <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
        </w:docDefaults>
        {string.Concat(Styles.Select(s => s.Xml))}
        </w:styles>
        """;

    static readonly (string Id, string Xml)[] Styles =
    [
        ("Normal", """<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>"""),
        ("Title", """<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>"""),
        ("Heading1", Heading(1, 32)),
        ("Heading2", Heading(2, 28)),
        ("Heading3", Heading(3, 26)),
        ("Heading4", Heading(4, 24)),
        ("Heading5", Heading(5, 22)),
        ("Heading6", Heading(6, 22)),
        ("Code", """<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>"""),
        ("Quote", """<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>"""),
        ("ListParagraph", """<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>"""),
        ("DefaultParagraphFont", """<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>"""),
        ("Hyperlink", """<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>"""),
        ("TableNormal", """<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>"""),
        ("TableGrid", """<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>"""),
        ("PlainTable1", """<w:style w:type="table" w:styleId="PlainTable1"><w:name w:val="Plain Table 1"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr><w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:tcPr></w:tblStylePr></w:style>"""),
        ("GridTable4Accent1", """<w:style w:type="table" w:styleId="GridTable4Accent1"><w:name w:val="Grid Table 4 Accent 1"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:left w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:right w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/></w:tblBorders></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:left w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:right w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr></w:style>"""),
        ("ThreeLineTable", """<w:style w:type="table" w:styleId="ThreeLineTable"><w:name w:val="Three Line Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="59"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="12" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblStylePr w:type="firstRow"><w:tblPr/><w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="auto"/></w:tcBorders></w:tcPr></w:tblStylePr></w:style>"""),
        ("TOCHeading","""<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>"""),
        .. Enumerable.Range(1, 9).Select(level => ($"TOC{level}", Toc(level))),
    ];

    static string Toc(int level) =>
        $"""<w:style w:type="paragraph" w:styleId="TOC{level}"><w:name w:val="toc {level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/>{(level > 1 ? $"""<w:ind w:left="{220 * (level - 1)}"/>""" : "")}</w:pPr></w:style>""";

    static string Heading(int level, int halfPoints) =>
        $"""<w:style w:type="paragraph" w:styleId="Heading{level}"><w:name w:val="heading {level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="{(level == 1 ? 360 : 240)}" w:after="120"/><w:outlineLvl w:val="{level - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="{halfPoints}"/><w:szCs w:val="{halfPoints}"/></w:rPr></w:style>""";
}
