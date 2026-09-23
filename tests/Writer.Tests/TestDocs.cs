using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Docx;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Tests;

/// <summary>Builders for small test documents.</summary>
public static class TestDocs
{
    public static string FixtureDir(string format) => Path.Combine(AppContext.BaseDirectory, "Fixtures", format);

    public static IEnumerable<string> Fixtures(string format, string pattern) =>
        Directory.GetFiles(FixtureDir(format), pattern, SearchOption.AllDirectories).OrderBy(f => f, StringComparer.Ordinal);

    /// <summary>A blank template document with the given blocks in the body; extra may add numbering, styles or relationships.</summary>
    public static byte[] Docx(IEnumerable<OpenXmlElement> blocks, Action<MainDocumentPart>? extra = null)
    {
        var ms = new MemoryStream();
        DocxTemplate.WriteBlank(ms);
        ms.Position = 0;
        using (var package = WordprocessingDocument.Open(ms, true))
        {
            var main = package.MainDocumentPart!;
            var body = main.Document!.Body!;
            var section = body.GetFirstChild<W.SectionProperties>();
            foreach (var block in blocks) body.InsertBefore(block, section);
            extra?.Invoke(main);
        }
        return ms.ToArray();
    }

    public static byte[] Docx(params OpenXmlElement[] blocks) => Docx(blocks, null);

    public static Document OpenDocx(byte[] bytes) => new DocxAdapter().Open(new MemoryStream(bytes));

    public static W.Paragraph P(string text, string? style = null)
    {
        var p = new W.Paragraph();
        if (style is not null) p.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = style });
        p.Append(new W.Run(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return p;
    }

    public static W.TableCell Cell(string text) => new(text.Split('\n').Select(line => (OpenXmlElement)P(line)));

    /// <summary>Bytes with a valid PNG header for the given pixel size. Enough for header parsing; not a viewable image.</summary>
    public static byte[] FakePng(int width, int height)
    {
        var bytes = new byte[33];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52 }.CopyTo(bytes, 0);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(16), width);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(20), height);
        bytes[24] = 8;
        bytes[25] = 6;
        return bytes;
    }

    /// <summary>A two-page PDF with a title line, two body lines and a second page, built with PdfPig.</summary>
    public static byte[] Pdf()
    {
        var builder = new UglyToad.PdfPig.Writer.PdfDocumentBuilder();
        var font = builder.AddStandard14Font(UglyToad.PdfPig.Fonts.Standard14Fonts.Standard14Font.Helvetica);
        var page = builder.AddPage(UglyToad.PdfPig.Content.PageSize.A4);
        page.AddText("Quarterly Report", 18, new UglyToad.PdfPig.Core.PdfPoint(50, 780), font);
        page.AddText("Revenue grew 25% year over year.", 11, new UglyToad.PdfPig.Core.PdfPoint(50, 740), font);
        page.AddText("Costs were flat.", 11, new UglyToad.PdfPig.Core.PdfPoint(50, 725), font);
        var second = builder.AddPage(UglyToad.PdfPig.Content.PageSize.A4);
        second.AddText("Second page", 11, new UglyToad.PdfPig.Core.PdfPoint(50, 780), font);
        return builder.Build();
    }

    public static W.Paragraph ListP(string text, int numId, int level)
    {
        var p = P(text);
        p.ParagraphProperties = new W.ParagraphProperties(new W.NumberingProperties(
            new W.NumberingLevelReference { Val = level }, new W.NumberingId { Val = numId }));
        return p;
    }

    /// <summary>numId 1 = bullets, numId 2 = decimal then lower letters.</summary>
    public static void AddNumbering(MainDocumentPart main)
    {
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = new W.Numbering(
            new W.AbstractNum(
                new W.Level(new W.NumberingFormat { Val = W.NumberFormatValues.Bullet }, new W.LevelText { Val = "•" }) { LevelIndex = 0 },
                new W.Level(new W.NumberingFormat { Val = W.NumberFormatValues.Bullet }, new W.LevelText { Val = "◦" }) { LevelIndex = 1 })
            { AbstractNumberId = 1 },
            new W.AbstractNum(
                new W.Level(new W.NumberingFormat { Val = W.NumberFormatValues.Decimal }, new W.LevelText { Val = "%1." }) { LevelIndex = 0 },
                new W.Level(new W.NumberingFormat { Val = W.NumberFormatValues.LowerLetter }, new W.LevelText { Val = "%2." }) { LevelIndex = 1 })
            { AbstractNumberId = 2 },
            new W.NumberingInstance(new W.AbstractNumId { Val = 1 }) { NumberID = 1 },
            new W.NumberingInstance(new W.AbstractNumId { Val = 2 }) { NumberID = 2 });
    }
}
