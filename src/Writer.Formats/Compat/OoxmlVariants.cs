using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Docx;
using Writer.Formats.Pptx;
using Writer.Formats.Xlsx;

namespace Writer.Formats.Compat;

/// <summary>Macro-enabled, template and slideshow OOXML (.docm .dotx .dotm .xlsm .xltx .xltm .pptm .potx .potm .ppsx
/// .ppsm): the same package as docx/xlsx/pptx under another content type. Opened as the plain document type with the
/// VBA project left out, so 另存为 gives a clean .docx/.xlsx/.pptx. The main part decides, not the extension.</summary>
static class OoxmlVariants
{
    public static Document Open(byte[] bytes, ZipArchive zip)
    {
        var ms = new MemoryStream(bytes);
        try
        {
            if (zip.Entries.Any(e => e.FullName.StartsWith("word/", StringComparison.Ordinal)))
            {
                var package = WordprocessingDocument.Open(ms, true);
                if (package.MainDocumentPart?.VbaProjectPart is { } vba) package.MainDocumentPart.DeletePart(vba);
                if (package.DocumentType != WordprocessingDocumentType.Document) package.ChangeDocumentType(WordprocessingDocumentType.Document);
                return new DocxDocument(package);
            }
            if (zip.Entries.Any(e => e.FullName.StartsWith("xl/", StringComparison.Ordinal)))
            {
                var package = SpreadsheetDocument.Open(ms, true);
                if (package.WorkbookPart?.VbaProjectPart is { } vba) package.WorkbookPart.DeletePart(vba);
                if (package.DocumentType != SpreadsheetDocumentType.Workbook) package.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                return new XlsxDocument(package);
            }
            if (zip.Entries.Any(e => e.FullName.StartsWith("ppt/", StringComparison.Ordinal)))
            {
                var package = PresentationDocument.Open(ms, true);
                if (package.PresentationPart?.VbaProjectPart is { } vba) package.PresentationPart.DeletePart(vba);
                if (package.DocumentType != PresentationDocumentType.Presentation) package.ChangeDocumentType(PresentationDocumentType.Presentation);
                return new PptxDocument(package);
            }
        }
        catch (Exception ex) when (ex is not WriterException)
        {
            throw new WriterException(ErrorCode.FormatError, $"Not a valid Office file: {ex.Message}", "Check that the file opens in Office.");
        }
        throw new WriterException(ErrorCode.FormatError, "Not a Word, Excel or PowerPoint package", $"The zip has no word/, xl/ or ppt/ main part; it holds {string.Join(", ", zip.Entries.Select(e => e.FullName).Take(12))}.");
    }
}
