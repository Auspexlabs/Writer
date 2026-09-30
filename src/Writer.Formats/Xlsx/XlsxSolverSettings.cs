using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;

namespace Writer.Formats.Xlsx;

// Solver setup is per worksheet. Store it separately from result cells and from
// Excel's own add-in metadata, keyed by the sheet relationship (stable on rename).
static class XlsxSolverSettings
{
    static readonly XNamespace Ns = "urn:writer:spreadsheet-solver:v1";
    static CustomXmlPart? Find(XlsxDocument doc) => doc.Workbook.CustomXmlParts.FirstOrDefault(p =>
    {
        try { using var s = p.GetStream(); return XDocument.Load(s).Root?.Name == Ns + "solver"; }
        catch (System.Xml.XmlException) { return false; }
    });
    static XElement Load(CustomXmlPart? part)
    {
        if (part is null) return new XElement(Ns + "solver");
        using var stream = part.GetStream(); return XDocument.Load(stream).Root!;
    }
    public static string Read(XlsxSheet sheet)
    {
        var id = sheet.Doc.Workbook.GetIdOfPart(sheet.Part);
        return Load(Find(sheet.Doc)).Elements(Ns + "sheet").FirstOrDefault(e => (string?)e.Attribute("id") == id)?.Value ?? "null";
    }
    public static void Write(XlsxSheet sheet, string json)
    {
        var spec = JsonNode.Parse(json);
        if (spec is not null && (spec is not JsonObject || string.IsNullOrWhiteSpace(spec["target"]?.GetValue<string>()) || string.IsNullOrWhiteSpace(spec["variables"]?.GetValue<string>())))
            throw new WriterException(ErrorCode.Validation, "Solver setup needs target and variables", "Use {target,variables,mode,constraints,method,integer,nonnegative} or null.");
        var part = Find(sheet.Doc); var root = Load(part); var id = sheet.Doc.Workbook.GetIdOfPart(sheet.Part);
        var current = sheet.Doc.Workbook.Workbook!.Sheets!.ChildElements.Select(e => e.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value).ToHashSet();
        root.Elements(Ns + "sheet").Where(e => (string?)e.Attribute("id") == id || !current.Contains((string?)e.Attribute("id") ?? "")).Remove();
        if (spec is not null) root.Add(new XElement(Ns + "sheet", new XAttribute("id", id), spec.ToJsonString()));
        if (!root.HasElements) { if (part is not null) sheet.Doc.Workbook.DeletePart(part); return; }
        part ??= sheet.Doc.Workbook.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(root.ToString(SaveOptions.DisableFormatting))); part.FeedData(stream);
    }
}
