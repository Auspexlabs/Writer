using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
namespace Writer.Formats.Pptx;

static class PptxFontReplace
{
    internal static void Apply(PptxDocument doc, string json)
    {
        using var data = JsonDocument.Parse(json);
        var root = data.RootElement;
        var from = root.TryGetProperty("from", out var f) && f.ValueKind == JsonValueKind.String ? f.GetString()!.Trim() : "";
        var to = root.TryGetProperty("to", out var t) && t.ValueKind == JsonValueKind.String ? t.GetString()!.Trim() : "";
        if (from.Length == 0 || to.Length == 0 || to.Length > 255 || to.Any(char.IsControl))
            throw new WriterException(ErrorCode.Validation, "Both font names are required", "Use {from,to} with non-empty font names.");
        var visited = new HashSet<OpenXmlPart>();
        void Visit(OpenXmlPart part)
        {
            if (!visited.Add(part)) return;
            if (part.RootElement is { } xml)
                foreach (var element in xml.Descendants().Where(e => e.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/main" && e.LocalName is "latin" or "ea" or "cs" or "font" or "buFont"))
                {
                    var face = element.GetAttribute("typeface", "");
                    if (string.Equals(face.Value, from, StringComparison.OrdinalIgnoreCase))
                        element.SetAttribute(new OpenXmlAttribute("typeface", "", to));
                }
            foreach (var child in part.Parts) Visit(child.OpenXmlPart);
        }
        Visit(doc.Presentation);
    }
}
