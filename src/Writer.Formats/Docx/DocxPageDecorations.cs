using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

static partial class DocxSection
{
    const string VmlNamespace = "urn:schemas-microsoft-com:vml";
    static bool WatermarkShape(OpenXmlElement e) => e.NamespaceUri == VmlNamespace && e.LocalName == "shape"
        && e.GetAttributes().Any(a => a.LocalName == "id" && ((a.Value ?? "").StartsWith("WriterWatermark", StringComparison.Ordinal) || (a.Value ?? "").StartsWith("PowerPlusWaterMarkObject", StringComparison.Ordinal)));
    static bool WatermarkParagraph(W.Paragraph p) => p.Descendants().Any(WatermarkShape);

    static void ReadDecorations(DocxDocument doc, Dictionary<string, string> props)
    {
        if (doc.Main.Document!.GetFirstChild<W.DocumentBackground>()?.Color?.Value is { } color) props["pageColor"] = color;
        var shape = doc.Main.HeaderParts.SelectMany(p => p.Header?.Descendants() ?? []).FirstOrDefault(WatermarkShape);
        var text = shape?.Descendants().FirstOrDefault(e => e.NamespaceUri == VmlNamespace && e.LocalName == "textpath");
        if (text?.GetAttributes().FirstOrDefault(a => a.LocalName == "string").Value is { Length: > 0 } watermark) props["watermark"] = watermark;
    }

    public static void SetPageColor(DocxDocument doc, string value)
    {
        var document = doc.Main.Document!;
        document.RemoveAllChildren<W.DocumentBackground>();
        if (value != "none") document.AddChild(new W.DocumentBackground { Color = value });
        var settings = (doc.Main.DocumentSettingsPart ?? doc.Main.AddNewPart<DocumentSettingsPart>()).Settings ??= new W.Settings();
        settings.RemoveAllChildren<W.DisplayBackgroundShape>();
        if (value != "none") settings.AddChild(new W.DisplayBackgroundShape());
    }

    public static void SetWatermark(DocxDocument doc, string value)
    {
        foreach (var header in doc.Main.HeaderParts)
            foreach (var p in (header.Header?.Elements<W.Paragraph>() ?? []).Where(WatermarkParagraph).ToList())
            {
                foreach (var shape in p.Descendants().Where(WatermarkShape).ToList()) shape.Remove();
                foreach (var pict in p.Descendants<W.Picture>().Where(x => !x.HasChildren).ToList()) pict.Remove();
                foreach (var run in p.Descendants<W.Run>().Where(x => x.ChildElements.All(c => c is W.RunProperties)).ToList()) run.Remove();
                if (p.ChildElements.All(c => c is W.ParagraphProperties)) p.Remove();
            }
        if (value.Length == 0) return;
        var seen = new HashSet<HeaderPart>();
        foreach (var section in Sections(doc))
        {
            var slots = new List<bool> { false };
            if (DocxRun.On(section.GetFirstChild<W.TitlePage>())) slots.Add(true);
            foreach (var first in slots)
            {
                var part = Part(doc, section, true, first) as HeaderPart;
                if (part is null)
                {
                    part = doc.Main.AddNewPart<HeaderPart>();
                    part.Header = new W.Header();
                    AddReference(section, new W.HeaderReference { Type = first ? W.HeaderFooterValues.First : W.HeaderFooterValues.Default, Id = doc.Main.GetIdOfPart(part) });
                }
                if (!seen.Add(part)) continue;
                var escaped = System.Security.SecurityElement.Escape(value);
                var paragraph = new W.Paragraph($"""
                    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
                      <w:r><w:pict><v:shape id="WriterWatermark" type="#_x0000_t136" style="position:absolute;width:400pt;height:100pt;rotation:315;z-index:-251654144;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="#C0C0C0" stroked="f">
                        <v:path textpathok="t"/><v:textpath on="t" fitshape="t" style="font-family:Calibri;font-size:1pt" string="{escaped}"/>
                      </v:shape></w:pict></w:r>
                    </w:p>
                    """);
                (part.Header ??= new W.Header()).Append(paragraph);
            }
        }
    }
}
