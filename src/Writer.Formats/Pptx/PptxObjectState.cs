using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>Selection-pane metadata uses Office's own nonvisual drawing properties and locks.</summary>
static class PptxObjectState
{
    static OpenXmlElement? NonVisual(OpenXmlElement element) => element.ChildElements.FirstOrDefault(e => e.LocalName.StartsWith("nv", StringComparison.Ordinal));
    static P.NonVisualDrawingProperties? Drawing(OpenXmlElement element) => NonVisual(element)?.GetFirstChild<P.NonVisualDrawingProperties>();
    static OpenXmlElement? LockParent(OpenXmlElement element) => NonVisual(element)?.ChildElements.FirstOrDefault(e => e.LocalName.StartsWith("cNv", StringComparison.Ordinal) && e.LocalName != "cNvPr");
    static bool Enabled(OpenXmlElement? element, string name) => element?.GetAttributes().Any(a => a.LocalName == name && a.Value is "1" or "true") == true;
    public static void Read(OpenXmlElement element, Dictionary<string,string> props)
    {
        if (Drawing(element) is { } nv) { if (nv.Name?.Value is { } name) props["name"] = name; if (nv.Hidden?.Value == true) props["hidden"] = "true"; }
        var locks = LockParent(element)?.ChildElements.FirstOrDefault(e => e.LocalName.EndsWith("Locks", StringComparison.Ordinal));
        if (Enabled(locks,"noMove") || Enabled(locks,"noResize") || Enabled(locks,"noSelect")) props["locked"] = "true";
    }
    public static bool Set(OpenXmlElement element, string name, string value)
    {
        if (name is not ("name" or "hidden" or "locked")) return false;
        var nv = Drawing(element);
        if (nv is null) return false;
        if (name == "name") nv.Name = value;
        else if (name == "hidden") nv.Hidden = value == "true" ? true : null;
        else if (LockParent(element) is OpenXmlCompositeElement parent)
        {
            var locks = parent.ChildElements.FirstOrDefault(e => e.LocalName.EndsWith("Locks", StringComparison.Ordinal));
            if (locks is null && value == "true")
            {
                locks = element switch { P.Shape => new A.ShapeLocks(), P.Picture => new A.PictureLocks(), P.GroupShape => new A.GroupShapeLocks(), P.ConnectionShape => new A.ConnectionShapeLocks(), _ => new A.GraphicFrameLocks() };
                parent.PrependChild(locks);
            }
            if (locks is not null) foreach (var key in new[] { "noMove", "noResize", "noSelect" })
            {
                if (value == "true") locks.SetAttribute(new OpenXmlAttribute(key, "", "1"));
                else locks.RemoveAttribute(key, "");
            }
        }
        return true;
    }
}
