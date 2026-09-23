using DocumentFormat.OpenXml;
using Writer.Core;

namespace Writer.Formats.Common;

/// <summary>Positioning helpers shared by the Open XML adapters.</summary>
static class OoxmlTree
{
    /// <summary>Inserts before the top-level element of the child at the 1-based position among the parent's projected children.
    /// Returns false when the position means "append", so the caller can apply its own append rule.</summary>
    public static bool InsertBefore(Node parent, OpenXmlElement container, OpenXmlElement element, int? index)
    {
        if (index is not { } i) return false;
        var children = parent.Children;
        i = Math.Max(1, i);
        if (i > children.Count) return false;
        var top = (OpenXmlElement)children[i - 1].Anchor;
        while (top.Parent is not null && !ReferenceEquals(top.Parent, container)) top = top.Parent;
        if (!ReferenceEquals(top.Parent, container)) return false;
        container.InsertBefore(element, top);
        return true;
    }
}
