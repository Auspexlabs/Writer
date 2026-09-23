using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Writer.Core;

namespace Writer.Formats.Common;

/// <summary>Makes hand-written XML fragments parseable and swaps them into an Open XML tree.</summary>
public static class RawXml
{
    static readonly Dictionary<string, string> Known = new()
    {
        ["w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ["r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ["wp"] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        ["a"] = "http://schemas.openxmlformats.org/drawingml/2006/main",
        ["pic"] = "http://schemas.openxmlformats.org/drawingml/2006/picture",
        ["c"] = "http://schemas.openxmlformats.org/drawingml/2006/chart",
        ["m"] = "http://schemas.openxmlformats.org/officeDocument/2006/math",
        ["mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ["w14"] = "http://schemas.microsoft.com/office/word/2010/wordml",
        ["w15"] = "http://schemas.microsoft.com/office/word/2012/wordml",
        ["wps"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        ["wpg"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        ["a14"] = "http://schemas.microsoft.com/office/drawing/2010/main",
        ["a16"] = "http://schemas.microsoft.com/office/drawing/2014/main",
        ["v"] = "urn:schemas-microsoft-com:vml",
        ["o"] = "urn:schemas-microsoft-com:office:office",
        ["w10"] = "urn:schemas-microsoft-com:office:word",
        ["p"] = "http://schemas.openxmlformats.org/presentationml/2006/main",
        ["p14"] = "http://schemas.microsoft.com/office/powerpoint/2010/main",
        ["x"] = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    };

    static readonly Regex Prefixes = new(@"[<\s/]([A-Za-z][\w.-]*):[A-Za-z]");

    /// <summary>The fragment with xmlns declarations added to its root start tag for every known prefix it uses but does not declare.</summary>
    public static string WithNamespaces(string xml, IEnumerable<KeyValuePair<string, string>> declaredByDocument)
    {
        var text = xml.Trim();
        var end = text.IndexOf('>');
        if (!text.StartsWith('<') || end < 0)
            throw new WriterException(ErrorCode.Validation, "Raw value must be one XML element", "Example: <w:p><w:r><w:t>text</w:t></w:r></w:p>");
        var startTag = text[..end];
        var known = new Dictionary<string, string>(Known);
        foreach (var (prefix, uri) in declaredByDocument) known[prefix] = uri;
        var missing = new List<string>();
        foreach (Match m in Prefixes.Matches(text))
        {
            var prefix = m.Groups[1].Value;
            if (prefix is "xmlns" or "xml" || missing.Contains(prefix) || !known.ContainsKey(prefix)) continue;
            if (startTag.Contains($"xmlns:{prefix}=", StringComparison.Ordinal)) continue;
            missing.Add(prefix);
        }
        if (missing.Count == 0) return text;
        var nameEnd = startTag.IndexOfAny([' ', '\t', '\r', '\n', '/']);
        if (nameEnd < 0) nameEnd = startTag.Length;
        return text[..nameEnd] + string.Concat(missing.Select(p => $" xmlns:{p}=\"{known[p]}\"")) + text[nameEnd..];
    }

    /// <summary>Replaces an element in its parent with the element parsed from raw XML. The new element must be of the same type.</summary>
    public static void Replace(OpenXmlElement element, string xml, IEnumerable<KeyValuePair<string, string>> declaredByDocument)
    {
        var parent = element.Parent
            ?? throw new WriterException(ErrorCode.Validation, "This element cannot be replaced", "Replace one of its children instead.");
        var replacement = Parse(parent, xml, declaredByDocument);
        if (replacement.GetType() != element.GetType())
            throw new WriterException(ErrorCode.Validation,
                $"Raw XML must be a <{element.Prefix}:{element.LocalName}> element, got <{replacement.Prefix}:{replacement.LocalName}>",
                "Start from 'get --raw' output and edit that.");
        parent.ReplaceChild(replacement, element);
    }

    /// <summary>One element parsed from raw XML as a child of container, not attached anywhere yet.</summary>
    public static OpenXmlElement Parse(OpenXmlElement container, string xml, IEnumerable<KeyValuePair<string, string>> declaredByDocument)
    {
        var wrapper = container.CloneNode(false);
        try
        {
            wrapper.InnerXml = WithNamespaces(xml, declaredByDocument);
        }
        catch (Exception ex) when (ex is not WriterException)
        {
            throw new WriterException(ErrorCode.Validation, $"Raw XML could not be parsed: {ex.Message}", "Check that the XML is well-formed and uses the right prefixes.");
        }
        var element = wrapper.FirstChild
            ?? throw new WriterException(ErrorCode.Validation, "Raw XML is empty", "Pass one element.");
        element.Remove();
        return element;
    }
}
