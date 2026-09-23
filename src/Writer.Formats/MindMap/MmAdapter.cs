using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Writer.Core;
using Writer.Formats.Common;

namespace Writer.Formats.MindMap;

/// <summary>FreeMind .mm files: a &lt;map&gt; of nested &lt;node&gt; elements. Elements and attributes writer does not model
/// (fonts, edges, hooks, arrow links, timestamps) ride along untouched, as does the file's whitespace.</summary>
public sealed class MmAdapter : IFormatAdapter
{
    public string Format => "mm";
    public IReadOnlyList<string> Extensions { get; } = [".mm"];
    public bool CanWrite => true;

    public Document Create() =>
        Fresh();

    static MmDocument Fresh()
    {
        var doc = new MmDocument(XDocument.Parse("<map version=\"1.0.1\">\n<node TEXT=\"Central Topic\"/>\n</map>\n", LoadOptions.PreserveWhitespace));
        doc.EnsureIds();
        return doc;
    }

    public Document Open(Stream stream)
    {
        XDocument xml;
        try
        {
            xml = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException ex)
        {
            throw new WriterException(ErrorCode.FormatError, $"Not a FreeMind file: {ex.Message}", "A .mm file is XML with a <map> root element.");
        }
        if (xml.Root?.Name.LocalName != "map")
            throw new WriterException(ErrorCode.FormatError, "Not a FreeMind file: the root element is not <map>", "A .mm file is XML with a <map> root element.");
        return new MmDocument(xml);
    }
}

public sealed class MmDocument(XDocument xml) : Document
{
    internal XElement Map => xml.Root!;

    public override string Format => "mm";
    public override Node Root => new MmRoot(this);

    public string Render()
    {
        var ms = new MemoryStream();
        Save(ms);
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    /// <summary>UTF-8 without BOM, no indentation added, declaration only when the source had one.</summary>
    public override void Save(Stream stream)
    {
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false), OmitXmlDeclaration = xml.Declaration is null, Indent = false, NewLineHandling = NewLineHandling.Entitize,
        };
        var ms = new MemoryStream();
        using (var writer = XmlWriter.Create(ms, settings)) xml.Save(writer);
        // ponytail: XmlWriter closes empty elements as "<x />" while FreeMind writes "<x/>"; a text replace keeps untouched leaves diff-free.
        stream.Write(Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(ms.ToArray()).Replace(" />", "/>")));
    }

    /// <summary>Gives every node an ID, so the UI and agents can address topics by [@id]. Runs before any change;
    /// reading never adds them, so untouched files stay byte for byte the same.</summary>
    internal void EnsureIds()
    {
        foreach (var e in xml.Descendants().Where(e => e.Name.LocalName == "node" && string.IsNullOrEmpty((string?)e.Attribute("ID"))).ToList())
            e.SetAttributeValue("ID", NewId());
    }

    internal string NewId()
    {
        var used = xml.Descendants().Select(e => (string?)e.Attribute("ID")).ToHashSet();
        string id;
        do id = "ID_" + Random.Shared.Next(100_000_000, int.MaxValue).ToString(CultureInfo.InvariantCulture);
        while (used.Contains(id));
        return id;
    }
}

sealed class MmRoot(MmDocument doc) : Node
{
    public override string Kind => "document";
    public override string Format => "mm";
    public override object Anchor => doc;

    protected override IEnumerable<Node> ProjectChildren() => MmTopic.Nodes(doc.Map).Select(e => (Node)new MmTopic(doc, e));

    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string> { ["format"] = "mm" };

    public override string GetRaw() => doc.Render();

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        if (MmTopic.Nodes(doc.Map).Any())
            throw new WriterException(ErrorCode.Validation, "A map has one root topic; add under /topic[1]", "Example: add map.mm /topic[1] --type topic --prop text=\"Idea\"");
        return MmTopic.Create(doc, doc.Map, kind, props, null);
    }
}

sealed class MmTopic(MmDocument doc, XElement node) : Node
{
    public override string Kind => "topic";
    public override object Anchor => node;

    protected override IEnumerable<Node> ProjectChildren() => Nodes(node).Select(e => (Node)new MmTopic(doc, e));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = Rich("NODE").FirstOrDefault() is { } rich ? RichText(rich) : (string?)node.Attribute("TEXT") ?? "" };
        if (Rich("NOTE").FirstOrDefault() is { } note) props["note"] = RichText(note);
        if ((string?)node.Attribute("FOLDED") == "true") props["collapsed"] = "true";
        var side = (string?)node.Attribute("POSITION");
        if (side is "left" or "right") props["side"] = side;
        if ((string?)node.Attribute("LINK") is { Length: > 0 } link) props["link"] = link;
        if (Color("COLOR") is { } color) props["color"] = color;
        if (Color("BACKGROUND_COLOR") is { } fill) props["fill"] = fill;
        if ((string?)Icons().FirstOrDefault()?.Attribute("BUILTIN") is { Length: > 0 } icon) props["icon"] = icon;
        if ((string?)node.Attribute("ID") is { Length: > 0 } id) props["id"] = id;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        doc.EnsureIds();
        switch (name)
        {
            case "text": SetText(value); break;
            case "md": SetText(string.Concat(InlineMarkdown.Parse(value).Select(r => r.Text))); break;
            case "html": SetText(string.Concat(InlineHtml.Parse(value).Select(r => r.Text))); break;
            case "note": SetNote(value); break;
            case "collapsed": node.SetAttributeValue("FOLDED", value == "true" ? "true" : null); break;
            case "side": node.SetAttributeValue("POSITION", value); break;
            case "link": node.SetAttributeValue("LINK", value.Length == 0 ? null : value); break;
            case "color": node.SetAttributeValue("COLOR", value == "none" ? null : "#" + value.ToLowerInvariant()); break;
            case "fill": node.SetAttributeValue("BACKGROUND_COLOR", value == "none" ? null : "#" + value.ToLowerInvariant()); break;
            case "icon": SetIcon(value); break;
        }
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        doc.EnsureIds();
        return Create(doc, node, kind, props, index);
    }

    public override void Remove()
    {
        doc.EnsureIds();
        if (IsRoot) throw new WriterException(ErrorCode.Validation, "The root topic cannot be removed", "Remove its children, or delete the file.");
        Detach();
    }

    public override void MoveTo(Node newParent, int? index)
    {
        doc.EnsureIds();
        if (newParent is not MmTopic target)
            throw new WriterException(ErrorCode.Validation, "A map has one root topic; move under a topic instead", "Example: --to /topic[1] --index 1");
        if (IsRoot) throw new WriterException(ErrorCode.Validation, "The root topic cannot be moved", "Move its children instead.");
        var to = (XElement)target.Anchor;
        if (to.AncestorsAndSelf().Contains(node))
            throw new WriterException(ErrorCode.Validation, "A topic cannot move under itself", "Pick a topic outside this branch.");
        Detach();
        Place(to, node, index);
    }

    public override string GetRaw() => node.ToString(SaveOptions.DisableFormatting);

    public override void SetRaw(string raw)
    {
        XElement parsed;
        try
        {
            parsed = XElement.Parse(raw, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException ex)
        {
            throw new WriterException(ErrorCode.Validation, $"Raw XML could not be parsed: {ex.Message}", "Start from 'get --raw' output and edit that.");
        }
        if (parsed.Name.LocalName != "node")
            throw new WriterException(ErrorCode.Validation, "Raw XML must be a <node> element", "Start from 'get --raw' output and edit that.");
        node.ReplaceWith(parsed);
    }

    internal static MmTopic Create(MmDocument doc, XElement parent, string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        if (kind != "topic") throw new WriterException(ErrorCode.UnsupportedKind, $"Cannot add {kind} to a mind map", "Mind maps hold topics only.");
        var elem = new XElement(parent.Name.Namespace + "node", new XAttribute("TEXT", ""), new XAttribute("ID", doc.NewId()));
        Place(parent, elem, index);
        var topic = new MmTopic(doc, elem);
        foreach (var (name, value) in props) topic.SetProp(name, value);
        return topic;
    }

    internal static IEnumerable<XElement> Nodes(XElement parent) => parent.Elements().Where(e => e.Name.LocalName == "node");

    bool IsRoot => node.Parent is null || node.Parent.Name.LocalName == "map";

    IEnumerable<XElement> Rich(string type) => node.Elements().Where(e => e.Name.LocalName == "richcontent" && (string?)e.Attribute("TYPE") == type);

    IEnumerable<XElement> Icons() => node.Elements().Where(e => e.Name.LocalName == "icon");

    string? Color(string attribute) => (string?)node.Attribute(attribute) is { Length: > 0 } v ? v.TrimStart('#').ToUpperInvariant() : null;

    void SetText(string value)
    {
        node.SetAttributeValue("TEXT", value);
        Rich("NODE").Remove();
    }

    void SetNote(string value)
    {
        var notes = Rich("NOTE").ToList();
        if (value.Length == 0)
        {
            notes.Remove();
            return;
        }
        var html = new XElement("html", new XElement("head"), new XElement("body", value.ReplaceLineEndings("\n").Split('\n').Select(line => new XElement("p", line))));
        if (notes.Count > 0)
        {
            notes[0].ReplaceNodes(html);
            return;
        }
        var rich = new XElement(node.Name.Namespace + "richcontent", new XAttribute("TYPE", "NOTE"), html);
        if (Nodes(node).FirstOrDefault() is { } first) first.AddBeforeSelf(rich, Indent(first));
        else node.Add(rich);
    }

    void SetIcon(string value)
    {
        var icons = Icons().ToList();
        if (value.Length == 0 || value == "none")
        {
            icons.Remove();
            return;
        }
        if (icons.Count > 0)
        {
            icons[0].SetAttributeValue("BUILTIN", value);
            return;
        }
        var icon = new XElement(node.Name.Namespace + "icon", new XAttribute("BUILTIN", value));
        if (Nodes(node).FirstOrDefault() is { } first) first.AddBeforeSelf(icon, Indent(first));
        else node.Add(icon);
    }

    /// <summary>The text of a richcontent block: one line per paragraph, whitespace collapsed.</summary>
    static string RichText(XElement rich)
    {
        var leaves = rich.Descendants().Where(e => IsBlock(e) && !e.Descendants().Any(IsBlock)).ToList();
        var parts = leaves.Count > 0 ? leaves.Select(e => Collapse(e.Value)) : [Collapse(rich.Value)];
        return string.Join("\n", parts.Where(p => p.Length > 0));
    }

    static bool IsBlock(XElement e) => e.Name.LocalName is "p" or "li" or "div" or "pre" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6";

    static string Collapse(string s) => Regex.Replace(s, @"\s+", " ").Trim();

    /// <summary>Inserts a node at a 1-based position among the parent's child nodes (null appends), copying a neighbour's indentation.</summary>
    static void Place(XElement parent, XElement elem, int? index)
    {
        var siblings = Nodes(parent).ToList();
        var at = index is { } i ? Math.Clamp(i - 1, 0, siblings.Count) : siblings.Count;
        if (at < siblings.Count) siblings[at].AddBeforeSelf(elem, Indent(siblings[at]));
        else if (siblings.Count > 0) siblings[^1].AddAfterSelf(Indent(siblings[^1]), elem);
        else parent.Add(elem);
    }

    void Detach()
    {
        Whitespace(node)?.Remove();
        node.Remove();
    }

    static XText? Whitespace(XNode n) => n.PreviousNode is XText t && string.IsNullOrWhiteSpace(t.Value) ? t : null;

    static XText? Indent(XElement e) => Whitespace(e) is { } t ? new XText(t) : null;
}
