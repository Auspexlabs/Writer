using System.Globalization;
using System.Text;
using System.Text.Json;
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
    /// <summary>Opened from an .xmind file: exporting it to .mm writes this very map.</summary>
    public bool FromXmind { get; init; }

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
        var icons = Icons().Select(i => (string?)i.Attribute("BUILTIN")).Where(i => !string.IsNullOrEmpty(i)).ToList();
        if (icons.Count > 0) props["icon"] = string.Join(",", icons);
        var labels = Attrs("label").Select(a => (string?)a.Attribute("VALUE")).Where(v => !string.IsNullOrEmpty(v)).ToList();
        if (labels.Count > 0) props["labels"] = string.Join(",", labels);
        if (Picture() is { } picture)
        {
            if ((string?)picture.Attribute("URI") is { Length: > 0 } uri) props["image"] = uri;
            if ((string?)picture.Attribute("WIDTH") is { Length: > 0 } w && (string?)picture.Attribute("HEIGHT") is { Length: > 0 } h) props["imageSize"] = w + "," + h;
        }
        if (Font() is { } font)
        {
            foreach (var (prop, attribute) in FontFlags)
                if ((string?)font.Attribute(attribute) == "true") props[prop] = "true";
            if ((string?)font.Attribute("SIZE") is { Length: > 0 } size) props["size"] = size;
            if ((string?)font.Attribute("NAME") is { Length: > 0 } name) props["font"] = name;
        }
        if (Attr("free") is { } free) props["free"] = free;
        if (Attr("summary") is { } summary) props["summary"] = summary;
        foreach (var name in MapAttrs) if (Attr(name) is { } v) props[name] = v;
        if (Attr("mono") == "true") props["mono"] = "true";
        if (Cloud() is { } cloud) props["cloud"] = (string?)cloud.Attribute("COLOR") is { Length: > 0 } c ? c.TrimStart('#').ToUpperInvariant() : "F0F0F0";
        var arrows = Arrows().Select(a => new Rel((string?)a.Attribute("DESTINATION") ?? "", (string?)a.Attribute("MIDDLE_LABEL") ?? "",
            ((string?)a.Attribute("COLOR") ?? "").TrimStart('#').ToUpperInvariant(), ArrowsOf((string?)a.Attribute("STARTARROW"), (string?)a.Attribute("ENDARROW")))).Where(r => r.To.Length > 0).ToList();
        if (arrows.Count > 0)
            props["rels"] = NodeJson.Compact(w =>
            {
                w.WriteStartArray();
                foreach (var r in arrows)
                {
                    w.WriteStartObject();
                    w.WriteString("to", r.To); w.WriteString("label", r.Label); w.WriteString("color", r.Color); w.WriteString("arrows", r.Arrows);
                    w.WriteEndObject();
                }
                w.WriteEndArray();
            });
        if ((string?)node.Attribute("ID") is { Length: > 0 } id) props["id"] = id;
        return props;
    }

    /// <summary>One relationship line as the rels prop carries it: a JSON array of these.</summary>
    sealed record Rel(string To, string Label, string Color, string Arrows);

    static string ArrowsOf(string? start, string? end)
    {
        var s = start is not null && start != "None";
        var e = end is null || end != "None";
        return s && e ? "both" : s ? "start" : e ? "end" : "none";
    }

    static List<Rel> ParseRels(string json)
    {
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array) throw new JsonException("expected an array");
        string Str(JsonElement e, string name) => e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";
        return doc.RootElement.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.Object).Select(e => new Rel(Str(e, "to"), Str(e, "label"), Str(e, "color"), Str(e, "arrows"))).ToList();
    }

    /// <summary>How the map (or one branch) is drawn: attribute rows FreeMind carries along.</summary>
    static readonly string[] MapAttrs = ["structure", "theme", "lines"];

    /// <summary>Topic props kept as flags on FreeMind's &lt;font&gt; element (STRIKETHROUGH is Freeplane's).</summary>
    static readonly (string Prop, string Attribute)[] FontFlags = [("bold", "BOLD"), ("italic", "ITALIC"), ("strike", "STRIKETHROUGH")];

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
            case "bold" or "italic" or "strike": SetFont(FontFlags.First(f => f.Prop == name).Attribute, value == "true" ? "true" : null); break;
            case "size": SetFont("SIZE", value is "" or "0" ? null : value); break;
            case "font": SetFont("NAME", value.Length == 0 ? null : value); break;
            case "free": SetAttr("free", value); break;
            case "labels": SetLabels(value); break;
            case "summary": SetAttr("summary", value); break;
            case "structure" or "theme" or "lines": SetAttr(name, value == "none" ? "" : value); break;
            case "mono": SetAttr("mono", value == "true" ? "true" : ""); break;
            case "cloud": SetCloud(value); break;
            case "rels": SetRels(value); break;
            case "image": SetPicture("URI", value.Length == 0 ? null : value); break;
            case "imageSize":
                var size = value.Split(',');
                SetPicture("WIDTH", size.Length == 2 ? size[0].Trim() : null);
                SetPicture("HEIGHT", size.Length == 2 ? size[1].Trim() : null);
                break;
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

    XElement? Font() => node.Elements().FirstOrDefault(e => e.Name.LocalName == "font");

    // What FreeMind has no element for rides in its <attribute NAME=".." VALUE=".."/> rows, which every FreeMind keeps and shows as a small table
    IEnumerable<XElement> Attrs(string name) => node.Elements().Where(e => e.Name.LocalName == "attribute" && (string?)e.Attribute("NAME") == name);

    string? Attr(string name) => (string?)Attrs(name).FirstOrDefault()?.Attribute("VALUE") is { Length: > 0 } v ? v : null;

    /// <summary>Sets the one attribute row called name; an empty value removes it.</summary>
    void SetAttr(string name, string value)
    {
        var rows = Attrs(name).ToList();
        if (value.Length == 0)
        {
            foreach (var row in rows) { Whitespace(row)?.Remove(); row.Remove(); }
            return;
        }
        if (rows.Count > 0) rows[0].SetAttributeValue("VALUE", value);
        else Insert(new XElement(node.Name.Namespace + "attribute", new XAttribute("NAME", name), new XAttribute("VALUE", value)));
    }

    /// <summary>One attribute of the topic's &lt;font&gt;: null clears it, and a font left without attributes goes away.</summary>
    void SetFont(string attribute, string? value)
    {
        var font = Font();
        if (value is null)
        {
            if (font is null) return;
            font.SetAttributeValue(attribute, null);
            if (!font.HasAttributes)
            {
                Whitespace(font)?.Remove();
                font.Remove();
            }
            return;
        }
        if (font is null)
        {
            font = new XElement(node.Name.Namespace + "font");
            Insert(font);
        }
        font.SetAttributeValue(attribute, value);
    }

    /// <summary>Adds a child element of the topic before its first child topic, as FreeMind writes them.</summary>
    void Insert(XElement elem)
    {
        if (Nodes(node).FirstOrDefault() is { } first) first.AddBeforeSelf(elem, Indent(first));
        else node.Add(elem);
    }

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
        Insert(new XElement(node.Name.Namespace + "richcontent", new XAttribute("TYPE", "NOTE"), html));
    }

    /// <summary>The icons and markers as a comma-separated list of builtin names, replacing what was there; none clears them.</summary>
    void SetIcon(string value)
    {
        foreach (var icon in Icons().ToList()) { Whitespace(icon)?.Remove(); icon.Remove(); }
        if (value == "none") return;
        foreach (var name in value.Split(',').Select(s => s.Trim()).Where(s => s.Length > 0))
            Insert(new XElement(node.Name.Namespace + "icon", new XAttribute("BUILTIN", name)));
    }

    /// <summary>XMind-style labels: one FreeMind attribute row called label per entry of the comma-separated list.</summary>
    void SetLabels(string value)
    {
        foreach (var row in Attrs("label").ToList()) { Whitespace(row)?.Remove(); row.Remove(); }
        foreach (var label in value.Split(',').Select(s => s.Trim()).Where(s => s.Length > 0))
            Insert(new XElement(node.Name.Namespace + "attribute", new XAttribute("NAME", "label"), new XAttribute("VALUE", label)));
    }

    // A boundary around a branch is FreeMind's cloud
    XElement? Cloud() => node.Elements().FirstOrDefault(e => e.Name.LocalName == "cloud");

    void SetCloud(string value)
    {
        var cloud = Cloud();
        if (value == "none")
        {
            if (cloud is null) return;
            Whitespace(cloud)?.Remove();
            cloud.Remove();
            return;
        }
        if (cloud is null) { cloud = new XElement(node.Name.Namespace + "cloud"); Insert(cloud); }
        cloud.SetAttributeValue("COLOR", "#" + value.ToLowerInvariant());
    }

    // Relationship lines are FreeMind arrowlinks on the source topic (MIDDLE_LABEL as Freeplane writes it)
    IEnumerable<XElement> Arrows() => node.Elements().Where(e => e.Name.LocalName == "arrowlink");

    void SetRels(string value)
    {
        List<Rel> rels;
        try { rels = ParseRels(value); }
        catch (JsonException ex) { throw new WriterException(ErrorCode.Validation, $"rels: {ex.Message}", "Example: rels=[{\"to\":\"ID_2\",\"label\":\"because\",\"color\":\"B5563A\",\"arrows\":\"end\"}]"); }
        foreach (var old in Arrows().ToList()) { Whitespace(old)?.Remove(); old.Remove(); }
        foreach (var rel in rels.Where(r => r.To.Length > 0))
        {
            var arrows = rel.Arrows.Length > 0 ? rel.Arrows : "end";
            var link = new XElement(node.Name.Namespace + "arrowlink", new XAttribute("DESTINATION", rel.To),
                new XAttribute("STARTARROW", arrows is "both" or "start" ? "Default" : "None"), new XAttribute("ENDARROW", arrows is "both" or "end" ? "Default" : "None"),
                new XAttribute("ID", "Arrow_" + doc.NewId()));
            if (!string.IsNullOrEmpty(rel.Color)) link.SetAttributeValue("COLOR", "#" + rel.Color.TrimStart('#').ToLowerInvariant());
            if (!string.IsNullOrEmpty(rel.Label)) link.SetAttributeValue("MIDDLE_LABEL", rel.Label);
            Insert(link);
        }
    }

    // A picture is Freeplane's external-object hook (URI plus our WIDTH / HEIGHT); FreeMind keeps hooks it does not know
    XElement? Picture() => node.Elements().FirstOrDefault(e => e.Name.LocalName == "hook" && (string?)e.Attribute("NAME") == "ExternalObject");

    void SetPicture(string attribute, string? value)
    {
        var hook = Picture();
        if (value is null)
        {
            if (hook is null) return;
            if (attribute == "URI") { Whitespace(hook)?.Remove(); hook.Remove(); }
            else hook.SetAttributeValue(attribute, null);
            return;
        }
        if (hook is null)
        {
            if (attribute != "URI") return;
            hook = new XElement(node.Name.Namespace + "hook", new XAttribute("NAME", "ExternalObject"));
            Insert(hook);
        }
        hook.SetAttributeValue(attribute, value);
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
