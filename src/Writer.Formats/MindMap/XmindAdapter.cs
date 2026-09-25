using System.Globalization;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Writer.Core;

namespace Writer.Formats.MindMap;

/// <summary>XMind files (.xmind: a zip holding content.json, or content.xml with styles.xml in XMind 8) open as a FreeMind
/// map with the first sheet's topics, folding, notes, links, labels, markers, pictures, boundaries, summaries, relationships,
/// detached topics and structure. Read-only: 'export file.xmind --to file.mm' writes the map.</summary>
public sealed class XmindAdapter : IFormatAdapter
{
    public string Format => "xmind";
    public IReadOnlyList<string> Extensions { get; } = [".xmind"];
    public bool CanWrite => false;

    public Document Create() =>
        throw new WriterException(ErrorCode.FormatReadonly, "xmind files are opened, not created", "Create a .mm map instead; an opened .xmind exports with --to map.mm.");

    public Document Open(Stream stream)
    {
        ZipArchive zip;
        try
        {
            zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        }
        catch (InvalidDataException ex)
        {
            throw new WriterException(ErrorCode.FormatError, $"Not an XMind file: {ex.Message}", "An .xmind file is a zip with content.json or content.xml inside.");
        }
        using (zip)
        {
            var sheet = zip.GetEntry("content.json") is { } json ? FromJson(json)
                : zip.GetEntry("content.xml") is { } xml ? FromXml(xml, zip.GetEntry("styles.xml"))
                : throw new WriterException(ErrorCode.FormatError, "Not an XMind file: no content.json or content.xml inside", "An .xmind file is a zip with content.json or content.xml inside.");
            return Build(sheet, zip);
        }
    }

    sealed class Topic
    {
        public string Id = "", Title = "", Notes = "", Href = "", Fill = "", Color = "", Size = "", Structure = "", Image = "", ImageSize = "", Position = "";
        public bool Folded, Bold, Italic;
        public readonly List<string> Labels = [], Markers = [], Boundaries = [];
        public readonly List<Topic> Children = [], Detached = [];
        public readonly List<(string Range, Topic Topic)> Summaries = [];
    }

    sealed record Sheet(Topic Root, List<(string From, string To, string Label)> Rels);

    // ----- content.json (XMind 2020 and later) -----
    static Sheet FromJson(ZipArchiveEntry entry)
    {
        using var s = entry.Open();
        using var doc = JsonDocument.Parse(s);
        var sheets = doc.RootElement.ValueKind == JsonValueKind.Array ? doc.RootElement.EnumerateArray().ToList() : [doc.RootElement];
        var sheet = sheets.FirstOrDefault(e => e.ValueKind == JsonValueKind.Object && e.TryGetProperty("rootTopic", out _));
        if (sheet.ValueKind != JsonValueKind.Object)
            throw new WriterException(ErrorCode.FormatError, "content.json has no sheet with a rootTopic", "Save the file again from XMind and retry.");
        var rels = new List<(string, string, string)>();
        if (sheet.TryGetProperty("relationships", out var rs) && rs.ValueKind == JsonValueKind.Array)
            foreach (var r in rs.EnumerateArray())
                if (Str(r, "end1Id") is { Length: > 0 } a && Str(r, "end2Id") is { Length: > 0 } b) rels.Add((a, b, Str(r, "title")));
        return new Sheet(JsonTopic(sheet.GetProperty("rootTopic")), rels);
    }

    static Topic JsonTopic(JsonElement t)
    {
        var x = new Topic { Id = Str(t, "id"), Title = Str(t, "title"), Href = Str(t, "href"), Structure = Structure(Str(t, "structureClass")), Folded = Str(t, "branch") == "folded" };
        if (t.TryGetProperty("notes", out var n) && n.ValueKind == JsonValueKind.Object && n.TryGetProperty("plain", out var p)) x.Notes = Str(p, "content");
        if (t.TryGetProperty("labels", out var ls) && ls.ValueKind == JsonValueKind.Array)
            x.Labels.AddRange(ls.EnumerateArray().Where(l => l.ValueKind == JsonValueKind.String).Select(l => l.GetString()!.Trim()).Where(l => l.Length > 0));
        if (t.TryGetProperty("markers", out var ms) && ms.ValueKind == JsonValueKind.Array)
            foreach (var m in ms.EnumerateArray()) if (Marker(Str(m, "markerId")) is { } mk) x.Markers.Add(mk);
        if (t.TryGetProperty("image", out var im) && im.ValueKind == JsonValueKind.Object)
        {
            x.Image = Str(im, "src");
            var (w, h) = (Num(im, "width"), Num(im, "height"));
            if (w > 0 && h > 0) x.ImageSize = $"{w},{h}";
        }
        if (t.TryGetProperty("style", out var st) && st.ValueKind == JsonValueKind.Object && st.TryGetProperty("properties", out var pr) && pr.ValueKind == JsonValueKind.Object) Style(x, name => Str(pr, name));
        if (t.TryGetProperty("position", out var pos) && pos.ValueKind == JsonValueKind.Object) x.Position = $"{Num(pos, "x")},{Num(pos, "y")}";
        if (t.TryGetProperty("children", out var ch) && ch.ValueKind == JsonValueKind.Object)
        {
            if (ch.TryGetProperty("attached", out var at) && at.ValueKind == JsonValueKind.Array) x.Children.AddRange(at.EnumerateArray().Select(JsonTopic));
            if (ch.TryGetProperty("detached", out var de) && de.ValueKind == JsonValueKind.Array) x.Detached.AddRange(de.EnumerateArray().Select(JsonTopic));
            if (ch.TryGetProperty("summary", out var su) && su.ValueKind == JsonValueKind.Array && t.TryGetProperty("summaries", out var sums) && sums.ValueKind == JsonValueKind.Array)
            {
                var pool = su.EnumerateArray().Select(JsonTopic).ToList();
                foreach (var s in sums.EnumerateArray())
                    if (pool.FirstOrDefault(q => q.Id == Str(s, "topicId")) is { } topic) x.Summaries.Add((Str(s, "range"), topic));
            }
        }
        if (t.TryGetProperty("boundaries", out var bs) && bs.ValueKind == JsonValueKind.Array)
            foreach (var b in bs.EnumerateArray()) if (Str(b, "range") is { Length: > 0 } r) x.Boundaries.Add(r);
        return x;
    }

    static string Str(JsonElement e, string name) => e.ValueKind == JsonValueKind.Object && e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";

    static int Num(JsonElement e, string name) => e.ValueKind == JsonValueKind.Object && e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? (int)Math.Round(v.GetDouble()) : 0;

    // ----- content.xml (XMind 8) -----
    static Sheet FromXml(ZipArchiveEntry entry, ZipArchiveEntry? stylesEntry)
    {
        XDocument xml;
        try
        {
            using var s = entry.Open();
            xml = XDocument.Load(s);
        }
        catch (System.Xml.XmlException ex)
        {
            throw new WriterException(ErrorCode.FormatError, $"content.xml could not be read: {ex.Message}", "Save the file again from XMind and retry.");
        }
        var styles = new Dictionary<string, XElement>();
        if (stylesEntry is not null)
        {
            using var ss = stylesEntry.Open();
            foreach (var st in XDocument.Load(ss).Descendants().Where(e => e.Name.LocalName == "style"))
                if (A(st, "id") is { Length: > 0 } id && El(st, "topic-properties") is { } tp) styles[id] = tp;
        }
        var sheet = xml.Root?.Descendants().FirstOrDefault(e => e.Name.LocalName == "sheet") ?? throw new WriterException(ErrorCode.FormatError, "content.xml has no sheet", "Save the file again from XMind and retry.");
        var rootEl = El(sheet, "topic") ?? throw new WriterException(ErrorCode.FormatError, "content.xml has a sheet without a topic", "Save the file again from XMind and retry.");
        var rels = sheet.Elements().Where(e => e.Name.LocalName == "relationships").SelectMany(r => r.Elements().Where(e => e.Name.LocalName == "relationship"))
            .Select(r => (A(r, "end1"), A(r, "end2"), Text(r, "title"))).Where(r => r.Item1.Length > 0 && r.Item2.Length > 0).ToList();
        return new Sheet(XmlTopic(rootEl, styles), rels);
    }

    static Topic XmlTopic(XElement t, Dictionary<string, XElement> styles)
    {
        var x = new Topic { Id = A(t, "id"), Title = Text(t, "title"), Href = A(t, "href"), Structure = Structure(A(t, "structure-class")), Folded = A(t, "branch") == "folded" };
        if (El(t, "notes") is { } notes) x.Notes = Text(notes, "plain");
        x.Labels.AddRange(Kids(El(t, "labels"), "label").Select(e => e.Value.Trim()).Where(l => l.Length > 0));
        foreach (var m in Kids(El(t, "marker-refs"), "marker-ref")) if (Marker(A(m, "marker-id")) is { } mk) x.Markers.Add(mk);
        if (El(t, "img") is { } img)
        {
            x.Image = A(img, "src");
            if (int.TryParse(A(img, "width"), out var w) && int.TryParse(A(img, "height"), out var h) && w > 0 && h > 0) x.ImageSize = $"{w},{h}";
        }
        if (A(t, "style-id") is { Length: > 0 } sid && styles.TryGetValue(sid, out var tp)) Style(x, name => A(tp, name.Split(':')[1]));
        if (El(t, "position") is { } pos) x.Position = $"{Round(A(pos, "x"))},{Round(A(pos, "y"))}";
        foreach (var group in Kids(El(t, "children"), "topics"))
        {
            var list = Kids(group, "topic").Select(e => XmlTopic(e, styles)).ToList();
            switch (A(group, "type"))
            {
                case "attached": x.Children.AddRange(list); break;
                case "detached": x.Detached.AddRange(list); break;
                case "summary":
                    foreach (var s in Kids(El(t, "summaries"), "summary"))
                        if (list.FirstOrDefault(q => q.Id == A(s, "topic-id")) is { } topic) x.Summaries.Add((A(s, "range"), topic));
                    break;
            }
        }
        foreach (var b in Kids(El(t, "boundaries"), "boundary")) if (A(b, "range") is { Length: > 0 } r) x.Boundaries.Add(r);
        return x;
    }

    static string A(XElement e, string name) => e.Attributes().FirstOrDefault(a => a.Name.LocalName == name)?.Value ?? "";

    static XElement? El(XElement? e, string name) => e?.Elements().FirstOrDefault(c => c.Name.LocalName == name);

    static IEnumerable<XElement> Kids(XElement? e, string name) => e?.Elements().Where(c => c.Name.LocalName == name) ?? [];

    static string Text(XElement e, string name) => El(e, name)?.Value.Trim() ?? "";

    static int Round(string v) => double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? (int)Math.Round(d) : 0;

    // ----- what both share -----
    static void Style(Topic x, Func<string, string> prop)
    {
        x.Fill = Hex(prop("svg:fill"));
        x.Color = Hex(prop("fo:color"));
        x.Bold = prop("fo:font-weight").Contains("bold", StringComparison.OrdinalIgnoreCase);
        x.Italic = prop("fo:font-style").Contains("italic", StringComparison.OrdinalIgnoreCase);
        var size = new string(prop("fo:font-size").TakeWhile(c => char.IsDigit(c)).ToArray());
        if (size.Length > 0) x.Size = size;
    }

    static string Hex(string v) => v.StartsWith('#') && v.Length == 7 ? v[1..].ToUpperInvariant() : "";

    /// <summary>XMind's structure class → the map's structure prop; empty for the default and the unknown.</summary>
    static string Structure(string cls) => cls switch
    {
        _ when cls.Contains("logic.right") => "logic",
        _ when cls.Contains("logic.left") => "logic-left",
        _ when cls.Contains("org-chart") => "org",
        _ when cls.Contains("tree") => "tree",
        _ when cls.Contains("timeline") => "timeline",
        _ => "",
    };

    /// <summary>XMind marker ids → the builtin names the map draws; null drops the marker.</summary>
    static string? Marker(string id) => id switch
    {
        _ when id.StartsWith("priority-") && id.Length == 10 && char.IsDigit(id[9]) && id[9] != '0' => "full-" + id[9],
        "task-start" => "0%", "task-oct" or "task-quarter" => "25%", "task-3oct" or "task-half" or "task-5oct" => "50%", "task-3quar" or "task-7oct" => "75%", "task-done" => "100%",
        "flag-red" => "flag", "flag-orange" => "flag-orange", "flag-yellow" => "flag-yellow", "flag-green" => "flag-green", "flag-blue" => "flag-blue", "flag-purple" => "flag-pink",
        _ when id.StartsWith("flag-") => "flag-black",
        "star-red" or "star-orange" => "star-red", "star-green" => "star-green", "star-blue" => "star-blue", "star-purple" => "star-purple",
        _ when id.StartsWith("star-") => "star",
        "symbol-question" => "help", "symbol-attention" or "symbol-exclam" => "messagebox_warning", "symbol-wrong" => "button_cancel", "symbol-right" => "button_ok", "symbol-info" => "info", "symbol-idea" or "c_symbol_idea" => "idea",
        _ => null,
    };

    static (int, int)? Range(string range, int count)
    {
        var parts = range.Trim('(', ')', ' ').Split(',');
        if (parts.Length != 2 || !int.TryParse(parts[0], out var i) || !int.TryParse(parts[1], out var j)) return null;
        (i, j) = (Math.Min(i, j), Math.Max(i, j));
        return i >= 0 && j < count ? (i, j) : null;
    }

    /// <summary>The FreeMind map: structure first (ids kept, so summaries and relationships still find their topics), then
    /// every other property through the mm adapter's own writers.</summary>
    static MmDocument Build(Sheet sheet, ZipArchive zip)
    {
        var seen = new HashSet<string>(); var n = 0;
        void Ids(Topic t) { if (t.Id.Length == 0 || !seen.Add(t.Id)) { do t.Id = "ID_" + ++n; while (!seen.Add(t.Id)); } foreach (var c in t.Children.Concat(t.Detached).Concat(t.Summaries.Select(s => s.Topic))) Ids(c); }
        Ids(sheet.Root);
        var map = new XElement("map", new XAttribute("version", "1.0.1"));
        var props = new List<(XElement Elem, string Name, string Value)>();
        var index = new Dictionary<string, XElement>();
        var rels = new Dictionary<string, List<(string To, string Label)>>();
        var detached = new List<(XElement Elem, string At)>();
        XElement Emit(Topic t)
        {
            var e = new XElement("node", new XAttribute("TEXT", t.Title), new XAttribute("ID", t.Id));
            if (t.Folded) e.SetAttributeValue("FOLDED", "true");
            index[t.Id] = e;
            void P(string name, string value) { if (value.Length > 0) props.Add((e, name, value)); }
            P("note", t.Notes); P("labels", string.Join(",", t.Labels)); P("icon", string.Join(",", t.Markers.Distinct()));
            P("fill", t.Fill); P("color", t.Color); P("size", t.Size); P("structure", t.Structure);
            if (t.Bold) P("bold", "true");
            if (t.Italic) P("italic", "true");
            if (t.Href.StartsWith("xmind:#")) { var to = t.Href[7..]; if (to.Length > 0) (rels[t.Id] = rels.GetValueOrDefault(t.Id) ?? []).Add((to, "")); }
            else P("link", t.Href);
            if (t.Image.Length > 0 && Picture(zip, t.Image) is { } uri) { P("image", uri); P("imageSize", t.ImageSize); }
            foreach (var c in t.Children) e.Add(Emit(c));
            foreach (var (range, topic) in t.Summaries)
            {
                var se = Emit(topic); e.Add(se);
                if (Range(range, t.Children.Count) is var (i, j)) props.Add((se, "summary", t.Children[i].Id + ":" + t.Children[j].Id));
            }
            foreach (var range in t.Boundaries)
                if (Range(range, t.Children.Count) is var (i, j)) for (var k = i; k <= j; k++) props.Add((index[t.Children[k].Id], "cloud", "CFE2F3"));
            foreach (var d in t.Detached) detached.Add((Emit(d), d.Position.Length > 0 ? d.Position : "0,0"));
            return e;
        }
        var root = Emit(sheet.Root);
        map.Add(root);
        foreach (var (elem, at) in detached) { root.Add(elem); props.Add((elem, "free", at)); } // floating topics hang off the centre wherever they were
        foreach (var (from, to, label) in sheet.Rels) (rels[from] = rels.GetValueOrDefault(from) ?? []).Add((to, label));
        foreach (var (from, list) in rels)
        {
            var known = list.Where(r => index.ContainsKey(r.To)).ToList();
            if (index.TryGetValue(from, out var elem) && known.Count > 0)
                props.Add((elem, "rels", NodeJson.Compact(w =>
                {
                    w.WriteStartArray();
                    foreach (var (to, label) in known) { w.WriteStartObject(); w.WriteString("to", to); w.WriteString("label", label); w.WriteEndObject(); }
                    w.WriteEndArray();
                })));
        }
        var doc = new MmDocument(new XDocument(map)) { FromXmind = true };
        foreach (var (elem, name, value) in props) new MmTopic(doc, elem).SetProp(name, value);
        return doc;
    }

    /// <summary>The picture behind an xap: reference as a data URI, so the map stays one file.</summary>
    static string? Picture(ZipArchive zip, string src)
    {
        var path = src.StartsWith("xap:") ? src[4..] : src;
        if (zip.GetEntry(path) is not { } entry) return null;
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        var mime = Path.GetExtension(path).ToLowerInvariant() switch { ".jpg" or ".jpeg" => "image/jpeg", ".gif" => "image/gif", ".svg" => "image/svg+xml", ".webp" => "image/webp", ".bmp" => "image/bmp", _ => "image/png" };
        return $"data:{mime};base64,{Convert.ToBase64String(ms.ToArray())}";
    }
}
