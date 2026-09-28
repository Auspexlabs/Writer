using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Cite;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>The document's sources, kept where Word keeps them: the bibliography part (customXml, b:Sources), which Word's References ›
/// Manage Sources lists and its citations and bibliography draw from. And the citation style the document is written in: the
/// document variable WriterCitationStyle, with the nearest style Word has written on b:Sources so Word's own style box agrees.</summary>
static class DocxSources
{
    static readonly XNamespace B = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
    const string StyleVariable = "WriterCitationStyle";
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    static readonly string[] MonthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

    /// <summary>The style names a document gives: mla, apa, chicago (notes and bibliography), chicago-date (author-date).</summary>
    public static readonly string[] StyleNames = ["mla", "apa", "chicago", "chicago-date", "gb7714", "ieee"];

    public static CiteStyle? StyleOf(string name) => name switch
    {
        "mla" => CiteStyle.Mla,
        "apa" => CiteStyle.Apa,
        "chicago" => CiteStyle.Chicago,
        "chicago-date" => CiteStyle.ChicagoDate,
        "gb7714" => CiteStyle.Gb7714,
        "ieee" => CiteStyle.Ieee,
        _ => null,
    };

    public static string NameOf(CiteStyle style) => StyleNames[(int)style];

    /// <summary>The document's citation style: its own (WriterCitationStyle), else what Word's style box was set to (Word's Chicago
    /// cites author-date), else none.</summary>
    public static CiteStyle? Style(DocxDocument doc)
    {
        var own = doc.Main.DocumentSettingsPart?.Settings?.GetFirstChild<W.DocumentVariables>()?.Elements<W.DocumentVariable>()
            .FirstOrDefault(v => v.Name?.Value == StyleVariable)?.Val?.Value;
        if (own is not null && StyleOf(own) is { } style) return style;
        var word = (string?)Root(doc)?.Attribute("StyleName") ?? (string?)Root(doc)?.Attribute("SelectedStyle") ?? "";
        return word.Contains("MLA", StringComparison.OrdinalIgnoreCase) ? CiteStyle.Mla
            : word.Contains("IEEE", StringComparison.OrdinalIgnoreCase) ? CiteStyle.Ieee
            : word.Contains("GB", StringComparison.OrdinalIgnoreCase) ? CiteStyle.Gb7714
            : word.Contains("APA", StringComparison.OrdinalIgnoreCase) ? CiteStyle.Apa
            : word.Contains("Chicago", StringComparison.OrdinalIgnoreCase) || word.Contains("Turabian", StringComparison.OrdinalIgnoreCase) ? CiteStyle.ChicagoDate
            : null;
    }

    public static void SetStyle(DocxDocument doc, string name)
    {
        var style = StyleOf(name) ?? throw new WriterException(ErrorCode.Validation, $"No citation style '{name}'", "Use mla, apa, chicago (notes and bibliography) or chicago-date (author-date).");
        var settings = (doc.Main.DocumentSettingsPart ?? doc.Main.AddNewPart<DocumentSettingsPart>()).Settings ??= new W.Settings();
        var variables = settings.GetFirstChild<W.DocumentVariables>();
        variables?.Elements<W.DocumentVariable>().Where(v => v.Name?.Value == StyleVariable).ToList().ForEach(v => v.Remove());
        if (variables is null) settings.AddChild(variables = new W.DocumentVariables());
        variables.Append(new W.DocumentVariable { Name = StyleVariable, Val = name });
        // Word's own style files: the ones every Word has, nearest to MLA 9, APA 7 and Chicago 18
        var (file, word, version) = style switch
        {
            CiteStyle.Gb7714 => ("\\GB7714.XSL", "GB7714", "2015"),
            CiteStyle.Ieee => ("\\IEEE.XSL", "IEEE", "2006"),
            CiteStyle.Mla => ("\\MLASeventhEditionOfficeOnline.xsl", "MLA", "7"),
            CiteStyle.Apa => ("\\APASixthEditionOfficeOnline.xsl", "APA", "6"),
            _ => ("\\CHICAGO.XSL", "Chicago", "16"),
        };
        Edit(doc, root =>
        {
            root.SetAttributeValue("SelectedStyle", file);
            root.SetAttributeValue("StyleName", word);
            root.SetAttributeValue("Version", version);
        });
    }

    // ----- the list -----

    public static List<Source> Read(DocxDocument doc) => Root(doc)?.Elements(B + "Source").Select(FromXml).Where(s => s.Tag.Length > 0).ToList() ?? [];

    /// <summary>The sources as JSON: an array of objects with tag, type, authors ([{last, first}] or [{name}] for an organisation),
    /// editors, title, container, publisher, place, year, month, day, volume, issue, pages, edition, url, doi, isbn, accessed.</summary>
    public static string Json(IEnumerable<Source> sources) => NodeJson.Compact(w =>
    {
        w.WriteStartArray();
        foreach (var s in sources) Write(w, s);
        w.WriteEndArray();
    });

    static void Write(Utf8JsonWriter w, Source s)
    {
        w.WriteStartObject();
        w.WriteString("tag", s.Tag);
        w.WriteString("type", s.Type);
        People(w, "authors", s.Authors);
        People(w, "editors", s.Editors);
        foreach (var (name, value) in Fields(s)) if (value.Length > 0) w.WriteString(name, value);
        w.WriteEndObject();
    }

    static void People(Utf8JsonWriter w, string name, IReadOnlyList<Person> people)
    {
        if (people.Count == 0) return;
        w.WriteStartArray(name);
        foreach (var p in people)
        {
            w.WriteStartObject();
            if (p.Org) w.WriteString("name", p.Last);
            else
            {
                w.WriteString("last", p.Last);
                if (p.First.Length > 0) w.WriteString("first", p.First);
            }
            w.WriteEndObject();
        }
        w.WriteEndArray();
    }

    static IEnumerable<(string Name, string Value)> Fields(Source s) =>
    [
        ("title", s.Title), ("container", s.Container), ("publisher", s.Publisher), ("place", s.Place), ("year", s.Year), ("month", s.Month), ("day", s.Day),
        ("volume", s.Volume), ("issue", s.Issue), ("pages", s.Pages), ("edition", s.Edition), ("url", s.Url), ("doi", s.Doi), ("isbn", s.Isbn), ("accessed", s.Accessed),
    ];

    /// <summary>Adds or changes a source from JSON (the fields of Json; a field given as "" or none is cleared, one left out stays), or
    /// removes one ({"tag": …, "remove": true}), which a citation still naming it forbids. A new source without a tag gets one as Word
    /// makes them: the first author's surname, three letters, and the year (Peg15). Returns the tag.</summary>
    public static string Define(DocxDocument doc, string json, Func<string, bool> cited)
    {
        using var parsed = JsonDocument.Parse(json);
        var o = parsed.RootElement;
        if (o.ValueKind != JsonValueKind.Object) throw new WriterException(ErrorCode.Validation, "A source is a JSON object", "Give {\"type\":\"article\",\"authors\":[{\"last\":\"Pegg\",\"first\":\"Ian L.\"}],\"title\":\"…\",\"year\":\"2015\"}.");
        var tag = Str(o, "tag") ?? "";
        var existing = Read(doc).FirstOrDefault(s => string.Equals(s.Tag, tag, StringComparison.OrdinalIgnoreCase));
        if (o.TryGetProperty("remove", out var remove) && remove.ValueKind == JsonValueKind.True)
        {
            if (existing is null) throw new WriterException(ErrorCode.PathNotFound, $"No source '{tag}'", "Read the document's sources to see their tags.");
            if (cited(existing.Tag)) throw new WriterException(ErrorCode.Validation, $"Source '{existing.Tag}' is cited in the document", "Remove its citations first, as Word asks.");
            Edit(doc, root => root.Elements(B + "Source").Where(e => string.Equals((string?)e.Element(B + "Tag"), existing.Tag, StringComparison.OrdinalIgnoreCase)).Remove());
            return existing.Tag;
        }
        var s = Apply(existing ?? new Source(), o);
        if (s.Title.Length == 0 && s.Authors.Count == 0) throw new WriterException(ErrorCode.Validation, "A source needs a title or an author", "Give at least its title.");
        if (existing is null) s = s with { Tag = tag.Length > 0 ? Unique(doc, tag) : Unique(doc, NewTag(s)) };
        Edit(doc, root =>
        {
            var element = root.Elements(B + "Source").FirstOrDefault(x => string.Equals((string?)x.Element(B + "Tag"), s.Tag, StringComparison.OrdinalIgnoreCase));
            if (element is null) root.Add(element = new XElement(B + "Source", new XElement(B + "Tag", s.Tag), new XElement(B + "Guid", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}")));
            ToXml(element, s);
        });
        return s.Tag;
    }

    /// <summary>A source from JSON on its own (the fields of Json), its tag made as Word makes one when it has none: what the cite
    /// command formats.</summary>
    public static Source Parse(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Object) throw new WriterException(ErrorCode.Validation, "A source is a JSON object", "Give {\"type\":\"article\",\"title\":\"…\"}.");
        var s = Apply(new Source(), parsed.RootElement);
        return s with { Tag = Str(parsed.RootElement, "tag") is { Length: > 0 } tag ? tag : NewTag(s) };
    }

    /// <summary>A source with the fields of a JSON object written: one given as "" or none is cleared, one left out stays.</summary>
    static Source Apply(Source s, JsonElement o)
    {
        string Field(string name, string was) => Str(o, name) is { } v ? (v == "none" ? "" : v.Trim()) : was;
        return s with
        {
            Type = Field("type", s.Type) is { Length: > 0 } t ? Normalize(t) : "article",
            Authors = o.TryGetProperty("authors", out var a) ? PeopleOf(a) : s.Authors,
            Editors = o.TryGetProperty("editors", out var e) ? PeopleOf(e) : s.Editors,
            Title = Field("title", s.Title), Container = Field("container", s.Container), Publisher = Field("publisher", s.Publisher), Place = Field("place", s.Place),
            Year = Field("year", s.Year), Month = Field("month", s.Month), Day = Field("day", s.Day), Volume = Field("volume", s.Volume), Issue = Field("issue", s.Issue),
            Pages = Field("pages", s.Pages), Edition = Field("edition", s.Edition), Url = Field("url", s.Url), Doi = Field("doi", s.Doi), Isbn = Field("isbn", s.Isbn),
            Accessed = Field("accessed", s.Accessed),
        };
    }

    static string? Str(JsonElement o, string name) => o.TryGetProperty(name, out var v) ? v.ValueKind switch
    {
        JsonValueKind.String => v.GetString(),
        JsonValueKind.Number => v.GetRawText(),
        JsonValueKind.Null => "",
        _ => null,
    } : null;

    static List<Person> PeopleOf(JsonElement list)
    {
        if (list.ValueKind == JsonValueKind.String) return (list.GetString() ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Select(ParseName).ToList();
        if (list.ValueKind != JsonValueKind.Array) return [];
        var people = new List<Person>();
        foreach (var p in list.EnumerateArray())
        {
            if (p.ValueKind == JsonValueKind.String) { people.Add(ParseName(p.GetString() ?? "")); continue; }
            if (Str(p, "name") is { Length: > 0 } org) people.Add(new(org.Trim(), "", true));
            else if (Str(p, "last") is { Length: > 0 } last) people.Add(new(last.Trim(), (Str(p, "first") ?? "").Trim()));
        }
        return people;
    }

    /// <summary>A name as people write it: "Pegg, Ian L." or "Ian L. Pegg".</summary>
    static Person ParseName(string name)
    {
        var comma = name.IndexOf(',');
        if (comma > 0) return new(name[..comma].Trim(), name[(comma + 1)..].Trim());
        var parts = name.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length < 2 ? new(name.Trim(), "", parts.Length == 0) : new(parts[^1], string.Join(" ", parts[..^1]));
    }

    static string Normalize(string type) => type.Trim().ToLowerInvariant() switch
    {
        "journal" or "journal-article" or "journalarticle" => "article",
        "book-chapter" or "booksection" or "section" => "chapter",
        "website" or "web" or "internetsite" or "web page" => "webpage",
        "proceedings" or "proceedings-article" or "paper" => "conference",
        "dissertation" => "thesis",
        var t when t is "article" or "book" or "chapter" or "magazine" or "newspaper" or "webpage" or "report" or "conference" or "thesis" or "other" => t,
        _ => "other",
    };

    static string NewTag(Source s)
    {
        var who = s.Authors.Count > 0 ? s.Authors[0].Last : s.Title;
        var letters = new string(who.Where(char.IsLetter).Take(3).ToArray());
        if (letters.Length == 0) letters = "Src";
        var year = s.Year.Length >= 2 && s.Year.All(char.IsDigit) ? s.Year[^2..] : "";
        return char.ToUpperInvariant(letters[0]) + letters[1..].ToLowerInvariant() + year;
    }

    static string Unique(DocxDocument doc, string tag)
    {
        var taken = Read(doc).Select(s => s.Tag).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (!taken.Contains(tag)) return tag;
        for (var i = 1; ; i++) if (!taken.Contains(tag + i.ToString(Inv))) return tag + i.ToString(Inv);
    }

    // ----- b:Source -----

    static readonly (string Word, string Ours)[] Types =
    [
        ("JournalArticle", "article"), ("Book", "book"), ("BookSection", "chapter"), ("ArticleInAPeriodical", "magazine"), ("InternetSite", "webpage"),
        ("DocumentFromInternetSite", "webpage"), ("Report", "report"), ("ConferenceProceedings", "conference"), ("Misc", "other"),
    ];

    /// <summary>Where each type keeps what holds it: a journal's name, the book of a chapter, a website's, a periodical's, a conference's.</summary>
    static string ContainerElement(string type) => type switch
    {
        "article" => "JournalName",
        "chapter" => "BookTitle",
        "magazine" or "newspaper" => "PeriodicalTitle",
        "webpage" => "InternetSiteTitle",
        "conference" => "ConferenceName",
        _ => "PublicationTitle",
    };

    static Source FromXml(XElement e)
    {
        string V(string name) => ((string?)e.Element(B + name) ?? "").Trim();
        var word = V("SourceType");
        var type = Types.FirstOrDefault(t => t.Word.Equals(word, StringComparison.OrdinalIgnoreCase)).Ours ?? "other";
        if (type == "report" && V("ThesisType").Length > 0) type = "thesis";
        var container = V(ContainerElement(type));
        if (container.Length == 0) container = new[] { "JournalName", "BookTitle", "PeriodicalTitle", "InternetSiteTitle", "ConferenceName", "PublicationTitle" }.Select(V).FirstOrDefault(v => v.Length > 0) ?? "";
        var accessed = V("YearAccessed").Length > 0 ? AccessedOf(V("YearAccessed"), V("MonthAccessed"), V("DayAccessed")) : "";
        return new Source
        {
            Tag = V("Tag"), Type = type, Authors = Names(e, "Author"), Editors = Names(e, "Editor"), Title = V("Title"), Container = container,
            Publisher = V("Publisher") is { Length: > 0 } pub ? pub : V("Institution"), Place = V("City"), Year = V("Year"), Month = V("Month"), Day = V("Day"),
            Volume = V("Volume"), Issue = V("Issue"), Pages = V("Pages"), Edition = V("Edition"), Url = V("URL"), Doi = V("DOI"), Isbn = V("StandardNumber"), Accessed = accessed,
        };
    }

    static List<Person> Names(XElement source, string role)
    {
        var holder = source.Element(B + "Author")?.Element(B + role);
        if (holder is null) return [];
        if (holder.Element(B + "Corporate") is { } corporate) return [new(corporate.Value.Trim(), "", true)];
        return holder.Element(B + "NameList")?.Elements(B + "Person").Select(p =>
        {
            var first = string.Join(" ", new[] { (string?)p.Element(B + "First"), (string?)p.Element(B + "Middle") }.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x!.Trim()));
            return new Person(((string?)p.Element(B + "Last") ?? "").Trim(), first);
        }).Where(p => p.Last.Length > 0).ToList() ?? [];
    }

    static string AccessedOf(string year, string month, string day)
    {
        var m = int.TryParse(month, NumberStyles.Integer, Inv, out var n) ? n : Array.FindIndex(MonthNames, x => x.StartsWith(month, StringComparison.OrdinalIgnoreCase) && month.Length >= 3) + 1;
        return m is >= 1 and <= 12 && int.TryParse(day, NumberStyles.Integer, Inv, out var d) ? $"{year}-{m:00}-{d:00}" : year;
    }

    /// <summary>Writes a source's fields onto its b:Source, leaving the elements Word keeps that these do not name.</summary>
    static void ToXml(XElement e, Source s)
    {
        void Set(string name, string value)
        {
            e.Elements(B + name).Remove();
            if (value.Length > 0) e.Add(new XElement(B + name, value));
        }
        Set("SourceType", Types.First(t => t.Ours == (s.Type == "newspaper" ? "magazine" : s.Type == "thesis" ? "report" : s.Type)).Word);
        foreach (var name in new[] { "JournalName", "BookTitle", "PeriodicalTitle", "InternetSiteTitle", "ConferenceName", "PublicationTitle" }) e.Elements(B + name).Remove();
        Set(ContainerElement(s.Type), s.Container);
        Set("ThesisType", s.Type == "thesis" ? "Thesis" : "");
        var names = new XElement(B + "Author");
        void Role(string role, IReadOnlyList<Person> people)
        {
            if (people.Count == 0) return;
            if (people.Count == 1 && people[0].Org) { names.Add(new XElement(B + role, new XElement(B + "Corporate", people[0].Last))); return; }
            names.Add(new XElement(B + role, new XElement(B + "NameList", people.Select(p =>
            {
                var person = new XElement(B + "Person", new XElement(B + "Last", p.Last));
                var given = p.First.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                if (given.Length > 0) person.Add(new XElement(B + "First", given[0]));
                if (given.Length > 1) person.Add(new XElement(B + "Middle", given[1]));
                return person;
            }))));
        }
        Role("Author", s.Authors);
        Role("Editor", s.Editors);
        var old = e.Element(B + "Author");
        foreach (var keep in old?.Elements().Where(x => x.Name != B + "Author" && x.Name != B + "Editor") ?? []) names.Add(keep); // translators, composers…
        old?.Remove();
        if (names.HasElements) e.Add(names);
        Set("Title", s.Title);
        Set("Publisher", s.Publisher);
        e.Elements(B + "Institution").Remove();
        Set("City", s.Place);
        Set("Year", s.Year);
        Set("Month", int.TryParse(s.Month, NumberStyles.Integer, Inv, out var m) && m is >= 1 and <= 12 ? MonthNames[m - 1] : s.Month);
        Set("Day", s.Day);
        Set("Volume", s.Volume);
        Set("Issue", s.Issue);
        Set("Pages", s.Pages);
        Set("Edition", s.Edition);
        Set("URL", s.Url);
        Set("DOI", s.Doi);
        Set("StandardNumber", s.Isbn);
        var accessed = System.Text.RegularExpressions.Regex.Match(s.Accessed, @"^(\d{4})(?:-(\d{1,2})-(\d{1,2}))?");
        Set("YearAccessed", accessed.Success ? accessed.Groups[1].Value : "");
        Set("MonthAccessed", accessed.Success && accessed.Groups[2].Success ? MonthNames[int.Parse(accessed.Groups[2].Value, Inv) - 1] : "");
        Set("DayAccessed", accessed.Success && accessed.Groups[3].Success ? int.Parse(accessed.Groups[3].Value, Inv).ToString(Inv) : "");
    }

    // ----- the part -----

    /// <summary>The bibliography part's root, b:Sources, when the document has one.</summary>
    static XElement? Root(DocxDocument doc) => Part(doc) is { } part ? Load(part)?.Root : null;

    static CustomXmlPart? Part(DocxDocument doc) => doc.Main.CustomXmlParts.FirstOrDefault(p => Load(p)?.Root?.Name == B + "Sources");

    static XDocument? Load(CustomXmlPart part)
    {
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XDocument.Load(stream);
        }
        catch (System.Xml.XmlException) { return null; }
    }

    /// <summary>Changes the bibliography part, made as Word makes it (with its data store item naming the bibliography schema) when the
    /// document has none.</summary>
    static void Edit(DocxDocument doc, Action<XElement> change)
    {
        var part = Part(doc);
        // made in MLA, the style a document without one is written in (Word would make it APA)
        var xml = part is not null ? Load(part)! : new XDocument(new XElement(B + "Sources", new XAttribute(XNamespace.Xmlns + "b", B.NamespaceName), new XAttribute("xmlns", B.NamespaceName),
            new XAttribute("SelectedStyle", "\\MLASeventhEditionOfficeOnline.xsl"), new XAttribute("StyleName", "MLA"), new XAttribute("Version", "7")));
        change(xml.Root!);
        if (part is null)
        {
            part = doc.Main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            var props = part.AddNewPart<CustomXmlPropertiesPart>();
            props.DataStoreItem = new DataStoreItem(new SchemaReferences(new SchemaReference { Uri = B.NamespaceName })) { ItemId = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}" };
        }
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        xml.Save(stream, SaveOptions.DisableFormatting);
    }
}
