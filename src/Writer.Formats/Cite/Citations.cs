using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Writer.Formats.Cite;

/// <summary>The citation styles: MLA 9th edition, APA 7th, and Chicago 18th in its two systems, notes and bibliography (Chicago)
/// and author-date (ChicagoDate).</summary>
public enum CiteStyle { Mla, Apa, Chicago, ChicagoDate }

/// <summary>A person as a source names them, or an organisation (Org) with its whole name in Last.</summary>
public sealed record Person(string Last, string First = "", bool Org = false);

/// <summary>A source of the document's source list: the fields Word's source manager keeps (b:Source), under names of their own.
/// Type is article (a journal), book, chapter (in an edited book), magazine, newspaper, webpage, report, conference, thesis or
/// other; Container is what holds the source: the journal, the book of a chapter, the website, the newspaper, the conference.</summary>
public sealed record Source
{
    public string Tag { get; init; } = "";
    public string Type { get; init; } = "article";
    public IReadOnlyList<Person> Authors { get; init; } = [];
    public IReadOnlyList<Person> Editors { get; init; } = [];
    public string Title { get; init; } = "";
    public string Container { get; init; } = "";
    public string Publisher { get; init; } = "";
    public string Place { get; init; } = "";
    public string Year { get; init; } = "";
    public string Month { get; init; } = "";
    public string Day { get; init; } = "";
    public string Volume { get; init; } = "";
    public string Issue { get; init; } = "";
    public string Pages { get; init; } = "";
    public string Edition { get; init; } = "";
    public string Url { get; init; } = "";
    public string Doi { get; init; } = "";
    public string Isbn { get; init; } = "";
    /// <summary>When a web source was read, yyyy-mm-dd.</summary>
    public string Accessed { get; init; } = "";
}

/// <summary>A stretch of formatted text: citations and entries italicise titles.</summary>
public readonly record struct Span(string Text, bool Italic = false);

/// <summary>One source a citation names, with the place in it (a page, a range, "para. 4").</summary>
public sealed record Cited(Source Source, string Locator = "");

/// <summary>Formats sources in a citation style: the entries of the works-cited list (bibliography), the citations in the text, and
/// Chicago's notes. What depends on the whole list (APA's 2015a and 2015b, MLA's short title for an author of two works, initials for
/// two authors of one surname) comes from the list given to the constructor.</summary>
public sealed class Citations
{
    readonly CiteStyle _style;
    readonly Dictionary<string, string> _yearSuffix = new(StringComparer.Ordinal); // tag → a, b… (author-date styles)
    readonly HashSet<string> _manyWorks = new(StringComparer.Ordinal); // tags of sources whose authors have another work in the list
    readonly HashSet<string> _sharedSurname = new(StringComparer.OrdinalIgnoreCase); // surnames of two different first authors

    public Citations(CiteStyle style, IEnumerable<Source> sources)
    {
        _style = style;
        var list = sources.ToList();
        foreach (var group in list.GroupBy(AuthorKey).Where(g => g.Key.Length > 0 && g.Count() > 1))
        {
            foreach (var s in group) _manyWorks.Add(s.Tag);
            if (style is CiteStyle.Apa or CiteStyle.ChicagoDate)
                foreach (var same in group.GroupBy(s => YearOf(s)).Where(g => g.Count() > 1))
                {
                    var i = 0;
                    foreach (var s in same.OrderBy(s => SortTitle(s.Title), StringComparer.Ordinal)) _yearSuffix[s.Tag] = ((char)('a' + i++)).ToString();
                }
        }
        foreach (var g in list.Where(s => s.Authors.Count > 0 && !s.Authors[0].Org).GroupBy(s => Fold(s.Authors[0].Last)))
            if (g.Select(s => Fold(s.Authors[0].First)).Distinct().Count() > 1) _sharedSurname.Add(g.Key);
    }

    public CiteStyle Style => _style;

    /// <summary>The title a works-cited list goes under.</summary>
    public static string ListTitle(CiteStyle style) => style switch
    {
        CiteStyle.Mla => "Works Cited",
        CiteStyle.Apa => "References",
        CiteStyle.ChicagoDate => "References",
        _ => "Bibliography",
    };

    /// <summary>Whether the style cites in notes (Chicago's notes and bibliography) rather than in parentheses.</summary>
    public static bool InNotes(CiteStyle style) => style == CiteStyle.Chicago;

    /// <summary>The list in the style's order: by author (else title), then year or title.</summary>
    public IEnumerable<Source> Sort(IEnumerable<Source> sources) => sources.OrderBy(s => SortKey(s), StringComparer.Ordinal);

    string SortKey(Source s)
    {
        var who = s.Authors.Count > 0 ? string.Join(" ", s.Authors.Select(a => Fold(a.Last) + " " + Fold(a.First))) : SortTitle(s.Title);
        var year = YearOf(s).PadLeft(4, '0') + (_yearSuffix.GetValueOrDefault(s.Tag) ?? "");
        return _style is CiteStyle.Apa or CiteStyle.ChicagoDate ? who + "\u0001" + year + "\u0001" + SortTitle(s.Title) : who + "\u0001" + SortTitle(s.Title) + "\u0001" + year;
    }

    // ----- the works-cited list -----

    /// <summary>A source's entry in the works-cited list (Works Cited, References, Bibliography).</summary>
    public List<Span> Entry(Source s) => _style switch
    {
        CiteStyle.Mla => MlaEntry(s),
        CiteStyle.Apa => ApaEntry(s),
        CiteStyle.ChicagoDate => ChicagoEntry(s, dated: true),
        _ => ChicagoEntry(s, dated: false),
    };

    List<Span> MlaEntry(Source s)
    {
        var o = new Out();
        if (s.Authors.Count > 0) o.Add(Names(s.Authors, NameOrder.FirstInverted, "and", etAlFrom: 3, etAlKeep: 1)).End('.').Add(" ");
        if (Italic(s)) o.Italic(TitleCase(s.Title)).End('.'); else o.Quoted(TitleCase(s.Title), '.');
        // the containers' elements, each ending in a comma, the last in a period
        var parts = new List<List<Span>>();
        if (s.Container.Length > 0) parts.Add([new(TitleCase(s.Container), true)]);
        if (s.Editors.Count > 0) parts.Add([new("edited by " + Names(s.Editors, NameOrder.Normal, "and", 3, 1))]);
        if (s.Edition.Length > 0) parts.Add([new(Ordinal(s.Edition) + " ed.")]);
        if (s.Volume.Length > 0) parts.Add([new("vol. " + s.Volume)]);
        if (s.Issue.Length > 0) parts.Add([new("no. " + Range(s.Issue))]);
        if (s.Publisher.Length > 0 && s.Type is not ("article" or "magazine" or "newspaper") && !(s.Type == "webpage" && Same(s.Publisher, s.Container)))
            parts.Add([new(s.Publisher)]);
        var date = MlaDate(s);
        if (date.Length > 0) parts.Add([new(date)]);
        if (s.Pages.Length > 0) parts.Add([new((IsRange(s.Pages) ? "pp. " : "p. ") + MlaRange(s.Pages))]);
        if (s.Doi.Length > 0) parts.Add([new(DoiUrl(s.Doi))]);
        else if (s.Url.Length > 0) parts.Add([new(Regex.Replace(s.Url.Trim(), "^https?://", ""))]);
        for (var i = 0; i < parts.Count; i++) o.Add(" ").Add(parts[i]).End(i == parts.Count - 1 ? '.' : ',');
        if (date.Length == 0 && DateOf(s.Accessed) is { } seen) o.Add(" Accessed " + MlaDay(seen)).End('.');
        return o.Spans;
    }

    List<Span> ApaEntry(Source s)
    {
        var o = new Out();
        var date = ApaDate(s);
        if (s.Authors.Count > 0) o.Add(ApaNames(s.Authors)).Add(" (" + date + ").");
        else
        {
            // no author: the title moves to the author's place
            if (Italic(s)) o.Italic(s.Title); else o.Add(s.Title);
            if (s.Edition.Length > 0 && s.Type is "book" or "report") o.Add(" (" + Ordinal(s.Edition) + " ed.)");
            o.End('.').Add(" (" + date + ").");
        }
        if (s.Authors.Count > 0)
        {
            o.Add(" ");
            if (Italic(s)) o.Italic(s.Title); else o.Add(s.Title);
            if (s.Edition.Length > 0 && s.Type is "book" or "report") o.Add(" (" + Ordinal(s.Edition) + " ed.)");
            o.End('.');
        }
        switch (s.Type)
        {
            case "article" or "magazine" or "newspaper":
                if (s.Container.Length > 0)
                {
                    o.Add(" ").Italic(s.Container);
                    if (s.Volume.Length > 0) o.Add(", ").Italic(s.Volume);
                    if (s.Issue.Length > 0) o.Add("(" + Range(s.Issue) + ")");
                    if (s.Pages.Length > 0) o.Add(", " + Range(s.Pages));
                    o.End('.');
                }
                break;
            case "chapter":
                o.Add(" In ");
                if (s.Editors.Count > 0) o.Add(ApaEditors(s.Editors) + ", ");
                o.Italic(s.Container.Length > 0 ? s.Container : "");
                if (s.Pages.Length > 0) o.Add(" (" + (IsRange(s.Pages) ? "pp. " : "p. ") + Range(s.Pages) + ")");
                o.End('.');
                if (s.Publisher.Length > 0) o.Add(" " + s.Publisher).End('.');
                break;
            case "webpage":
                if (s.Container.Length > 0 && !(s.Authors.Count == 1 && s.Authors[0].Org && Same(s.Authors[0].Last, s.Container))) o.Add(" " + s.Container).End('.');
                break;
            default:
                if (s.Container.Length > 0 && s.Type == "conference") o.Add(" ").Italic(s.Container).End('.');
                if (s.Publisher.Length > 0 && !(s.Authors.Count == 1 && Same(s.Authors[0].Last, s.Publisher))) o.Add(" " + s.Publisher).End('.');
                break;
        }
        if (s.Doi.Length > 0) o.Add(" " + DoiUrl(s.Doi));
        else if (s.Url.Length > 0)
        {
            if (s.Type == "webpage" && ApaDate(s) == "n.d." && DateOf(s.Accessed) is { } seen) o.Add(" Retrieved " + LongDate(seen) + ", from");
            o.Add(" " + s.Url);
        }
        return o.Spans;
    }

    List<Span> ChicagoEntry(Source s, bool dated)
    {
        var o = new Out();
        var year = YearOf(s, "n.d.") + (_yearSuffix.GetValueOrDefault(s.Tag) ?? "");
        void Title() { if (Italic(s)) o.Italic(TitleCase(s.Title)).End('.'); else o.Quoted(TitleCase(s.Title), '.'); }
        if (s.Authors.Count > 0)
        {
            o.Add(Names(s.Authors, NameOrder.FirstInverted, "and", etAlFrom: 7, etAlKeep: 3)).End('.');
            if (dated) o.Add(" " + year).End('.');
            o.Add(" ");
            Title();
        }
        else
        {
            Title(); // no author: the title leads, the year after it
            if (dated) o.Add(" " + year).End('.');
        }
        switch (s.Type)
        {
            case "article":
                if (s.Container.Length > 0)
                {
                    o.Add(" ").Italic(TitleCase(s.Container));
                    if (s.Volume.Length > 0) o.Add(" " + s.Volume);
                    if (dated) { if (s.Issue.Length > 0) o.Add(" (" + Range(s.Issue) + ")"); }
                    else
                    {
                        if (s.Issue.Length > 0) o.Add(", no. " + Range(s.Issue));
                        o.Add(" (" + ChicagoMonthYear(s) + ")");
                    }
                    o.Add(s.Pages.Length > 0 ? ": " + Range(s.Pages) : "").End('.');
                }
                break;
            case "magazine" or "newspaper":
                if (s.Container.Length > 0) o.Add(" ").Italic(TitleCase(s.Container)).End(dated ? '.' : ',');
                if (!dated && ChicagoDay(s) is { Length: > 0 } day) o.Add(" " + day).End('.');
                else if (dated && ChicagoMonthDay(s) is { Length: > 0 } md) o.Add(" " + md).End('.');
                break;
            case "chapter":
                if (s.Container.Length > 0)
                {
                    o.Add(" In ").Italic(TitleCase(s.Container));
                    if (s.Editors.Count > 0) o.Add(", edited by " + Names(s.Editors, NameOrder.Normal, "and", 7, 3));
                    if (s.Pages.Length > 0) o.Add(", " + Range(s.Pages));
                    o.End('.');
                }
                if (Imprint(s, dated) is { Length: > 0 } imprint) o.Add(" " + imprint).End('.');
                break;
            case "webpage":
                if (s.Container.Length > 0) o.Add(" " + s.Container).End('.');
                if (!dated && ChicagoDay(s) is { Length: > 0 } day2) o.Add(" " + day2).End('.');
                else if (dated && ChicagoMonthDay(s) is { Length: > 0 } md2) o.Add(" " + md2).End('.');
                else if (DateOf(s.Accessed) is { } seen && YearOf(s).Length == 0) o.Add(" Accessed " + LongDate(seen)).End('.');
                break;
            default:
                if (s.Edition.Length > 0) o.Add(" " + Ordinal(s.Edition) + " ed").End('.');
                if (s.Type == "conference" && s.Container.Length > 0) o.Add(" Paper presented at " + s.Container).End('.');
                if (Imprint(s, dated) is { Length: > 0 } imprint2) o.Add(" " + imprint2).End('.');
                break;
        }
        if (s.Doi.Length > 0) o.Add(" " + DoiUrl(s.Doi)).End('.');
        else if (s.Url.Length > 0) o.Add(" " + s.Url).End('.');
        return o.Spans;
    }

    /// <summary>Chicago's place: publisher, year (the year stays out in author-date, where it follows the author).</summary>
    static string Imprint(Source s, bool dated)
    {
        var who = s.Place.Length > 0 && s.Publisher.Length > 0 ? s.Place + ": " + s.Publisher : s.Publisher.Length > 0 ? s.Publisher : s.Place;
        if (dated) return who;
        var year = YearOf(s);
        return who.Length > 0 && year.Length > 0 ? who + ", " + year : who + year;
    }

    // ----- citations in the text -----

    /// <summary>A citation in parentheses (MLA, APA, Chicago author-date) of one or more sources. noAuthor leaves the name out, as when
    /// the sentence names the author (Pegg shows … (287)); noYear the year.</summary>
    public List<Span> InText(IReadOnlyList<Cited> cites, bool noAuthor = false, bool noYear = false)
    {
        var o = new Out();
        o.Add("(");
        var ordered = _style == CiteStyle.Apa && cites.Count > 1 ? cites.OrderBy(c => SortKey(c.Source), StringComparer.Ordinal).ToList() : cites.ToList();
        for (var i = 0; i < ordered.Count; i++)
        {
            if (i > 0) o.Add("; ");
            o.Add(_style switch
            {
                CiteStyle.Mla => MlaCite(ordered[i], noAuthor),
                CiteStyle.Apa => ApaCite(ordered[i], noAuthor, noYear),
                _ => ChicagoDateCite(ordered[i], noAuthor, noYear),
            });
        }
        o.Add(")");
        return o.Spans;
    }

    List<Span> MlaCite(Cited c, bool noAuthor)
    {
        var s = c.Source;
        var o = new Out();
        var who = noAuthor ? "" : ShortNames(s, "and", etAlFrom: 3);
        if (who.Length > 0) o.Add(who);
        if (who.Length == 0 && !noAuthor || _manyWorks.Contains(s.Tag) && !noAuthor)
        {
            if (who.Length > 0) o.Add(", ");
            ShortTitle(o, s);
        }
        if (c.Locator.Length > 0) o.Add((o.Length > 0 ? " " : "") + MlaRange(c.Locator));
        return o.Spans;
    }

    List<Span> ApaCite(Cited c, bool noAuthor, bool noYear)
    {
        var s = c.Source;
        var o = new Out();
        if (!noAuthor)
        {
            var who = ShortNames(s, "&", etAlFrom: 3);
            if (who.Length > 0) o.Add(who); else ShortTitle(o, s);
        }
        if (!noYear) o.Add((o.Length > 0 ? ", " : "") + ApaYear(s));
        if (c.Locator.Length > 0) o.Add((o.Length > 0 ? ", " : "") + Locator(c.Locator));
        return o.Spans;
    }

    List<Span> ChicagoDateCite(Cited c, bool noAuthor, bool noYear)
    {
        var s = c.Source;
        var o = new Out();
        if (!noAuthor)
        {
            var who = ShortNames(s, "and", etAlFrom: 3);
            if (who.Length > 0) o.Add(who); else ShortTitle(o, s);
        }
        if (!noYear) o.Add((o.Length > 0 ? " " : "") + YearOf(s, "n.d.") + (_yearSuffix.GetValueOrDefault(s.Tag) ?? ""));
        if (c.Locator.Length > 0) o.Add((o.Length > 0 ? ", " : "") + Range(c.Locator));
        return o.Spans;
    }

    // ----- Chicago's notes -----

    /// <summary>A Chicago note citing a source: the full note the first time the source is cited, a short note (author, short title,
    /// page) after that. Several sources are separated by semicolons.</summary>
    public List<Span> Note(IReadOnlyList<Cited> cites, IReadOnlySet<string> citedBefore)
    {
        var o = new Out();
        for (var i = 0; i < cites.Count; i++)
        {
            if (i > 0) o.Add("; ");
            var c = cites[i];
            o.Add(citedBefore.Contains(c.Source.Tag) ? ShortNote(c) : FullNote(c));
        }
        o.End('.');
        return o.Spans;
    }

    List<Span> FullNote(Cited c)
    {
        var s = c.Source;
        var o = new Out();
        if (s.Authors.Count > 0) o.Add(Names(s.Authors, NameOrder.Normal, "and", etAlFrom: 3, etAlKeep: 1)).Add(", ");
        var loc = c.Locator.Length > 0 ? Range(c.Locator) : "";
        switch (s.Type)
        {
            case "article":
                o.Quoted(TitleCase(s.Title), ',').Add(" ").Italic(TitleCase(s.Container));
                if (s.Volume.Length > 0) o.Add(" " + s.Volume);
                if (s.Issue.Length > 0) o.Add(", no. " + Range(s.Issue));
                o.Add(" (" + ChicagoMonthYear(s) + ")");
                if (loc.Length > 0) o.Add(": " + loc);
                break;
            case "magazine" or "newspaper" or "webpage":
                o.Quoted(TitleCase(s.Title), ',');
                if (s.Container.Length > 0) { o.Add(" "); if (s.Type == "webpage") o.Add(s.Container); else o.Italic(TitleCase(s.Container)); o.Add(","); }
                if (ChicagoDay(s) is { Length: > 0 } day) o.Add(" " + day);
                else if (DateOf(s.Accessed) is { } seen) o.Add(" accessed " + LongDate(seen));
                if (loc.Length > 0) o.Add(", " + loc);
                break;
            case "chapter":
                o.Quoted(TitleCase(s.Title), ',').Add(" in ").Italic(TitleCase(s.Container));
                if (s.Editors.Count > 0) o.Add(", ed. " + Names(s.Editors, NameOrder.Normal, "and", 3, 1));
                o.Add(" (" + Imprint(s, dated: false) + ")");
                if (loc.Length > 0) o.Add(", " + loc);
                break;
            default:
                o.Italic(TitleCase(s.Title));
                if (s.Edition.Length > 0) o.Add(", " + Ordinal(s.Edition) + " ed.");
                if (Imprint(s, dated: false) is { Length: > 0 } imprint) o.Add(" (" + imprint + ")");
                if (loc.Length > 0) o.Add(", " + loc);
                break;
        }
        if (s.Doi.Length > 0) o.Add(", " + DoiUrl(s.Doi));
        else if (s.Url.Length > 0) o.Add(", " + s.Url);
        return o.Spans;
    }

    List<Span> ShortNote(Cited c)
    {
        var s = c.Source;
        var o = new Out();
        var who = ShortNames(s, "and", etAlFrom: 3);
        if (who.Length > 0) o.Add(who + ", ");
        ShortTitle(o, s, comma: c.Locator.Length > 0);
        if (c.Locator.Length > 0) o.Add(" " + Range(c.Locator));
        return o.Spans;
    }

    // ----- names -----

    enum NameOrder { Normal, FirstInverted }

    /// <summary>Authors as a list names them: "Pegg, Ian L., and Mark Smith", with et al. after etAlKeep names when there are
    /// etAlFrom or more.</summary>
    static string Names(IReadOnlyList<Person> people, NameOrder order, string and, int etAlFrom, int etAlKeep)
    {
        string One(Person p, bool inverted) => p.Org || p.First.Length == 0 ? p.Last : inverted ? p.Last + ", " + p.First : p.First + " " + p.Last;
        if (people.Count >= etAlFrom)
        {
            var kept = people.Take(etAlKeep).Select((p, i) => One(p, i == 0 && order == NameOrder.FirstInverted)).ToList();
            return string.Join(", ", kept) + (kept.Count == 1 && order == NameOrder.Normal ? " et al." : ", et al.");
        }
        var names = people.Select((p, i) => One(p, i == 0 && order == NameOrder.FirstInverted)).ToList();
        return names.Count switch
        {
            1 => names[0],
            2 => names[0] + (order == NameOrder.FirstInverted ? ", " : " ") + and + " " + names[1],
            _ => string.Join(", ", names.Take(names.Count - 1)) + ", " + and + " " + names[^1],
        };
    }

    /// <summary>APA's author list: surname and initials, an ampersand before the last; up to 20 names, then the first 19, an
    /// ellipsis and the last.</summary>
    static string ApaNames(IReadOnlyList<Person> people)
    {
        string One(Person p) => p.Org || p.First.Length == 0 ? p.Last : p.Last + ", " + Initials(p.First);
        var names = people.Select(One).ToList();
        if (names.Count > 20) return string.Join(", ", names.Take(19)) + ", . . . " + names[^1];
        return names.Count switch
        {
            1 => names[0],
            2 => names[0] + ", & " + names[1],
            _ => string.Join(", ", names.Take(names.Count - 1)) + ", & " + names[^1],
        };
    }

    static string ApaEditors(IReadOnlyList<Person> editors)
    {
        var names = editors.Select(p => p.Org || p.First.Length == 0 ? p.Last : Initials(p.First) + " " + p.Last).ToList();
        var list = names.Count switch { 1 => names[0], 2 => names[0] + " & " + names[1], _ => string.Join(", ", names.Take(names.Count - 1)) + ", & " + names[^1] };
        return list + (names.Count > 1 ? " (Eds.)" : " (Ed.)");
    }

    /// <summary>The names in a citation in the text: a surname (with initials when another first author has that surname), two
    /// surnames joined by `and`, or the first surname and et al. from etAlFrom authors.</summary>
    string ShortNames(Source s, string and, int etAlFrom)
    {
        if (s.Authors.Count == 0) return "";
        string Surname(Person p) => !p.Org && p.First.Length > 0 && _sharedSurname.Contains(Fold(p.Last)) ? Initials(p.First) + " " + p.Last : p.Last;
        if (s.Authors.Count >= etAlFrom) return Surname(s.Authors[0]) + " et al.";
        if (s.Authors.Count == 2) return Surname(s.Authors[0]) + " " + and + " " + Surname(s.Authors[1]);
        return Surname(s.Authors[0]);
    }

    /// <summary>A short title in a citation: up to four words before any subtitle, in quotation marks for a part (an article, a
    /// chapter, a page), in italics for a whole (a book, a report).</summary>
    void ShortTitle(Out o, Source s, bool comma = false)
    {
        var title = Regex.Split(s.Title, @"[:?!]")[0].Trim();
        var words = title.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var skip = words.Length > 1 && Articles.Contains(words[0]) && _style is CiteStyle.Mla or CiteStyle.Chicago ? 1 : 0;
        var kept = words.Skip(skip).Take(4).ToList();
        while (kept.Count > 1 && Minor.Contains(kept[^1].Trim(',', ';'))) kept.RemoveAt(kept.Count - 1);
        var shortTitle = string.Join(" ", kept).TrimEnd(',', ';');
        if (_style is CiteStyle.Mla or CiteStyle.Chicago) shortTitle = TitleCase(shortTitle);
        if (Italic(s)) { o.Italic(shortTitle); if (comma) o.Add(","); }
        else if (_style == CiteStyle.Apa) o.Add("“" + shortTitle + (comma ? ",”" : "”"));
        else o.Quoted(shortTitle, comma ? ',' : '\0');
    }

    /// <summary>Initials from given names: "Ian Lee" and "I. L." are both "I. L."; "Jean-Paul" is "J.-P.".</summary>
    public static string Initials(string first) => string.Join(" ", first.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(part =>
        string.Join("-", part.Split('-').Where(x => x.Length > 0).Select(x => char.ToUpperInvariant(x.TrimEnd('.')[0]) + "."))));

    // ----- dates, pages, links -----

    static readonly string[] MonthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    static readonly string[] MlaMonths = ["Jan.", "Feb.", "Mar.", "Apr.", "May", "June", "July", "Aug.", "Sept.", "Oct.", "Nov.", "Dec."];

    static string YearOf(Source s, string none = "") => s.Year.Trim() is { Length: > 0 } y ? y : none;

    /// <summary>The month as 1–12 from a number or a name ("3", "March", "Mar."), 0 when there is none.</summary>
    static int MonthOf(string month)
    {
        var m = month.Trim().TrimEnd('.');
        if (int.TryParse(m, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) && n is >= 1 and <= 12) return n;
        if (m.Length < 3) return 0;
        for (var i = 0; i < 12; i++) if (MonthNames[i].StartsWith(m, StringComparison.OrdinalIgnoreCase)) return i + 1;
        return 0;
    }

    static (int Year, int Month, int Day)? DateOf(string iso) =>
        Regex.Match(iso, @"^(\d{4})-(\d{1,2})-(\d{1,2})") is { Success: true } m
            ? (int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture), int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture), int.Parse(m.Groups[3].Value, CultureInfo.InvariantCulture)) : null;

    static string MlaDay((int Year, int Month, int Day) d) => d.Day + " " + MlaMonths[d.Month - 1] + " " + d.Year;
    static string LongDate((int Year, int Month, int Day) d) => MonthNames[d.Month - 1] + " " + d.Day + ", " + d.Year;

    static string MlaDate(Source s)
    {
        var y = YearOf(s); if (y.Length == 0) return "";
        var m = MonthOf(s.Month);
        if (m == 0 || s.Type is "article" or "book" or "chapter" or "report" or "thesis") return y;
        return (s.Day.Length > 0 ? s.Day.Trim() + " " : "") + MlaMonths[m - 1] + " " + y;
    }

    string ApaYear(Source s) => YearOf(s, "n.d.") + (_yearSuffix.GetValueOrDefault(s.Tag) is { } suffix ? (s.Year.Length > 0 ? suffix : "-" + suffix) : "");

    /// <summary>APA's date: the year, and for dated sources (a web page, a magazine, a newspaper) its month and day.</summary>
    string ApaDate(Source s)
    {
        var y = ApaYear(s);
        var m = MonthOf(s.Month);
        if (m == 0 || s.Year.Length == 0 || s.Type is "article" or "book" or "chapter" or "report" or "thesis") return y;
        return y + ", " + MonthNames[m - 1] + (s.Day.Length > 0 ? " " + s.Day.Trim() : "");
    }

    static string ChicagoMonthYear(Source s)
    {
        var m = MonthOf(s.Month);
        return m > 0 && s.Year.Length > 0 && s.Type != "article" ? MonthNames[m - 1] + " " + s.Year : YearOf(s, "n.d.");
    }

    static string ChicagoDay(Source s)
    {
        var y = YearOf(s); if (y.Length == 0) return "";
        var m = MonthOf(s.Month);
        return m == 0 ? y : MonthNames[m - 1] + " " + (s.Day.Length > 0 ? s.Day.Trim() + ", " : "") + y;
    }

    static string ChicagoMonthDay(Source s)
    {
        var m = MonthOf(s.Month);
        return m == 0 ? "" : MonthNames[m - 1] + (s.Day.Length > 0 ? " " + s.Day.Trim() : "");
    }

    /// <summary>A page range with an en dash: 287-292 is 287–292.</summary>
    public static string Range(string pages) => Regex.Replace(pages.Trim(), @"(?<=\d)\s*[-‐‑–—]+\s*(?=\d)", "–");

    /// <summary>MLA's page range: the second number in full up to 99, else its last two digits when the rest is the same (287–92).</summary>
    public static string MlaRange(string pages) => Regex.Replace(Range(pages), @"(\d+)–(\d+)", m =>
    {
        var (a, b) = (m.Groups[1].Value, m.Groups[2].Value);
        if (b.Length < 3 || a.Length != b.Length || a[..^2] != b[..^2]) return a + "–" + b;
        return a + "–" + b[^2..];
    });

    /// <summary>APA's page in a citation: p. 287, pp. 287–289; a place given in words (para. 4, Chapter 2) as it is.</summary>
    static string Locator(string locator)
    {
        var l = Range(locator);
        if (!char.IsDigit(l[0]) && !Regex.IsMatch(l, @"^[ivxlcdm]+([–,]|$)", RegexOptions.IgnoreCase)) return l;
        return (l.Contains('–') || l.Contains(',') ? "pp. " : "p. ") + l;
    }

    static bool IsRange(string pages) => Regex.IsMatch(pages, @"\d\s*[-‐‑–—,]\s*\d");

    public static string DoiUrl(string doi) => "https://doi.org/" + Regex.Replace(doi.Trim(), @"^(https?://(dx\.)?doi\.org/|doi:\s*)", "", RegexOptions.IgnoreCase);

    static string Ordinal(string edition)
    {
        if (!int.TryParse(edition.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) return edition.Trim();
        var suffix = (n % 100) is 11 or 12 or 13 ? "th" : (n % 10) switch { 1 => "st", 2 => "nd", 3 => "rd", _ => "th" };
        return n + suffix;
    }

    /// <summary>Whether the title of a source is set in italics: a whole work (a book, a report, a thesis, a stand-alone web page in
    /// APA); a part of one (an article, a chapter, a page of a site) goes in quotation marks.</summary>
    bool Italic(Source s) => s.Type is "book" or "report" or "thesis" or "other" || s.Type == "webpage" && _style == CiteStyle.Apa;

    static bool Same(string a, string b) => string.Equals(Fold(a), Fold(b), StringComparison.Ordinal);

    // ----- titles -----

    static readonly HashSet<string> Minor = new(StringComparer.OrdinalIgnoreCase)
    {
        "a", "an", "the", "and", "but", "or", "nor", "for", "so", "yet", "as", "at", "by", "from", "in", "into", "of", "off", "on", "onto",
        "out", "over", "per", "than", "to", "up", "upon", "via", "with", "within", "without", "vs", "versus", "about", "after", "against",
        "along", "among", "around", "before", "behind", "below", "beneath", "beside", "between", "beyond", "down", "during", "except",
        "inside", "like", "near", "outside", "since", "through", "throughout", "toward", "towards", "under", "until", "unto",
    };
    static readonly HashSet<string> Articles = new(StringComparer.OrdinalIgnoreCase) { "a", "an", "the" };

    /// <summary>Title case as MLA and Chicago set titles: the first and last words and every major word capitalised; articles,
    /// prepositions and coordinating conjunctions lower case. A word already carrying capitals past its first letter or digits (DNA,
    /// iPhone, Tc(IV), 99mTc) is left as it is.</summary>
    public static string TitleCase(string title)
    {
        var words = Regex.Split(title, @"(\s+)");
        var real = words.Select((w, i) => (w, i)).Where(x => !string.IsNullOrWhiteSpace(x.w)).Select(x => x.i).ToList();
        var afterBreak = true;
        var sb = new StringBuilder();
        for (var i = 0; i < words.Length; i++)
        {
            var w = words[i];
            if (string.IsNullOrWhiteSpace(w)) { sb.Append(w); continue; }
            var last = i == real[^1];
            var parts = w.Split('-');
            sb.Append(string.Join("-", parts.Select((part, j) => Word(part, j == 0 && (afterBreak || last || parts.Length > 1), j > 0))));
            afterBreak = Regex.IsMatch(w, @"[:?!—]$");
        }
        return sb.ToString();

        static string Word(string part, bool force, bool inCompound)
        {
            var m = Regex.Match(part, @"^([^\p{L}\p{N}]*)([\p{L}\p{N}’']+)(.*)$");
            if (!m.Success) return part;
            var (lead, core, tail) = (m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value);
            if (core.Skip(1).Any(char.IsUpper) || core.Any(char.IsDigit) || tail.Any(char.IsUpper)) return part;
            if (!force && Minor.Contains(core) && !(inCompound && core.Length > 3)) return lead + core.ToLowerInvariant() + tail;
            return lead + char.ToUpperInvariant(core[0]) + core[1..] + tail;
        }
    }

    /// <summary>A title for sorting: folded to plain lower-case letters, without a leading article.</summary>
    static string SortTitle(string title) => Regex.Replace(Fold(title), @"^(a|an|the)\s+", "");

    /// <summary>The key of a source's authors: two sources by the same people share it.</summary>
    static string AuthorKey(Source s) => string.Join("|", s.Authors.Select(a => Fold(a.Last) + "," + Fold(a.First)));

    /// <summary>Lower case without accents or punctuation, for comparing and sorting names and titles (the engine runs without culture data).</summary>
    public static string Fold(string text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var ch in text.ToLowerInvariant())
        {
            var i = Accented.IndexOf(ch);
            if (i >= 0) sb.Append(Plain[i]);
            else if (ch == 'ß') sb.Append("ss");
            else if (ch == 'æ') sb.Append("ae");
            else if (ch == 'œ') sb.Append("oe");
            else if (char.IsLetterOrDigit(ch) || ch == ' ') sb.Append(ch);
        }
        return sb.ToString().Trim();
    }

    const string Accented = "àáâãäåāăąçćčďđèéêëēėęěìíîïīįıñńňòóôõöøōőŕřśšşťùúûüūůűųýÿžźżł";
    const string Plain = "aaaaaaaaacccddeeeeeeeeiiiiiiinnnoooooooorrssstuuuuuuuuyyzzzl";
}

/// <summary>Builds formatted text: runs of the same slant join, and punctuation that would double (a title that ends in a question
/// mark takes no period) is dropped. Periods and commas go inside closing quotation marks, as American style has them.</summary>
sealed class Out
{
    readonly List<Span> _spans = [];

    public List<Span> Spans => _spans;
    public int Length => _spans.Sum(s => s.Text.Length);

    public Out Add(string text, bool italic = false)
    {
        if (text.Length == 0) return this;
        if (_spans.Count > 0 && _spans[^1].Italic == italic) _spans[^1] = new(_spans[^1].Text + text, italic);
        else _spans.Add(new(text, italic));
        return this;
    }

    public Out Add(IEnumerable<Span> spans)
    {
        foreach (var s in spans) Add(s.Text, s.Italic);
        return this;
    }

    public Out Italic(string text) => Add(text, true);

    /// <summary>A title in quotation marks followed by punctuation that goes inside them: “Title,” — none after a title that ends in
    /// its own (“Why?”); '\0' for none.</summary>
    public Out Quoted(string title, char punctuation)
    {
        var t = title.TrimEnd();
        var own = t.Length > 0 && t[^1] is '?' or '!';
        return Add("“" + t + (punctuation != '\0' && !own ? punctuation.ToString() : "") + "”");
    }

    /// <summary>Ends the text with a period or comma unless it already ends with a sentence's end (a period, a question mark, a
    /// closing quotation mark with punctuation inside).</summary>
    public Out End(char punctuation)
    {
        if (_spans.Count == 0) return this;
        var text = _spans[^1].Text.TrimEnd();
        if (text.Length == 0) return this;
        var last = text[^1];
        if (last == '”' && text.Length > 1 && text[^2] is '.' or ',' or '?' or '!')
        {
            if (punctuation == '.' && text[^2] == ',') _spans[^1] = new(text[..^2] + ".”", _spans[^1].Italic);
            return this;
        }
        if (last is '.' or '?' or '!') return this;
        if (last == ',' && punctuation == '.') { _spans[^1] = new(text[..^1] + ".", _spans[^1].Italic); return this; }
        // punctuation after italic text is roman, as the styles set it
        return Add(punctuation.ToString());
    }
}
