using Writer.Formats.Cite;

namespace Writer.Tests;

/// <summary>The citation formatter against the style guides' own examples: MLA 9, APA 7, Chicago 18 in notes and bibliography and
/// in author-date. Italic text shows between asterisks.</summary>
public class CitationsTests
{
    static string Show(IEnumerable<Span> spans) => string.Concat(spans.Select(s => s.Italic ? "*" + s.Text + "*" : s.Text));

    static readonly Source Pegg = new()
    {
        Tag = "Peg15", Type = "article", Authors = [new("Pegg", "Ian L.")], Title = "Behavior of technetium in nuclear waste vitrification processes",
        Container = "Journal of Radioanalytical and Nuclear Chemistry", Year = "2015", Volume = "305", Issue = "1", Pages = "287-292", Doi = "10.1007/s10967-014-3900-9",
    };
    static readonly Source Soderquist = new()
    {
        Tag = "Sod14", Type = "article", Authors = [new("Soderquist", "C. Z."), new("Schweiger", "M. J."), new("Kim", "D.-S."), new("Lukens", "W. W."), new("McCloy", "J. S.")],
        Title = "Redox-dependent solubility of technetium in low activity waste glass", Container = "Journal of Nuclear Materials", Year = "2014", Volume = "449", Issue = "1-3", Pages = "173-180",
        Doi = "https://doi.org/10.1016/j.jnucmat.2014.03.008",
    };
    static readonly Source Kuhn = new()
    {
        Tag = "Kuh62", Type = "book", Authors = [new("Kuhn", "Thomas S.")], Title = "The structure of scientific revolutions", Publisher = "University of Chicago Press", Place = "Chicago", Year = "1962",
    };
    static readonly Source Page = new()
    {
        Tag = "How21", Type = "webpage", Title = "How to cite a website in MLA", Container = "Scribbr", Year = "2021", Month = "3", Day = "5", Url = "https://www.scribbr.com/mla/website-citation/",
    };

    [Fact]
    public void Mla_lists_works_cited_in_title_case_with_the_containers_elements_the_doi_last_and_cites_author_and_page()
    {
        var mla = new Citations(CiteStyle.Mla, [Pegg, Soderquist, Kuhn, Page]);
        Assert.Equal("Pegg, Ian L. “Behavior of Technetium in Nuclear Waste Vitrification Processes.” *Journal of Radioanalytical and Nuclear Chemistry*, vol. 305, no. 1, 2015, pp. 287–92, https://doi.org/10.1007/s10967-014-3900-9.", Show(mla.Entry(Pegg)));
        Assert.Equal("Soderquist, C. Z., et al. “Redox-Dependent Solubility of Technetium in Low Activity Waste Glass.” *Journal of Nuclear Materials*, vol. 449, no. 1–3, 2014, pp. 173–80, https://doi.org/10.1016/j.jnucmat.2014.03.008.", Show(mla.Entry(Soderquist)));
        Assert.Equal("Kuhn, Thomas S. *The Structure of Scientific Revolutions*. University of Chicago Press, 1962.", Show(mla.Entry(Kuhn)));
        Assert.Equal("“How to Cite a Website in MLA.” *Scribbr*, 5 Mar. 2021, www.scribbr.com/mla/website-citation/.", Show(mla.Entry(Page)));
        Assert.Equal("(Pegg 288)", Show(mla.InText([new(Pegg, "288")])));
        Assert.Equal("(Soderquist et al. 175–79)", Show(mla.InText([new(Soderquist, "175-179")])));
        Assert.Equal("(288)", Show(mla.InText([new(Pegg, "288")], noAuthor: true)));
        Assert.Equal("(Pegg 288; Soderquist et al. 175)", Show(mla.InText([new(Pegg, "288"), new(Soderquist, "175")])));
        Assert.Equal("(“How to Cite”)", Show(mla.InText([new(Page)])));
        Assert.Equal(["How21", "Kuh62", "Peg15", "Sod14"], mla.Sort([Soderquist, Pegg, Page, Kuhn]).Select(s => s.Tag));
        Assert.Equal("Works Cited", Citations.ListTitle(CiteStyle.Mla));
    }

    [Fact]
    public void Apa_lists_references_with_initials_and_the_year_and_cites_author_year_and_page()
    {
        var apa = new Citations(CiteStyle.Apa, [Pegg, Soderquist, Kuhn, Page]);
        Assert.Equal("Pegg, I. L. (2015). Behavior of technetium in nuclear waste vitrification processes. *Journal of Radioanalytical and Nuclear Chemistry*, *305*(1), 287–292. https://doi.org/10.1007/s10967-014-3900-9", Show(apa.Entry(Pegg)));
        Assert.Equal("Soderquist, C. Z., Schweiger, M. J., Kim, D.-S., Lukens, W. W., & McCloy, J. S. (2014). Redox-dependent solubility of technetium in low activity waste glass. *Journal of Nuclear Materials*, *449*(1–3), 173–180. https://doi.org/10.1016/j.jnucmat.2014.03.008", Show(apa.Entry(Soderquist)));
        Assert.Equal("Kuhn, T. S. (1962). *The structure of scientific revolutions*. University of Chicago Press.", Show(apa.Entry(Kuhn)));
        Assert.Equal("*How to cite a website in MLA*. (2021, March 5). Scribbr. https://www.scribbr.com/mla/website-citation/", Show(apa.Entry(Page)));
        Assert.Equal("(Pegg, 2015, p. 288)", Show(apa.InText([new(Pegg, "288")])));
        Assert.Equal("(Soderquist et al., 2014, pp. 175–179)", Show(apa.InText([new(Soderquist, "175-179")])));
        Assert.Equal("(2015, p. 288)", Show(apa.InText([new(Pegg, "288")], noAuthor: true)));
        Assert.Equal("(Pegg, 2015; Soderquist et al., 2014)", Show(apa.InText([new(Soderquist), new(Pegg)])));
        Assert.Equal("References", Citations.ListTitle(CiteStyle.Apa));
    }

    [Fact]
    public void Apa_tells_one_authors_works_of_one_year_apart_by_letters_in_title_order()
    {
        var later = Pegg with { Tag = "Peg15b", Title = "A later look at technetium" };
        var apa = new Citations(CiteStyle.Apa, [Pegg, later]);
        Assert.Equal("(Pegg, 2015a)", Show(apa.InText([new(Pegg)]))); // by title, A left out: Behavior before Later
        Assert.Equal("(Pegg, 2015b)", Show(apa.InText([new(later)])));
        Assert.StartsWith("Pegg, I. L. (2015b). A later look", Show(apa.Entry(later)));
        Assert.Equal(["Peg15", "Peg15b"], apa.Sort([later, Pegg]).Select(s => s.Tag));
        var mla = new Citations(CiteStyle.Mla, [Pegg, later]);
        Assert.Equal("(Pegg, “Behavior of Technetium” 288)", Show(mla.InText([new(Pegg, "288")]))); // MLA names the work when the author has two
    }

    [Fact]
    public void Chicago_cites_in_notes_full_the_first_time_and_short_after_and_lists_a_bibliography()
    {
        var chicago = new Citations(CiteStyle.Chicago, [Pegg, Kuhn]);
        Assert.Equal("Ian L. Pegg, “Behavior of Technetium in Nuclear Waste Vitrification Processes,” *Journal of Radioanalytical and Nuclear Chemistry* 305, no. 1 (2015): 288, https://doi.org/10.1007/s10967-014-3900-9.",
            Show(chicago.Note([new(Pegg, "288")], new HashSet<string>())));
        Assert.Equal("Pegg, “Behavior of Technetium,” 290.", Show(chicago.Note([new(Pegg, "290")], new HashSet<string> { "Peg15" })));
        Assert.Equal("Pegg, “Behavior of Technetium.”", Show(chicago.Note([new(Pegg)], new HashSet<string> { "Peg15" }))); // the period inside the quotation marks
        Assert.Equal("Thomas S. Kuhn, *The Structure of Scientific Revolutions* (Chicago: University of Chicago Press, 1962), 45.", Show(chicago.Note([new(Kuhn, "45")], new HashSet<string>())));
        Assert.Equal("Kuhn, *Structure of Scientific Revolutions*, 45.", Show(chicago.Note([new(Kuhn, "45")], new HashSet<string> { "Kuh62" })));
        Assert.Equal("Pegg, Ian L. “Behavior of Technetium in Nuclear Waste Vitrification Processes.” *Journal of Radioanalytical and Nuclear Chemistry* 305, no. 1 (2015): 287–292. https://doi.org/10.1007/s10967-014-3900-9.", Show(chicago.Entry(Pegg)));
        Assert.Equal("Kuhn, Thomas S. *The Structure of Scientific Revolutions*. Chicago: University of Chicago Press, 1962.", Show(chicago.Entry(Kuhn)));
        Assert.True(Citations.InNotes(CiteStyle.Chicago));
        Assert.Equal("Bibliography", Citations.ListTitle(CiteStyle.Chicago));
    }

    [Fact]
    public void Chicago_author_date_puts_the_year_after_the_author_and_cites_author_year_page()
    {
        var dated = new Citations(CiteStyle.ChicagoDate, [Pegg, Page]);
        Assert.Equal("Pegg, Ian L. 2015. “Behavior of Technetium in Nuclear Waste Vitrification Processes.” *Journal of Radioanalytical and Nuclear Chemistry* 305 (1): 287–292. https://doi.org/10.1007/s10967-014-3900-9.", Show(dated.Entry(Pegg)));
        Assert.Equal("“How to Cite a Website in MLA.” 2021. Scribbr. March 5. https://www.scribbr.com/mla/website-citation/.", Show(dated.Entry(Page)));
        Assert.Equal("(Pegg 2015, 288)", Show(dated.InText([new(Pegg, "288")])));
    }

    [Theory]
    [InlineData("the art of war: a new translation", "The Art of War: A New Translation")]
    [InlineData("DNA repair in iPhone users", "DNA Repair in iPhone Users")]
    [InlineData("redox-dependent solubility of Tc(IV) and 99mTc", "Redox-Dependent Solubility of Tc(IV) and 99mTc")]
    [InlineData("what is it for?", "What Is It For?")]
    public void Title_case_capitalises_major_words_and_leaves_words_with_their_own_capitals(string title, string expected) =>
        Assert.Equal(expected, Citations.TitleCase(title));

    [Theory]
    [InlineData("287-292", "287–92")]
    [InlineData("101-108", "101–08")]
    [InlineData("498-532", "498–532")]
    [InlineData("22-27", "22–27")]
    [InlineData("1087-1089", "1087–89")]
    public void Mla_shortens_the_second_number_of_a_range(string pages, string expected) => Assert.Equal(expected, Citations.MlaRange(pages));

    [Theory]
    [InlineData("Ian L.", "I. L.")]
    [InlineData("Jean-Paul", "J.-P.")]
    [InlineData("D.-S.", "D.-S.")]
    public void Initials_come_from_given_names(string first, string expected) => Assert.Equal(expected, Citations.Initials(first));
}
