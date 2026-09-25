using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

public sealed class PptxAdapter : IFormatAdapter
{
    public string Format => "pptx";
    public IReadOnlyList<string> Extensions { get; } = [".pptx"];
    public bool CanWrite => true;

    public Document Create()
    {
        var ms = new MemoryStream();
        PptxTemplate.WriteBlank(ms);
        ms.Position = 0;
        return Open(ms);
    }

    public Document Open(Stream stream)
    {
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        ms.Position = 0;
        PresentationDocument package;
        try
        {
            package = PresentationDocument.Open(ms, true);
        }
        catch (Exception ex) when (ex is not WriterException)
        {
            throw new WriterException(ErrorCode.FormatError, $"Not a valid .pptx file: {ex.Message}", "Check that the file opens in PowerPoint.");
        }
        if (package.PresentationPart?.Presentation is null)
        {
            package.Dispose();
            throw new WriterException(ErrorCode.FormatError, "The .pptx has no presentation part", "Check that the file opens in PowerPoint.");
        }
        return new PptxDocument(package);
    }
}

/// <summary>An open PowerPoint package. Slide order comes from the presentation's slide id list.</summary>
public sealed class PptxDocument(PresentationDocument package) : Document
{
    public PresentationDocument Package { get; } = package;
    public PresentationPart Presentation => Package.PresentationPart!;
    internal IEnumerable<KeyValuePair<string, string>> Namespaces => Presentation.Presentation!.NamespaceDeclarations;

    public override string Format => "pptx";
    public override Node Root => new PptxRoot(this);

    public override void Save(Stream stream)
    {
        foreach (var slide in Slides) PptxAnim.Prune(slide);
        PptxSections.Normalize(this);
        using var clone = Package.Clone(stream);
    }

    public override void Dispose() => Package.Dispose();

    P.SlideIdList SlideIds => Presentation.Presentation!.SlideIdList ??= new P.SlideIdList();

    internal List<SlidePart> Slides =>
        SlideIds.Elements<P.SlideId>()
            .Select(id => id.RelationshipId?.Value is { } rel && Presentation.TryGetPartById(rel, out var part) ? part as SlidePart : null)
            .Where(p => p is not null).Select(p => p!).ToList();

    internal (long Width, long Height) SlideSize =>
        Presentation.Presentation!.SlideSize is { Cx.Value: var cx, Cy.Value: var cy } ? (cx, cy) : (PptxTemplate.SlideWidth, PptxTemplate.SlideHeight);

    internal P.SlideId SlideIdOf(SlidePart slide)
    {
        var rel = Presentation.GetIdOfPart(slide);
        return SlideIds.Elements<P.SlideId>().First(s => s.RelationshipId?.Value == rel);
    }

    internal IEnumerable<SlideLayoutPart> Layouts => Presentation.SlideMasterParts.SelectMany(m => m.SlideLayoutParts);

    /// <summary>A layout by name (case-insensitive), else the standard layout a name stands for (a key such as two or quote,
    /// PowerPoint's English or Chinese name, see <see cref="PptxTemplate.Layouts"/>): the deck's own of that type when it has one,
    /// otherwise made in its first master from the standard spec. The deck's content layout when none is asked for.</summary>
    internal SlideLayoutPart FindLayout(string? name)
    {
        var layouts = Layouts.ToList();
        if (layouts.Count == 0) throw new WriterException(ErrorCode.FormatError, "The presentation has no slide layouts", "Create the deck with 'writer create'.");
        if (name is null) return layouts.FirstOrDefault(l => LayoutType(l) == "obj") ?? layouts[0];
        var byName = layouts.FirstOrDefault(l => string.Equals(LayoutName(l), name, StringComparison.OrdinalIgnoreCase));
        if (byName is not null) return byName;
        var spec = PptxTemplate.Standard(name) ?? throw new WriterException(ErrorCode.Validation, $"No layout called '{name}'",
            $"Layouts in this file: {string.Join(", ", layouts.Select(LayoutName))}; standard ones (made when missing): {string.Join(", ", PptxTemplate.Layouts.Select(l => l.Key))}.");
        return layouts.FirstOrDefault(l => Is(l, spec)) ?? AddLayout(spec);
    }

    /// <summary>A layout of the deck is the standard one when its type says so (a quote layout, which has no type of its own, by name).</summary>
    static bool Is(SlideLayoutPart layout, PptxTemplate.LayoutSpec spec) => spec.Type == "cust"
        ? LayoutName(layout).Contains("quote", StringComparison.OrdinalIgnoreCase) || LayoutName(layout).Contains(spec.Chinese, StringComparison.Ordinal)
        : LayoutType(layout) == spec.Type;

    /// <summary>A standard layout added to the first master, sized to the deck, under a layout id no master or layout has.</summary>
    SlideLayoutPart AddLayout(PptxTemplate.LayoutSpec spec)
    {
        var master = Presentation.SlideMasterParts.First();
        var (width, height) = SlideSize;
        var layout = master.AddNewPart<SlideLayoutPart>();
        layout.SlideLayout = new P.SlideLayout(PptxTemplate.LayoutXml(spec, width, height));
        layout.AddPart(master);
        var used = (Presentation.Presentation!.SlideMasterIdList?.Elements<P.SlideMasterId>().Select(m => m.Id?.Value ?? 0u) ?? [])
            .Concat(Presentation.SlideMasterParts.SelectMany(m => m.SlideMaster?.SlideLayoutIdList?.Elements<P.SlideLayoutId>().Select(l => l.Id?.Value ?? 0u) ?? []));
        var list = master.SlideMaster!.SlideLayoutIdList ??= new P.SlideLayoutIdList();
        list.Append(new P.SlideLayoutId { Id = Math.Max(2147483648u, used.DefaultIfEmpty(0u).Max() + 1), RelationshipId = master.GetIdOfPart(layout) });
        return layout;
    }

    internal static string LayoutName(SlideLayoutPart layout) =>
        layout.SlideLayout?.CommonSlideData?.Name?.Value ?? layout.SlideLayout?.Type?.InnerText ?? "Layout";

    internal static string? LayoutType(SlideLayoutPart layout) => layout.SlideLayout?.Type?.InnerText;

    internal SlidePart AddSlide(string? layoutName, int? index)
    {
        var layout = FindLayout(layoutName);
        var slide = Presentation.AddNewPart<SlidePart>();
        slide.Slide = PptxTemplate.NewSlide(layout);
        slide.AddPart(layout);
        List(slide, index);
        return slide;
    }

    /// <summary>A copy of slide at a 1-based position (null = last), made by <see cref="PptxCopy.Slide"/>.</summary>
    internal SlidePart CopySlide(SlidePart slide, int? index)
    {
        var copy = PptxCopy.Slide(Presentation, slide);
        List(copy, index);
        return copy;
    }

    /// <summary>Puts a new slide part in the slide list at a 1-based position (null = last), under a slide id no other slide has.</summary>
    void List(SlidePart slide, int? index)
    {
        var ids = SlideIds.Elements<P.SlideId>().ToList();
        var newId = new P.SlideId { Id = Math.Max(256u, ids.Count == 0 ? 256u : ids.Max(s => s.Id?.Value ?? 0u) + 1), RelationshipId = Presentation.GetIdOfPart(slide) };
        if (index is { } i && i >= 1 && i <= ids.Count) SlideIds.InsertBefore(newId, ids[i - 1]);
        else SlideIds.Append(newId);
    }

    internal void RemoveSlide(SlidePart slide)
    {
        SlideIdOf(slide).Remove();
        Presentation.DeletePart(slide);
    }

    internal void MoveSlide(SlidePart slide, int? index)
    {
        var id = SlideIdOf(slide);
        id.Remove();
        var ids = SlideIds.Elements<P.SlideId>().ToList();
        if (index is { } i && i >= 1 && i <= ids.Count) SlideIds.InsertBefore(id, ids[i - 1]);
        else SlideIds.Append(id);
    }

    /// <summary>The slide's notes page, created the way PowerPoint does (with a notes master and its own theme copy when the deck has none).</summary>
    internal NotesSlidePart EnsureNotes(SlidePart slide)
    {
        if (slide.NotesSlidePart is { } existing) return existing;
        var master = Presentation.NotesMasterPart ?? AddNotesMaster();
        var notes = slide.AddNewPart<NotesSlidePart>();
        notes.NotesSlide = PptxTemplate.NewNotesSlide();
        notes.AddPart(master);
        notes.AddPart(slide);
        return notes;
    }

    NotesMasterPart AddNotesMaster()
    {
        var master = Presentation.AddNewPart<NotesMasterPart>();
        master.NotesMaster = new P.NotesMaster(PptxTemplate.NotesMasterXml);
        if ((Presentation.SlideMasterParts.FirstOrDefault()?.ThemePart ?? Presentation.ThemePart)?.Theme is { } theme)
            master.AddNewPart<ThemePart>().Theme = (DocumentFormat.OpenXml.Drawing.Theme)theme.CloneNode(true);
        var list = new P.NotesMasterIdList(new P.NotesMasterId { Id = Presentation.GetIdOfPart(master) });
        var presentation = Presentation.Presentation!;
        if (presentation.SlideMasterIdList is { } after) presentation.InsertAfter(list, after);
        else presentation.PrependChild(list);
        return master;
    }

    internal static P.Shape? NotesBody(NotesSlidePart notes) =>
        notes.NotesSlide?.CommonSlideData?.ShapeTree?.Elements<P.Shape>().FirstOrDefault(s =>
            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape is { } ph && (ph.Type?.InnerText ?? "body") == "body");

    internal static uint NextShapeId(SlidePart slide)
    {
        var ids = slide.Slide?.Descendants<P.NonVisualDrawingProperties>().Select(p => p.Id?.Value ?? 0u).DefaultIfEmpty(1u).Max() ?? 1u;
        return Math.Max(ids, 1u) + 1;
    }

    internal static OpenXmlElement ShapeTree(SlidePart slide) => slide.Slide!.CommonSlideData!.ShapeTree!;
}
