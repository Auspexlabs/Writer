using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>
/// Copies within a deck, the markup deep-cloned with what it refers to. A slide keeps its layout, pictures and media (pointed at,
/// as PowerPoint shares them) and gets notes, charts, diagrams and embedded objects of its own; comments stay with the original. A
/// shape, picture or table copied onto another slide takes its relationships along; every copied element gets drawing ids that
/// no other shape on its slide has.
/// </summary>
static class PptxCopy
{
    const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string A16Ns = "http://schemas.microsoft.com/office/drawing/2014/main";
    /// <summary>Picture.cs keeps a picture's original for reset under this prefix and the id of the relationship it shows.</summary>
    const string KeptPrefix = "wrOrig";

    /// <summary>Parts a copy points at rather than copies: pictures, slides a link jumps to, and what the whole deck shares
    /// (layouts, masters, themes).</summary>
    static bool Shared(OpenXmlPart part) =>
        part is ImagePart or SlidePart or SlideLayoutPart or SlideMasterPart or NotesMasterPart or HandoutMasterPart or ThemePart or PresentationPart;

    /// <summary>A new slide part (not listed yet) holding a copy of slide under the same relationship ids.</summary>
    public static SlidePart Slide(PresentationPart presentation, SlidePart slide)
    {
        var copy = presentation.AddNewPart<SlidePart>();
        Relationships(slide, copy, slide, copy);
        copy.Slide = (P.Slide)slide.Slide!.CloneNode(true);
        return copy;
    }

    /// <summary>element (on slide from) ready to go onto slide to: deep-cloned, with drawing ids and creation ids of its own and,
    /// on another slide, the relationships its markup names (r:embed, r:link, r:id …) made there.</summary>
    public static T Element<T>(T element, SlidePart from, SlidePart to) where T : OpenXmlElement
    {
        var copy = (T)element.CloneNode(true);
        Ids(copy, to, keep: false);
        foreach (var creation in copy.Descendants().Where(e => e.LocalName == "creationId" && e.NamespaceUri == A16Ns))
            creation.SetAttribute(new OpenXmlAttribute("", "id", "", "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}"));
        if (!ReferenceEquals(from, to)) Relink(copy, from, to);
        return copy;
    }

    /// <summary>Numbers the drawing ids (p:cNvPr) in element so that none repeats one on the slide; keep leaves the free ones as they are.</summary>
    public static void Ids(OpenXmlElement element, SlidePart slide, bool keep)
    {
        var used = slide.Slide!.Descendants<P.NonVisualDrawingProperties>().Select(p => p.Id?.Value ?? 0u).ToHashSet();
        var next = Math.Max(used.DefaultIfEmpty(1u).Max(), 1u) + 1;
        foreach (var nv in element.Descendants<P.NonVisualDrawingProperties>())
        {
            if (keep && nv.Id?.Value is { } id && used.Add(id)) continue;
            while (!used.Add(next)) next++;
            nv.Id = next;
        }
    }

    static void Relink(OpenXmlElement element, SlidePart from, SlidePart to)
    {
        var ids = new Dictionary<string, string>();
        foreach (var e in element.Descendants().Prepend(element).ToList())
            foreach (var a in e.GetAttributes().Where(a => a.NamespaceUri == RelNs && !string.IsNullOrEmpty(a.Value)))
            {
                if (!ids.TryGetValue(a.Value!, out var id)) ids[a.Value!] = id = Carry(a.Value!, from, to);
                if (id != a.Value) e.SetAttribute(new OpenXmlAttribute(a.Prefix, a.LocalName, a.NamespaceUri, id));
            }
    }

    /// <summary>The id on slide to of a relationship to what relationship id names on slide from: the same picture (with the original
    /// kept for its reset), link or media file; a part of from's own, such as an embedded object, copied.</summary>
    static string Carry(string id, SlidePart from, SlidePart to)
    {
        if (from.TryGetPartById(id, out var part))
        {
            var carried = Shared(part) ? IdOf(to, part) ?? to.CreateRelationshipToPart(part) : IdOf(to, Part(part, to, null, from, to))!;
            if (from.TryGetPartById(KeptPrefix + id, out var kept) && !to.TryGetPartById(KeptPrefix + carried, out _))
            {
                if (IdOf(to, kept) is null) to.CreateRelationshipToPart(kept, KeptPrefix + carried);
                else Bytes(kept, to.AddImagePart(kept.ContentType, KeptPrefix + carried)); // a part takes one relationship per target
            }
            return carried;
        }
        if (from.HyperlinkRelationships.FirstOrDefault(r => r.Id == id) is { } link) return to.AddHyperlinkRelationship(link.Uri, link.IsExternal).Id;
        if (from.ExternalRelationships.FirstOrDefault(r => r.Id == id) is { } external) return to.AddExternalRelationship(external.RelationshipType, external.Uri).Id;
        if (from.DataPartReferenceRelationships.FirstOrDefault(r => r.Id == id) is { } media) return Reference(to, media, null)?.Id ?? id;
        return id;
    }

    static string? IdOf(OpenXmlPartContainer owner, OpenXmlPart part)
    {
        foreach (var pair in owner.Parts) if (ReferenceEquals(pair.OpenXmlPart, part)) return pair.RelationshipId;
        return null;
    }

    /// <summary>Gives copy the relationships source has, under the same ids: shared parts pointed at, the others copied; a notes page
    /// points back at the slide copy. Comments stay with the original.</summary>
    static void Relationships(OpenXmlPartContainer source, OpenXmlPartContainer copy, SlidePart slide, SlidePart slideCopy)
    {
        foreach (var pair in source.Parts)
        {
            var part = pair.OpenXmlPart;
            if (ReferenceEquals(part, slide)) copy.CreateRelationshipToPart(slideCopy, pair.RelationshipId);
            else if (part is SlideCommentsPart or PowerPointCommentPart) continue;
            else if (Shared(part)) copy.CreateRelationshipToPart(part, pair.RelationshipId);
            else Part(part, copy, pair.RelationshipId, slide, slideCopy);
        }
        foreach (var link in source.HyperlinkRelationships) copy.AddHyperlinkRelationship(link.Uri, link.IsExternal, link.Id);
        foreach (var external in source.ExternalRelationships) copy.AddExternalRelationship(external.RelationshipType, external.Uri, external.Id);
        foreach (var media in source.DataPartReferenceRelationships) Reference(copy, media, media.Id);
    }

    /// <summary>A new part of like's kind under owner, typed for the parts a slide can own (notes, charts with their styles,
    /// drawings and workbooks, diagrams, embedded objects, tags, VML); anything else goes in as an extended part with the same
    /// relationship and content type, which is the same thing on disk. No reflection: the release engine is NativeAOT.</summary>
    static OpenXmlPart NewPart(OpenXmlPartContainer owner, OpenXmlPart like, string id) => like switch
    {
        NotesSlidePart => owner.AddNewPart<NotesSlidePart>(like.ContentType, id),
        ChartPart => owner.AddNewPart<ChartPart>(like.ContentType, id),
        ExtendedChartPart => owner.AddNewPart<ExtendedChartPart>(like.ContentType, id),
        ChartStylePart => owner.AddNewPart<ChartStylePart>(like.ContentType, id),
        ChartColorStylePart => owner.AddNewPart<ChartColorStylePart>(like.ContentType, id),
        ChartDrawingPart => owner.AddNewPart<ChartDrawingPart>(like.ContentType, id),
        EmbeddedPackagePart => owner.AddNewPart<EmbeddedPackagePart>(like.ContentType, id),
        EmbeddedObjectPart => owner.AddNewPart<EmbeddedObjectPart>(like.ContentType, id),
        DiagramDataPart => owner.AddNewPart<DiagramDataPart>(like.ContentType, id),
        DiagramLayoutDefinitionPart => owner.AddNewPart<DiagramLayoutDefinitionPart>(like.ContentType, id),
        DiagramStylePart => owner.AddNewPart<DiagramStylePart>(like.ContentType, id),
        DiagramColorsPart => owner.AddNewPart<DiagramColorsPart>(like.ContentType, id),
        DiagramPersistLayoutPart => owner.AddNewPart<DiagramPersistLayoutPart>(like.ContentType, id),
        UserDefinedTagsPart => owner.AddNewPart<UserDefinedTagsPart>(like.ContentType, id),
        VmlDrawingPart => owner.AddNewPart<VmlDrawingPart>(like.ContentType, id),
        _ => owner.AddExtendedPart(like.RelationshipType, like.ContentType, Path.GetExtension(like.Uri.OriginalString).TrimStart('.'), id),
    };

    /// <summary>A copy of part under owner (id null: a new relationship id): its content as the document has it now, and its own
    /// relationships as <see cref="Relationships"/> gives them.</summary>
    static OpenXmlPart Part(OpenXmlPart part, OpenXmlPartContainer owner, string? id, SlidePart slide, SlidePart slideCopy)
    {
        id ??= "R" + Guid.NewGuid().ToString("N")[..16];
        var copy = NewPart(owner, part, id);
        (part.RootElement as OpenXmlPartRootElement)?.Save(); // markup changed in memory reaches the part's stream first
        Bytes(part, copy);
        Relationships(part, copy, slide, slideCopy);
        return copy;
    }

    static void Bytes(OpenXmlPart from, OpenXmlPart to)
    {
        using var stream = from.GetStream(FileMode.Open, FileAccess.Read);
        to.FeedData(stream);
    }

    /// <summary>A media reference (audio, video, or PowerPoint's media) from owner to the same media file.</summary>
    static DataPartReferenceRelationship? Reference(OpenXmlPartContainer owner, DataPartReferenceRelationship media, string? id)
    {
        if (owner is not SlidePart slide || media.DataPart is not MediaDataPart data) return null;
        return media switch
        {
            VideoReferenceRelationship => id is null ? slide.AddVideoReferenceRelationship(data) : slide.AddVideoReferenceRelationship(data, id),
            AudioReferenceRelationship => id is null ? slide.AddAudioReferenceRelationship(data) : slide.AddAudioReferenceRelationship(data, id),
            MediaReferenceRelationship => id is null ? slide.AddMediaReferenceRelationship(data) : slide.AddMediaReferenceRelationship(data, id),
            _ => null,
        };
    }
}
