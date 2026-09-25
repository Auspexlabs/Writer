using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace Writer.Formats.Pptx;

/// <summary>Sections (节): PowerPoint 2010's p14:sectionLst in the presentation's extLst. A section runs from its first slide to the next
/// section's first, so the model is the slide each starts at; the slide prop section names the one a slide starts. The save lays
/// the list out again from those starts, so a slide added, copied, moved or removed by any command lands in the section before it.</summary>
static class PptxSections
{
    const string Uri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";

    sealed record Sec(string Name, string Id, uint? Start);

    static P.PresentationExtension? Ext(PptxDocument doc) =>
        doc.Presentation.Presentation!.PresentationExtensionList?.Elements<P.PresentationExtension>().FirstOrDefault(e => e.Uri?.Value == Uri);

    static List<uint> Order(PptxDocument doc) =>
        doc.Presentation.Presentation!.SlideIdList?.Elements<P.SlideId>().Select(s => s.Id?.Value ?? 0u).ToList() ?? [];

    /// <summary>The sections in slide order, each with its first slide still in the deck (null: an empty one, kept after the one before it).</summary>
    static List<Sec> Read(PptxDocument doc)
    {
        if (Ext(doc)?.GetFirstChild<P14.SectionList>() is not { } list) return [];
        var order = Order(doc);
        return Sorted(order, list.Elements<P14.Section>().Select(s => new Sec(s.Name?.Value ?? "", s.Id?.Value ?? NewId(),
            s.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>().Select(e => e.Id?.Value ?? 0u).Where(order.Contains).OrderBy(order.IndexOf).Cast<uint?>().FirstOrDefault())));
    }

    static List<Sec> Sorted(List<uint> order, IEnumerable<Sec> secs)
    {
        double key = -1;
        return secs.Select(s => (Key: key = s.Start is { } f ? order.IndexOf(f) : key + 0.001, Sec: s)).ToList().OrderBy(k => k.Key).Select(k => k.Sec).ToList();
    }

    static string NewId() => "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";

    /// <summary>Writes the list: each section's slides from its start to the next start, the first one's from slide 1. No section with a slide: no list.</summary>
    static void Write(PptxDocument doc, List<Sec> secs)
    {
        var pres = doc.Presentation.Presentation!;
        var ext = Ext(doc);
        if (secs.All(s => s.Start is null))
        {
            ext?.Remove();
            if (pres.PresentationExtensionList is { HasChildren: false } empty) empty.Remove();
            return;
        }
        var order = Order(doc);
        var starts = secs.Select(s => s.Start is { } id ? order.IndexOf(id) : -1).ToList();
        starts[starts.FindIndex(i => i >= 0)] = 0;
        var list = new P14.SectionList();
        for (var i = 0; i < secs.Count; i++)
        {
            var end = starts.Skip(i + 1).FirstOrDefault(p => p >= 0, order.Count);
            var ids = starts[i] < 0 ? [] : order.Skip(starts[i]).Take(end - starts[i]);
            list.Append(new P14.Section(new P14.SectionSlideIdList(ids.Select(id => new P14.SectionSlideIdListEntry { Id = id }))) { Name = secs[i].Name, Id = secs[i].Id });
        }
        if (ext is null)
        {
            ext = new P.PresentationExtension { Uri = Uri };
            (pres.PresentationExtensionList ??= new P.PresentationExtensionList()).Append(ext);
        }
        ext.RemoveAllChildren();
        ext.Append(list);
    }

    /// <summary>The name of the section slide id starts, or null.</summary>
    internal static string? Starting(PptxDocument doc, uint slideId) => Read(doc).FirstOrDefault(s => s.Start == slideId)?.Name;

    /// <summary>A section starts at the slide under this name (renamed when one does); empty removes it, its slides joining the section
    /// before (the first section's the next). The first section made after slide 1 gets a Default Section in front, as in PowerPoint.</summary>
    internal static void Set(PptxDocument doc, uint slideId, string name)
    {
        var secs = Read(doc);
        var at = secs.FindIndex(s => s.Start == slideId);
        if (name.Length == 0) { if (at >= 0) secs.RemoveAt(at); }
        else if (at >= 0) secs[at] = secs[at] with { Name = name };
        else
        {
            if (!secs.Any(s => s.Start is not null) && Order(doc).IndexOf(slideId) > 0) secs.Insert(0, new Sec("Default Section", NewId(), Order(doc)[0]));
            secs.Add(new Sec(name, NewId(), slideId));
        }
        Write(doc, Sorted(Order(doc), secs));
    }

    /// <summary>Lays the list out again from the starts, for the save.</summary>
    internal static void Normalize(PptxDocument doc)
    {
        if (Ext(doc) is null) return;
        Write(doc, Read(doc));
    }
}
