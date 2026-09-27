using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WPS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace Writer.Formats.Docx;

/// <summary>
/// What Word draws in a paragraph that the tools do not edit: a chart, a SmartArt graphic, an embedded (OLE) object such as an Excel
/// sheet or a MathType equation, a group of pictures or a drawing canvas, an old VML drawing. Read-only, and never lost: it can be
/// moved or removed, and everything else in its paragraph can be edited around it. Its props say what an editor needs to show it as
/// Word does where the file keeps enough (a chart's cached values, SmartArt's laid-out shapes, an object's preview picture), else its
/// size and name. A paragraph holding only one (and no text) is a block of its own (/body/object[n]); one beside text is the
/// paragraph's child (/body/paragraph[n]/object[k]).
/// </summary>
sealed class DocxObject(DocxDocument doc, OpenXmlElement element, W.Paragraph? block = null) : Node
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string GroupUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    const string CanvasUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";

    public override string Kind => "object";
    public override object Anchor => block ?? element;

    /// <summary>The objects in a paragraph's own content (not in its text boxes, not an alternate content's fallback), in order: the
    /// mc:AlternateContent around a drawing when there is one, so the fallback goes with it.</summary>
    public static IEnumerable<OpenXmlElement> ElementsIn(W.Paragraph p)
    {
        foreach (var e in p.Descendants())
        {
            if (e is not (W.Drawing or W.EmbeddedObject or W.Picture) || e.Ancestors<W.Paragraph>().FirstOrDefault() != p || e.Ancestors<AlternateContentFallback>().Any()) continue;
            if (e is W.Drawing d && !IsObject(d)) continue;
            if (e is W.Picture && e.Parent is not W.Run) continue;
            yield return e.Parent is AlternateContentChoice { Parent: AlternateContent ac } ? ac : e;
        }
    }

    public static IEnumerable<Node> In(DocxDocument doc, W.Paragraph p) => ElementsIn(p).Select(e => (Node)new DocxObject(doc, e));

    /// <summary>A drawing that is neither a picture nor a shape or text box (DocxImage, DocxShape): a chart, SmartArt, a group, a canvas…</summary>
    static bool IsObject(W.Drawing d)
    {
        var data = d.Descendants<A.GraphicData>().FirstOrDefault();
        return data is not null && data.GetFirstChild<PIC.Picture>() is null && !d.Descendants<WPS.WordprocessingShape>().Any();
    }

    /// <summary>A paragraph with no text whose only content is one object (no picture, shape, equation or section break beside it).</summary>
    public static bool IsObjectParagraph(W.Paragraph p) =>
        p.ParagraphProperties?.SectionProperties is null
        && ElementsIn(p).Take(2).Count() == 1
        && DocxRuns.ParagraphText(p).Length == 0
        && !p.Descendants<PIC.Picture>().Any(pic => !ElementsIn(p).First().Descendants<PIC.Picture>().Contains(pic))
        && !p.Descendants<WPS.WordprocessingShape>().Any()
        && !p.Descendants<M.OfficeMath>().Any();

    public static DocxObject Block(DocxDocument doc, W.Paragraph p) => new(doc, ElementsIn(p).First(), p);

    W.Drawing? Drawing => element as W.Drawing ?? element.Descendants<W.Drawing>().FirstOrDefault();
    OpenXmlCompositeElement? Holder => (OpenXmlCompositeElement?)Drawing?.GetFirstChild<DW.Inline>() ?? Drawing?.GetFirstChild<DW.Anchor>();
    string? Uri => Drawing?.Descendants<A.GraphicData>().FirstOrDefault()?.Uri?.Value;

    string Type => element switch
    {
        W.EmbeddedObject => "ole",
        W.Picture => "vml",
        _ => Uri switch
        {
            OfficeGraphics.ChartUri or OfficeGraphics.ChartExUri => "chart",
            OfficeGraphics.DiagramUri => "smartart",
            GroupUri => "group",
            CanvasUri => "canvas",
            _ => "drawing",
        },
    };

    /// <summary>The part a relationship id of the document names, or null.</summary>
    OpenXmlPart? Part(string? id)
    {
        if (string.IsNullOrEmpty(id)) return null;
        try { return doc.Main.GetPartById(id); }
        catch (ArgumentOutOfRangeException) { return null; }
    }

    static string? RelAttr(OpenXmlElement? e, string localName) =>
        e?.GetAttributes().FirstOrDefault(a => a.LocalName == localName && a.NamespaceUri == RelNs).Value;

    static string? Attr(OpenXmlElement? e, string localName) => e?.GetAttributes().FirstOrDefault(a => a.LocalName == localName).Value;

    /// <summary>The VML shape of an OLE object or a w:pict: the one that says the size.</summary>
    OpenXmlElement? VmlShape => element.Descendants().FirstOrDefault(e => e.NamespaceUri == "urn:schemas-microsoft-com:vml" && e.LocalName is "shape" or "rect" or "roundrect" or "oval" or "group" or "line" or "polyline" or "image");

    /// <summary>A VML style's width or height in EMU ("width:415.5pt;height:250.5pt").</summary>
    static long? VmlLength(string? style, string name)
    {
        if (style is null) return null;
        foreach (var part in style.Split(';'))
        {
            var kv = part.Split(':', 2);
            if (kv.Length != 2 || kv[0].Trim() != name) continue;
            var v = kv[1].Trim();
            var unit = new string(v.SkipWhile(c => char.IsDigit(c) || c is '.' or '-').ToArray());
            if (!double.TryParse(v[..^unit.Length], NumberStyles.Float, Inv, out var n)) return null;
            return (long)Math.Round(n * unit switch { "pt" => 12700, "in" => 914400, "cm" => 360000, "mm" => 36000, "px" or "" => 9525, "pc" => 152400, _ => 12700 });
        }
        return null;
    }

    (long W, long H) Size
    {
        get
        {
            if (Holder?.GetFirstChild<DW.Extent>() is { } extent) return (extent.Cx?.Value ?? 0, extent.Cy?.Value ?? 0);
            var style = Attr(VmlShape, "style");
            long? w = VmlLength(style, "width"), h = VmlLength(style, "height");
            if (element is W.EmbeddedObject o)
            {
                w ??= long.TryParse(o.DxaOriginal?.Value, NumberStyles.Integer, Inv, out var dx) ? dx * 635 : null;
                h ??= long.TryParse(o.DyaOriginal?.Value, NumberStyles.Integer, Inv, out var dy) ? dy * 635 : null;
            }
            return (w ?? 0, h ?? 0);
        }
    }

    /// <summary>The picture Word shows for an OLE object or a VML drawing (v:imagedata), when the browser can draw it.</summary>
    ImagePart? Preview =>
        element is W.EmbeddedObject or W.Picture
        && Part(RelAttr(element.Descendants().FirstOrDefault(e => e.LocalName == "imagedata"), "id")) is ImagePart img
        && img.ContentType is "image/png" or "image/jpeg" or "image/gif" or "image/bmp" or "image/svg+xml" or "image/x-wmf" or "image/wmf" ? img : null;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var type = Type;
        var props = new Dictionary<string, string> { ["type"] = type, ["xml"] = element.OuterXml };
        var (w, h) = Size;
        if (block is null && element.Ancestors<W.Paragraph>().FirstOrDefault() is { } para)
            props["at"] = DocxFootnotes.OffsetOf(para, element).ToString(Inv);
        if (w > 0) props["width"] = w.ToString(Inv);
        if (h > 0) props["height"] = h.ToString(Inv);
        var docPr = Holder?.GetFirstChild<DW.DocProperties>();
        if (docPr?.Name?.Value is { Length: > 0 } name) props["name"] = name;
        if (docPr?.Description?.Value is { Length: > 0 } alt) props["alt"] = alt;
        if (docPr?.Id?.Value is { } id) props["id"] = id.ToString(Inv);
        if (Drawing?.GetFirstChild<DW.Anchor>() is not null) props["wrap"] = "float";
        var data = Drawing?.Descendants<A.GraphicData>().FirstOrDefault()?.FirstChild;
        switch (type)
        {
            case "chart":
                var part = Part(RelAttr(data, "id"));
                var json = part is null ? null : Uri == OfficeGraphics.ChartExUri ? OfficeGraphics.ChartExJson(part, doc.Scheme) : OfficeGraphics.ChartJson(part, doc.Scheme);
                if (json is not null)
                {
                    props["chart"] = json;
                    if (OfficeGraphics.ChartTitle(json) is { Length: > 0 } title) props["title"] = title;
                }
                break;
            case "smartart":
                var dataPart = Part(RelAttr(data, "dm"));
                var drawingPart = dataPart is null ? null : Part(OfficeGraphics.SmartArtDrawingId(dataPart));
                if (OfficeGraphics.SmartArtJson(drawingPart, dataPart, doc.Scheme) is { } smartart) props["smartart"] = smartart;
                break;
            case "ole":
                if (Attr(element.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject"), "ProgID") is { Length: > 0 } progId) props["progId"] = progId;
                break;
        }
        if (Preview is { } preview)
        {
            props["src"] = preview.Uri.OriginalString;
            if (preview.ContentType is "image/x-wmf" or "image/wmf") props["previewFormat"] = "wmf";
        }
        if (type == "vml" && element.Descendants<W.TextBoxContent>().FirstOrDefault() is { } box)
            props["text"] = string.Join("\n", box.Elements<W.Paragraph>().Select(DocxRuns.ParagraphText));
        return props;
    }

    public override (string ContentType, byte[] Data)? GetBinary()
    {
        if (Preview is not { } img) return null;
        using var stream = img.GetStream(FileMode.Open, FileAccess.Read);
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return (img.ContentType, ms.ToArray());
    }

    public override string GetRaw() => (block ?? element).OuterXml;

    /// <summary>History restores retain their IDs; copies sharing a document get fresh drawing/VML IDs.
    /// Related parts remain in the package after removal, so restoring their relationship IDs is lossless.</summary>
    internal static void UniqueIds(DocxDocument doc, OpenXmlElement copy)
    {
        var ids = doc.Main.Document!.Descendants<DW.DocProperties>().Select(e => e.Id?.Value ?? 0).ToHashSet();
        uint next = ids.Count == 0 ? 1 : ids.Max() + 1;
        foreach (var e in copy.Descendants<DW.DocProperties>())
        {
            var id = e.Id?.Value ?? 0;
            if (id == 0 || !ids.Add(id)) { while (ids.Contains(next)) next++; e.Id = next; ids.Add(next++); }
        }
        const string vml = "urn:schemas-microsoft-com:vml";
        var used = doc.Main.Document.Descendants().Where(e => e.NamespaceUri == vml).Select(e => Attr(e, "id")).OfType<string>().ToHashSet();
        var renamed = new Dictionary<string, string>();
        var number = 1025;
        foreach (var e in copy.Descendants().Where(e => e.NamespaceUri == vml && e.LocalName != "shapetype"))
        {
            var id = Attr(e, "id");
            if (id is null || used.Add(id)) continue;
            string replacement; do { replacement = "_x0000_i" + number++; } while (!used.Add(replacement));
            renamed[id] = replacement;
            e.SetAttribute(new OpenXmlAttribute("", "id", "", replacement));
        }
        foreach (var e in copy.Descendants())
            foreach (var a in e.GetAttributes().Where(a => a.LocalName is "ShapeID" or "spid").ToArray())
                if (a.Value is { } value && renamed.TryGetValue(value, out var replacement)) e.SetAttribute(new OpenXmlAttribute(a.Prefix, a.LocalName, a.NamespaceUri, replacement));
    }

    public override void SetProp(string name, string value)
    {
        if (name != "at" || block is not null) return;
        var paragraph = element.Ancestors<W.Paragraph>().FirstOrDefault();
        if (paragraph is null) return;
        var run = TakeRun();
        DocxFootnotes.Place(paragraph, run, int.Parse(value, Inv));
    }

    // ---- where it lives ----

    public override void Remove()
    {
        if (block is not null)
        {
            DocxBlocks.Detach(block);
            return;
        }
        var run = element.Parent as W.Run;
        element.Remove();
        if (run is not null && !run.ChildElements.Any(e => e is not W.RunProperties)) DocxImage.Cut(run);
    }

    /// <summary>Into a paragraph: at the end of its runs (or at index). Into the body or a cell: a paragraph of its own. The paragraph a block
    /// object leaves goes with it.</summary>
    public override void MoveTo(Node newParent, int? index)
    {
        var container = DocxBlocks.ContainerOf(newParent);
        if (block is not null && container is not W.Paragraph)
        {
            DocxBlocks.Move(newParent, block, index);
            return;
        }
        if (container.Ancestors().Contains(block ?? element) || ReferenceEquals(container, block))
            throw new WriterException(ErrorCode.Validation, "Cannot move an element into itself", "Pick another target.");
        var run = TakeRun();
        block = container is W.Paragraph ? null : new W.Paragraph(run);
        DocxBlocks.InsertAt(newParent, container, (OpenXmlElement?)block ?? run, index);
    }

    /// <summary>The object's run, out of its place: its own run when that holds nothing else, else a new one with its formatting.</summary>
    W.Run TakeRun()
    {
        var old = element.Parent as W.Run;
        if (block is not null) DocxBlocks.Detach(block);
        if (old is not null && old.ChildElements.All(e => e is W.RunProperties || ReferenceEquals(e, element)))
        {
            if (block is null) DocxImage.Cut(old); else old.Remove();
            return old;
        }
        element.Remove();
        var run = new W.Run();
        if (old?.RunProperties is { } rp) run.Append(rp.CloneNode(true));
        run.Append(element);
        return run;
    }
}
