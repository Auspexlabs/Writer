using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using Writer.Formats.Common;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

sealed class PptxRoot(PptxDocument doc) : Node
{
    public override string Kind => "document";
    public override string Format => "pptx";
    public override object Anchor => doc.Package;

    protected override IEnumerable<Node> ProjectChildren() => doc.Slides.Select(s => (Node)new PptxSlide(doc, s));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var (width, height) = doc.SlideSize;
        var props = new Dictionary<string, string>
        {
            ["format"] = "pptx",
            ["slides"] = doc.Slides.Count.ToString(CultureInfo.InvariantCulture),
            ["width"] = width.ToString(CultureInfo.InvariantCulture),
            ["height"] = height.ToString(CultureInfo.InvariantCulture),
        };
        if (doc.Package.PackageProperties.Title is { Length: > 0 } title) props["title"] = title;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        if (name == "title") doc.Package.PackageProperties.Title = value.Length > 0 ? value : null;
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var slide = doc.AddSlide(props.GetValueOrDefault("layout"), index);
        var node = new PptxSlide(doc, slide);
        foreach (var (name, value) in props)
            if (name != "layout") node.SetProp(name, value);
        return node;
    }
}

sealed class PptxSlide(PptxDocument doc, SlidePart slide) : Node
{
    public override string Kind => "slide";
    public override object Anchor => slide;
    OpenXmlElement Tree => PptxDocument.ShapeTree(slide);

    /// <summary>Decor inherited from the master and layout comes first (it is drawn underneath), then the slide's own shapes.</summary>
    protected override IEnumerable<Node> ProjectChildren()
    {
        foreach (var decor in PptxDecor.Of(doc, slide)) yield return decor;
        foreach (var child in Tree.ChildElements)
        {
            switch (child)
            {
                case P.Shape sp: yield return new PptxShape(doc, slide, sp); break;
                case P.Picture pic: yield return new PptxImage(doc, slide, pic); break;
                case P.GraphicFrame frame when PptxTable.TableOf(frame) is not null: yield return new PptxTable(doc, slide, frame); break;
            }
        }
    }

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string>();
        if (TitleShape() is { } title) props["title"] = PptxText.BodyText(title.TextBody);
        if (slide.SlideLayoutPart is { } layout) props["layout"] = PptxDocument.LayoutName(layout);
        if (PptxDecor.Background(slide)?.Background.BackgroundProperties is { } bg)
        {
            if (bg.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value is { } color) props["background"] = color.ToUpperInvariant();
            if (bg.GetFirstChild<A.BlipFill>() is not null) props["backgroundImage"] = "true";
        }
        if (slide.NotesSlidePart is { } notes && PptxDocument.NotesBody(notes) is { } body) props["notes"] = PptxText.BodyText(body.TextBody);
        if (slide.Slide?.Show?.Value == false) props["hidden"] = "true";
        if (Transition() is { } transition)
        {
            if (TransitionName(transition) is { } name) props["transition"] = name;
            if (int.TryParse(transition.Duration?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var ms)) props["duration"] = ms.ToString(CultureInfo.InvariantCulture);
        }
        props["id"] = doc.SlideIdOf(slide).Id?.Value.ToString(CultureInfo.InvariantCulture) ?? "";
        return props;
    }

    /// <summary>The background the slide shows when it declares none of its own: its layout's or master's, theme colours resolved.</summary>
    public override IReadOnlyDictionary<string, string>? GetComputed(IReadOnlyDictionary<string, string> props)
    {
        if (props.ContainsKey("background") || props.ContainsKey("backgroundImage")) return null;
        var background = PptxLook.For(slide, doc.Presentation).Background(PptxDecor.Background(slide)?.Background);
        return background is null or "none" ? null : new Dictionary<string, string> { ["background"] = background };
    }

    /// <summary>The child element PowerPoint writes for each transition the engine models.</summary>
    static readonly Dictionary<string, string> Transitions = new()
    {
        ["fade"] = "<p:fade/>",
        ["push"] = "<p:push dir=\"l\"/>",
        ["wipe"] = "<p:wipe dir=\"r\"/>",
        ["split"] = "<p:split orient=\"horz\" dir=\"out\"/>",
        ["cover"] = "<p:cover dir=\"l\"/>",
        ["cut"] = "<p:cut/>",
        ["dissolve"] = "<p:dissolve/>",
        ["zoom"] = "<p:zoom/>",
        ["random"] = "<p:random/>",
    };

    const string P14Ns = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    /// <summary>The slide's p:transition, whether bare or inside the mc:AlternateContent PowerPoint 2010+ wraps it in.</summary>
    P.Transition? Transition() =>
        slide.Slide!.Transition ?? slide.Slide.Elements<AlternateContent>().Select(ac => ac.GetFirstChild<AlternateContentChoice>()?.GetFirstChild<P.Transition>()).FirstOrDefault(t => t is not null);

    static OpenXmlElement? TransitionType(P.Transition transition) => transition.ChildElements.FirstOrDefault(c => c is not (P.SoundAction or P.ExtensionList));

    static string? TransitionName(P.Transition transition) => TransitionType(transition) is { } type
        ? type.NamespaceUri == PptxTemplate.PNs && Transitions.ContainsKey(type.LocalName) ? type.LocalName : "other"
        : null;

    /// <summary>Rewrites the transition: none removes it; a duration puts it in the AlternateContent form PowerPoint uses for p14:dur.</summary>
    void EditTransition(string? type, string? duration)
    {
        var transition = Transition() ?? new P.Transition();
        var holder = transition.Parent is AlternateContentChoice choice ? choice.Parent : transition.Parent is null ? null : transition;
        holder?.Remove();
        if (type == "none") return;
        if (transition.Parent is not null) transition.Remove();
        if (type is not null)
        {
            foreach (var old in transition.ChildElements.Where(c => c is not (P.SoundAction or P.ExtensionList)).ToList()) old.Remove();
            var fresh = new P.Transition($"<p:transition {PptxTemplate.Ns}>{Transitions[type]}</p:transition>").FirstChild!;
            fresh.Remove();
            transition.PrependChild(fresh);
        }
        if (duration is not null) transition.Duration = duration;
        OpenXmlElement element = transition;
        if (transition.Duration is not null)
        {
            var fallback = (P.Transition)transition.CloneNode(true);
            fallback.Duration = null;
            var p14 = new AlternateContentChoice(transition) { Requires = "p14" };
            p14.AddNamespaceDeclaration("p14", P14Ns);
            element = new AlternateContent(p14, new AlternateContentFallback(fallback));
        }
        var sld = slide.Slide!;
        if ((sld.Timing ?? (OpenXmlElement?)sld.SlideExtensionList) is { } before) sld.InsertBefore(element, before);
        else sld.Append(element);
    }

    /// <summary>Puts a slide-level element at a 1-based position among the slide's children; a position inside the decor block means first.</summary>
    internal void Insert(OpenXmlElement element, int? index)
    {
        if (index is { } i) index = Math.Max(i, Children.Count(c => c.Kind == "decor") + 1);
        if (!OoxmlTree.InsertBefore(this, Tree, element, index)) Tree.Append(element);
    }

    P.Shape? TitleShape() => Tree.Elements<P.Shape>().FirstOrDefault(s =>
        s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.InnerText is "title" or "ctrTitle");

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "title":
                var title = TitleShape() ?? AddTitleShape();
                PptxText.SetBody(title.TextBody ??= new P.TextBody(new A.BodyProperties(), new A.ListStyle()), slide, [new RunSpec(value)]);
                break;
            case "layout":
                var layout = doc.FindLayout(value);
                if (slide.SlideLayoutPart is { } current && !ReferenceEquals(current, layout)) slide.DeletePart(current);
                if (slide.SlideLayoutPart is null) slide.AddPart(layout);
                break;
            case "background":
                var data = slide.Slide!.CommonSlideData!;
                if (value == "none")
                {
                    data.Background?.Remove();
                    break;
                }
                data.Background = new P.Background(new P.BackgroundProperties(new A.SolidFill(new A.RgbColorModelHex { Val = value }), new A.EffectList()));
                break;
            case "notes":
                if (value.Length == 0 && slide.NotesSlidePart is null) break;
                var notes = PptxDocument.NotesBody(doc.EnsureNotes(slide))
                    ?? throw new WriterException(ErrorCode.Validation, "The notes page has no text placeholder", "Open the notes page in PowerPoint and add one.");
                PptxText.SetBody(notes.TextBody ??= new P.TextBody(new A.BodyProperties(), new A.ListStyle()), slide, [new RunSpec(value)]);
                break;
            case "hidden":
                slide.Slide!.Show = value == "true" ? false : null;
                break;
            case "transition":
                if (value == "other")
                    throw new WriterException(ErrorCode.Validation, "'other' only describes a transition the engine does not model", $"Choose one of: none, {string.Join(", ", Transitions.Keys)}.");
                EditTransition(value, null);
                break;
            case "duration":
                EditTransition(null, value);
                break;
        }
    }

    P.Shape AddTitleShape()
    {
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = PptxDocument.NextShapeId(slide), Name = "Title" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
            new P.ShapeProperties(new A.Transform2D(new A.Offset { X = 838200L, Y = 365125L }, new A.Extents { Cx = 10515600L, Cy = 1325563L })),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));
        Tree.Append(shape);
        return shape;
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var (element, consumed) = kind switch
        {
            "shape" => PptxShape.New(doc, slide, props),
            "image" => PptxImage.New(doc, slide, props),
            "table" => PptxTable.New(doc, slide, props),
            _ => throw new WriterException(ErrorCode.UnsupportedKind, $"Cannot add {kind} to a slide", "Slide children: shape, image, table."),
        };
        var node = Place(element, index);
        foreach (var (name, value) in props)
            if (!consumed.Contains(name)) node.SetProp(name, value);
        return node;
    }

    internal SlidePart Part => slide;

    /// <summary>Puts a shape, picture or table element on the slide and returns its node.</summary>
    internal Node Place(OpenXmlElement element, int? index)
    {
        Insert(element, index);
        return element switch
        {
            P.Shape sp => new PptxShape(doc, slide, sp),
            P.Picture pic => new PptxImage(doc, slide, pic),
            _ => new PptxTable(doc, slide, (P.GraphicFrame)element),
        };
    }

    /// <summary>A shape, picture or table put back from its raw XML. Its relationship ids are this slide's, as a removed element's
    /// still are; its drawing ids stay unless another shape has them now.</summary>
    public override Node AddRaw(string raw, int? index)
    {
        var element = RawXml.Parse(Tree, raw, doc.Namespaces);
        if (element is not (P.Shape or P.Picture) && !(element is P.GraphicFrame frame && PptxTable.TableOf(frame) is not null))
            throw new WriterException(ErrorCode.Validation, $"Raw XML on a slide must be a shape, picture or table, got <{element.Prefix}:{element.LocalName}>",
                "Pass one <p:sp>, <p:pic> or table <p:graphicFrame> as 'get --raw' printed it.");
        PptxCopy.Ids(element, slide, keep: true);
        return Place(element, index);
    }

    public override Node CopyTo(Node newParent, int? index) => newParent is PptxRoot
        ? new PptxSlide(doc, doc.CopySlide(slide, index))
        : throw new WriterException(ErrorCode.Validation, "A slide is copied into the presentation", "Use --to / with --index, --after or --before.");

    public override string GetRaw() => slide.Slide!.OuterXml;

    public override void SetRaw(string raw)
    {
        try
        {
            slide.Slide = new P.Slide(RawXml.WithNamespaces(raw, doc.Namespaces));
        }
        catch (Exception ex) when (ex is not WriterException)
        {
            throw new WriterException(ErrorCode.Validation, $"Raw slide XML could not be parsed: {ex.Message}", "Start from 'get --raw' output.");
        }
    }

    public override void Remove() => doc.RemoveSlide(slide);

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent.Kind != "document")
            throw new WriterException(ErrorCode.Validation, "Slides can only be reordered within the presentation", "Use --to / with --index.");
        doc.MoveSlide(slide, index);
    }
}

/// <summary>A p:sp: text box, placeholder or drawn shape.</summary>
sealed class PptxShape(PptxDocument doc, SlidePart slide, P.Shape shape) : Node
{
    public override string Kind => "shape";
    public override object Anchor => shape;

    protected override IEnumerable<Node> ProjectChildren() =>
        PptxText.Paragraphs(shape.TextBody).Select(p => (Node)new PptxParagraph(doc, slide, p));

    P.NonVisualDrawingProperties? Nv => shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
    P.PlaceholderShape? Placeholder => shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = PptxText.BodyText(shape.TextBody) };
        props["html"] = string.Join("<br>", Children.Where(c => c.Kind == "paragraph").Select(Exporter.HtmlOf));
        if (shape.NonVisualShapeProperties?.NonVisualShapeDrawingProperties?.TextBox?.Value == true) props["geometry"] = "textbox";
        else if (shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText is { } preset) props["geometry"] = preset;
        else if (shape.ShapeProperties?.GetFirstChild<A.CustomGeometry>() is not null) props["geometry"] = "custom";
        var xfrm = shape.ShapeProperties?.Transform2D ?? InheritedTransform();
        AddBox(props, xfrm?.Offset?.X?.Value, xfrm?.Offset?.Y?.Value, xfrm?.Extents?.Cx?.Value, xfrm?.Extents?.Cy?.Value);
        if (shape.ShapeProperties is { } spPr)
        {
            if (spPr.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value is { } fill) props["fill"] = fill.ToUpperInvariant();
            else if (spPr.GetFirstChild<A.NoFill>() is not null) props["fill"] = "none";
            if (spPr.GetFirstChild<A.Outline>() is { } line)
            {
                if (line.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value is { } lineColor) props["line"] = lineColor.ToUpperInvariant();
                else if (line.GetFirstChild<A.NoFill>() is not null) props["line"] = "none";
            }
        }
        var firstRun = PptxText.Paragraphs(shape.TextBody).SelectMany(p => p.Elements<A.Run>()).FirstOrDefault()?.RunProperties;
        if (firstRun?.GetFirstChild<A.LatinFont>()?.Typeface?.Value is { } font && !font.StartsWith('+')) props["font"] = font;
        if (firstRun?.FontSize?.Value is { } size) props["size"] = PptxText.Points(size);
        if (PptxText.Color(firstRun) is { } color) props["color"] = color;
        if (Placeholder is { } ph) props["placeholder"] = ph.Type?.InnerText switch { null => "body", "ctrTitle" => "title", "subTitle" => "subtitle", var t => t };
        if (Nv?.Name?.Value is { Length: > 0 } name) props["name"] = name;
        if (Nv?.Id?.Value is { } id) props["id"] = id.ToString(CultureInfo.InvariantCulture);
        return props;
    }

    public override IReadOnlyDictionary<string, string>? GetComputed(IReadOnlyDictionary<string, string> props) =>
        PptxLook.For(slide, doc.Presentation).Shape(shape, props);

    internal static void AddBox(Dictionary<string, string> props, long? x, long? y, long? w, long? h)
    {
        if (x is { } px) props["x"] = px.ToString(CultureInfo.InvariantCulture);
        if (y is { } py) props["y"] = py.ToString(CultureInfo.InvariantCulture);
        if (w is { } pw) props["w"] = pw.ToString(CultureInfo.InvariantCulture);
        if (h is { } ph) props["h"] = ph.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>Placeholders without their own position take it from the layout, then the master.</summary>
    A.Transform2D? InheritedTransform()
    {
        if (Placeholder is not { } ph) return null;
        foreach (var part in new OpenXmlPart?[] { slide.SlideLayoutPart, slide.SlideLayoutPart?.SlideMasterPart })
        {
            var shapes = part switch
            {
                SlideLayoutPart l => l.SlideLayout?.CommonSlideData?.ShapeTree?.Elements<P.Shape>(),
                SlideMasterPart m => m.SlideMaster?.CommonSlideData?.ShapeTree?.Elements<P.Shape>(),
                _ => null,
            };
            var match = shapes?.FirstOrDefault(s => Matches(s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape, ph));
            if (match?.ShapeProperties?.Transform2D is { } xfrm) return xfrm;
        }
        return null;
    }

    static bool Matches(P.PlaceholderShape? candidate, P.PlaceholderShape ph)
    {
        if (candidate is null) return false;
        var type = ph.Type?.InnerText ?? "body";
        var candidateType = candidate.Type?.InnerText ?? "body";
        if (ph.Index is not null && candidate.Index?.Value == ph.Index.Value) return true;
        if (type is "title" or "ctrTitle") return candidateType is "title" or "ctrTitle";
        return candidateType == type && (ph.Index is null || candidate.Index is null);
    }

    public override void SetProp(string name, string value)
    {
        var spPr = shape.ShapeProperties ??= new P.ShapeProperties();
        switch (name)
        {
            case "text":
                PptxText.SetBody(Body(), slide, [new RunSpec(value)]);
                break;
            case "md":
                PptxText.SetBody(Body(), slide, InlineMarkdown.Parse(value));
                break;
            case "html":
                PptxText.SetBody(Body(), slide, InlineHtml.Parse(value));
                break;
            case "geometry":
                var nv = shape.NonVisualShapeProperties!.NonVisualShapeDrawingProperties ??= new P.NonVisualShapeDrawingProperties();
                nv.TextBox = value == "textbox" ? true : null;
                spPr.RemoveAllChildren<A.PresetGeometry>();
                spPr.RemoveAllChildren<A.CustomGeometry>();
                var geometry = new A.PresetGeometry(new A.AdjustValueList()) { Preset = new A.ShapeTypeValues(value == "textbox" ? "rect" : value) };
                if (spPr.Transform2D is { } xfrm) spPr.InsertAfter(geometry, xfrm);
                else spPr.PrependChild(geometry);
                break;
            case "x" or "y" or "w" or "h":
                var transform = spPr.Transform2D ?? Materialise(spPr);
                var emu = long.Parse(value, CultureInfo.InvariantCulture);
                switch (name)
                {
                    case "x": transform.Offset!.X = emu; break;
                    case "y": transform.Offset!.Y = emu; break;
                    case "w": transform.Extents!.Cx = emu; break;
                    default: transform.Extents!.Cy = emu; break;
                }
                break;
            case "fill":
                PptxText.SetColor(spPr, value);
                if (value == "none") PptxText.InsertBeforeAny(spPr, new A.NoFill(), e => e is A.Outline or A.EffectList or A.EffectDag or A.Scene3DType or A.Shape3DType or A.ExtensionList);
                break;
            case "line":
                var line = spPr.GetFirstChild<A.Outline>();
                if (line is null)
                {
                    line = new A.Outline();
                    PptxText.InsertBeforeAny(spPr, line, e => e is A.EffectList or A.EffectDag or A.Scene3DType or A.Shape3DType or A.ExtensionList);
                }
                PptxText.SetColor(line, value);
                if (value == "none") line.PrependChild(new A.NoFill());
                break;
            case "font" or "size" or "color":
                foreach (var run in PptxText.Paragraphs(shape.TextBody).SelectMany(p => p.Elements<A.Run>()))
                {
                    var rPr = run.RunProperties ??= new A.RunProperties();
                    if (name == "font") PptxText.SetFont(rPr, value);
                    else if (name == "size") rPr.FontSize = PptxText.Hundredths(value);
                    else PptxText.SetColor(rPr, value);
                }
                break;
        }
    }

    P.TextBody Body() => shape.TextBody ??= new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph());

    A.Transform2D Materialise(P.ShapeProperties spPr)
    {
        var inherited = InheritedTransform();
        var xfrm = new A.Transform2D(
            new A.Offset { X = inherited?.Offset?.X?.Value ?? 914400L, Y = inherited?.Offset?.Y?.Value ?? 914400L },
            new A.Extents { Cx = inherited?.Extents?.Cx?.Value ?? 4572000L, Cy = inherited?.Extents?.Cy?.Value ?? 914400L });
        spPr.PrependChild(xfrm);
        return xfrm;
    }

    public static (OpenXmlElement Element, string[] Consumed) New(PptxDocument doc, SlidePart slide, IReadOnlyDictionary<string, string> props)
    {
        var geometry = props.GetValueOrDefault("geometry") ?? "textbox";
        var textBox = geometry == "textbox";
        var (width, _) = doc.SlideSize;
        var id = PptxDocument.NextShapeId(slide);
        var spPr = new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = Emu(props, "x", 914400L), Y = Emu(props, "y", 914400L) },
                new A.Extents { Cx = Emu(props, "w", width - 2 * 914400L), Cy = Emu(props, "h", 914400L) }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = new A.ShapeTypeValues(textBox ? "rect" : geometry) });
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = (textBox ? "TextBox " : "Shape ") + id },
                textBox ? new P.NonVisualShapeDrawingProperties { TextBox = true } : new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            spPr);
        if (!textBox)
            shape.Append(new P.ShapeStyle(
                new A.LineReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 2U },
                new A.FillReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 1U },
                new A.EffectReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
                new A.FontReference(new A.SchemeColor { Val = A.SchemeColorValues.Light1 }) { Index = A.FontCollectionIndexValues.Minor }));
        shape.Append(new P.TextBody(
            textBox ? new A.BodyProperties(new A.ShapeAutoFit()) { Wrap = A.TextWrappingValues.Square, RightToLeftColumns = false } : new A.BodyProperties { Anchor = A.TextAnchoringTypeValues.Center },
            new A.ListStyle(),
            new A.Paragraph()));
        return (shape, ["geometry", "x", "y", "w", "h"]);
    }

    internal static long Emu(IReadOnlyDictionary<string, string> props, string name, long fallback) =>
        props.TryGetValue(name, out var v) ? long.Parse(v, CultureInfo.InvariantCulture) : fallback;

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var body = Body();
        var paragraph = new A.Paragraph();
        if (body.Elements<A.Paragraph>().LastOrDefault()?.ParagraphProperties is { } pPr) paragraph.Append(pPr.CloneNode(true));
        if (!OoxmlTree.InsertBefore(this, body, paragraph, index)) body.Append(paragraph);
        var node = new PptxParagraph(doc, slide, paragraph);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override string GetRaw() => shape.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(shape, raw, doc.Namespaces);
    public override void Remove() => shape.Remove();

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent is not PptxSlide target)
            throw new WriterException(ErrorCode.Validation, "Shapes move between slides only", "Use --to /slide[n].");
        shape.Remove();
        target.Insert(shape, index);
    }

    public override Node CopyTo(Node newParent, int? index) => newParent is PptxSlide target
        ? target.Place(PptxCopy.Element(shape, slide, target.Part), index)
        : throw new WriterException(ErrorCode.Validation, "Shapes are copied onto slides", "Use --to /slide[n].");
}

sealed class PptxParagraph(PptxDocument doc, SlidePart slide, A.Paragraph p) : Node
{
    public override string Kind => "paragraph";
    public override object Anchor => p;

    protected override IEnumerable<Node> ProjectChildren() => p.Elements<A.Run>().Select(r => (Node)new PptxRun(doc, slide, r));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = PptxText.ParagraphText(p), ["html"] = Exporter.HtmlOf(this) };
        var pPr = p.ParagraphProperties;
        if (PptxText.ListOf(pPr) is { } list)
        {
            props["list"] = list;
            props["level"] = (pPr?.Level?.Value ?? 0).ToString(CultureInfo.InvariantCulture);
        }
        else if (pPr?.Level?.Value is { } level and > 0) props["level"] = level.ToString(CultureInfo.InvariantCulture);
        if (PptxText.AlignOf(pPr) is { } align) props["align"] = align;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "text": PptxText.SetParagraph(p, slide, [new RunSpec(value)]); break;
            case "md": PptxText.SetParagraph(p, slide, InlineMarkdown.Parse(value)); break;
            case "html": PptxText.SetParagraph(p, slide, InlineHtml.Parse(value)); break;
            case "list": PptxText.SetList(Properties(), value); break;
            case "level":
                var pPr = Properties();
                pPr.Level = int.Parse(value, CultureInfo.InvariantCulture);
                if (PptxText.ListOf(pPr) is { } kind) PptxText.SetList(pPr, kind);
                break;
            case "align": Properties().Alignment = PptxText.AlignValue(value); break;
        }
    }

    A.ParagraphProperties Properties() => p.ParagraphProperties ??= new A.ParagraphProperties();

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        if (!props.ContainsKey("text") && !props.ContainsKey("md"))
            throw new WriterException(ErrorCode.Validation, "A run needs text", "Add --prop text=\"...\" or --prop md=\"...\".");
        var run = new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(""));
        if (!OoxmlTree.InsertBefore(this, p, run, index))
        {
            if (p.GetFirstChild<A.EndParagraphRunProperties>() is { } end) p.InsertBefore(run, end);
            else p.Append(run);
        }
        var node = new PptxRun(doc, slide, run);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override string GetRaw() => p.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(p, raw, doc.Namespaces);

    public override void Remove()
    {
        var body = p.Parent;
        p.Remove();
        if (body is not null && !body.Elements<A.Paragraph>().Any()) body.Append(new A.Paragraph());
    }

    public override void MoveTo(Node newParent, int? index)
    {
        OpenXmlCompositeElement body = newParent switch
        {
            PptxShape s => ((P.Shape)s.Anchor).TextBody ?? throw NoBody(),
            PptxCell c => ((A.TableCell)c.Anchor).TextBody ?? throw NoBody(),
            _ => throw new WriterException(ErrorCode.Validation, $"Paragraphs cannot move under {newParent.Kind}", "Move them into a shape or a cell."),
        };
        Remove();
        if (!OoxmlTree.InsertBefore(newParent, body, p, index)) body.Append(p);
    }

    static WriterException NoBody() => new(ErrorCode.Validation, "The target has no text body", "Set its text first.");
}

sealed class PptxRun(PptxDocument doc, SlidePart slide, A.Run run) : Node
{
    public override string Kind => "run";
    public override object Anchor => run;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = run.Text?.Text ?? "" };
        var rPr = run.RunProperties;
        if (rPr?.Bold?.Value == true) props["bold"] = "true";
        if (rPr?.Italic?.Value == true) props["italic"] = "true";
        if (rPr?.Underline?.InnerText is { } u && u != "none") props["underline"] = "true";
        if (rPr?.Strike?.InnerText is "sngStrike" or "dblStrike") props["strike"] = "true";
        if (PptxText.Color(rPr) is { } color) props["color"] = color;
        if (rPr?.FontSize?.Value is { } size) props["size"] = PptxText.Points(size);
        if (rPr?.GetFirstChild<A.LatinFont>()?.Typeface?.Value is { } font && !font.StartsWith('+')) props["font"] = font;
        if (PptxText.Highlight(rPr) is { } highlight) props["highlight"] = highlight;
        if (PptxText.Link(rPr, slide) is { } link) props["link"] = link;
        return props;
    }

    public override void SetProp(string name, string value)
    {
        var rPr = run.RunProperties ??= new A.RunProperties();
        var on = value == "true";
        switch (name)
        {
            case "text":
                var lines = value.ReplaceLineEndings("\n").Split('\n');
                run.Text = new A.Text(lines[0]);
                OpenXmlElement cursor = run;
                for (var i = 1; i < lines.Length; i++)
                {
                    var br = new A.Break();
                    run.Parent!.InsertAfter(br, cursor);
                    cursor = br;
                    if (lines[i].Length == 0) continue;
                    var extra = new A.Run((A.RunProperties)rPr.CloneNode(true), new A.Text(lines[i]));
                    run.Parent.InsertAfter(extra, cursor);
                    cursor = extra;
                }
                break;
            case "md": run.Text = new A.Text(string.Concat(InlineMarkdown.Parse(value).Select(s => s.Text))); break;
            case "bold": rPr.Bold = on; break;
            case "italic": rPr.Italic = on; break;
            case "underline": rPr.Underline = on ? A.TextUnderlineValues.Single : A.TextUnderlineValues.None; break;
            case "strike": rPr.Strike = on ? A.TextStrikeValues.SingleStrike : A.TextStrikeValues.NoStrike; break;
            case "color": PptxText.SetColor(rPr, value); break;
            case "highlight": PptxText.SetHighlight(rPr, value); break;
            case "size": rPr.FontSize = PptxText.Hundredths(value); break;
            case "font": PptxText.SetFont(rPr, value); break;
            case "link": PptxText.SetLink(rPr, slide, value); break;
        }
        _ = doc;
    }

    public override string GetRaw() => run.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(run, raw, doc.Namespaces);
    public override void Remove() => run.Remove();

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent is not PptxParagraph target)
            throw new WriterException(ErrorCode.Validation, "Runs move between paragraphs only", "Use --to a paragraph path.");
        run.Remove();
        var p = (A.Paragraph)target.Anchor;
        if (OoxmlTree.InsertBefore(target, p, run, index)) return;
        if (p.GetFirstChild<A.EndParagraphRunProperties>() is { } end) p.InsertBefore(run, end);
        else p.Append(run);
    }
}

/// <summary>A p:pic. The part is the slide it sits on, or the layout or master when the picture is decor.</summary>
sealed class PptxImage(PptxDocument doc, OpenXmlPart part, P.Picture picture) : PictureNode
{
    public override object Anchor => picture;
    protected override OpenXmlCompositeElement Pic => picture;
    protected override OpenXmlPart Owner => part;

    protected override (long X, long Y, long W, long H) Frame
    {
        get
        {
            var xfrm = picture.ShapeProperties?.Transform2D;
            return (xfrm?.Offset?.X?.Value ?? 0, xfrm?.Offset?.Y?.Value ?? 0, xfrm?.Extents?.Cx?.Value ?? 0, xfrm?.Extents?.Cy?.Value ?? 0);
        }
        set
        {
            var xfrm = Transform();
            (xfrm.Offset!.X, xfrm.Offset.Y, xfrm.Extents!.Cx, xfrm.Extents.Cy) = (value.X, value.Y, value.W, value.H);
        }
    }

    A.Transform2D Transform()
    {
        var spPr = picture.ShapeProperties ??= new P.ShapeProperties();
        var xfrm = spPr.Transform2D ?? (A.Transform2D)spPr.PrependChild(new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 914400L, Cy = 914400L }));
        xfrm.Offset ??= new A.Offset { X = 0L, Y = 0L };
        xfrm.Extents ??= new A.Extents { Cx = 914400L, Cy = 914400L };
        return xfrm;
    }

    protected override void OwnProps(Dictionary<string, string> props)
    {
        var xfrm = picture.ShapeProperties?.Transform2D;
        PptxShape.AddBox(props, xfrm?.Offset?.X?.Value, xfrm?.Offset?.Y?.Value, xfrm?.Extents?.Cx?.Value, xfrm?.Extents?.Cy?.Value);
        var nv = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
        if (nv?.Description?.Value is { Length: > 0 } alt) props["alt"] = alt;
        if (nv?.Id?.Value is { } id) props["id"] = id.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>The bytes of the image part a blip refers to, resolved against the part that owns the blip.</summary>
    internal static (string ContentType, byte[] Data)? Bytes(OpenXmlPart owner, string? embed)
    {
        if (embed is null || !owner.TryGetPartById(embed, out var image)) return null;
        using var stream = image.GetStream(FileMode.Open, FileAccess.Read);
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        return (image.ContentType, ms.ToArray());
    }

    public static (OpenXmlElement Element, string[] Consumed) New(PptxDocument doc, SlidePart slide, IReadOnlyDictionary<string, string> props)
    {
        var src = props.GetValueOrDefault("src") ?? throw new WriterException(ErrorCode.Validation, "An image needs src", "Example: --prop src=chart.png");
        var (relId, info, name) = Embed(slide, src);
        var (width, height) = doc.SlideSize;
        var aspect = info.Height / (double)Math.Max(1, info.Width);
        long w, h;
        if (props.TryGetValue("w", out var pw) && props.TryGetValue("h", out var ph)) (w, h) = (long.Parse(pw, CultureInfo.InvariantCulture), long.Parse(ph, CultureInfo.InvariantCulture));
        else if (props.TryGetValue("w", out pw)) { w = long.Parse(pw, CultureInfo.InvariantCulture); h = (long)Math.Round(w * aspect); }
        else if (props.TryGetValue("h", out ph)) { h = long.Parse(ph, CultureInfo.InvariantCulture); w = (long)Math.Round(h / aspect); }
        else
        {
            w = Math.Min(info.Width * Units.EmuPerPx, width - 2 * 914400L);
            h = (long)Math.Round(w * aspect);
            if (h > height - 2 * 914400L) { h = height - 2 * 914400L; w = (long)Math.Round(h / aspect); }
        }
        var id = PptxDocument.NextShapeId(slide);
        var pic = new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name, Description = props.GetValueOrDefault("alt") is { Length: > 0 } alt ? alt : null },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(new A.Blip { Embed = relId }, new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = PptxShape.Emu(props, "x", 914400L), Y = PptxShape.Emu(props, "y", 914400L) },
                    new A.Extents { Cx = w, Cy = h }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        return (pic, ["src", "x", "y", "w", "h", "alt"]);
    }

    static (string RelId, ImageInfo Info, string Name) Embed(SlidePart slide, string src)
    {
        var (bytes, name) = ImageInfo.Load(src);
        var info = ImageInfo.Read(bytes);
        var part = slide.AddImagePart(info.ContentType);
        part.FeedData(new MemoryStream(bytes));
        return (slide.GetIdOfPart(part), info, name);
    }

    protected override void SetOwnProp(string name, string value)
    {
        switch (name)
        {
            case "x" or "y" or "w" or "h":
                var xfrm = Transform();
                var emu = long.Parse(value, CultureInfo.InvariantCulture);
                switch (name)
                {
                    case "x": xfrm.Offset!.X = emu; break;
                    case "y": xfrm.Offset!.Y = emu; break;
                    case "w": xfrm.Extents!.Cx = emu; break;
                    default: xfrm.Extents!.Cy = emu; break;
                }
                break;
            case "alt":
                if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties is { } nv) nv.Description = value.Length > 0 ? value : null;
                break;
        }
    }

    public override string GetRaw() => picture.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(picture, raw, doc.Namespaces);
    public override void Remove() => picture.Remove();

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent is not PptxSlide target)
            throw new WriterException(ErrorCode.Validation, "Pictures move between slides only", "Use --to /slide[n].");
        picture.Remove();
        target.Insert(picture, index);
    }

    public override Node CopyTo(Node newParent, int? index) => newParent is PptxSlide target
        ? target.Place(PptxCopy.Element(picture, (SlidePart)part, target.Part), index)
        : throw new WriterException(ErrorCode.Validation, "Pictures are copied onto slides", "Use --to /slide[n].");
}

sealed class PptxTable(PptxDocument doc, SlidePart slide, P.GraphicFrame frame) : Node
{
    public override string Kind => "table";
    public override object Anchor => frame;

    A.Table Table => TableOf(frame)!;

    public static A.Table? TableOf(P.GraphicFrame frame) => frame.Graphic?.GraphicData?.GetFirstChild<A.Table>();

    protected override IEnumerable<Node> ProjectChildren() => Table.Elements<A.TableRow>().Select(r => (Node)new PptxRow(doc, slide, r));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var rows = Table.Elements<A.TableRow>().ToList();
        var cols = Table.TableGrid?.Elements<A.GridColumn>().Count() ?? 0;
        if (cols == 0 && rows.Count > 0) cols = rows.Max(r => r.Elements<A.TableCell>().Count());
        var props = new Dictionary<string, string>
        {
            ["rows"] = rows.Count.ToString(CultureInfo.InvariantCulture),
            ["cols"] = cols.ToString(CultureInfo.InvariantCulture),
            ["data"] = NodeJson.Compact(w =>
            {
                w.WriteStartArray();
                foreach (var r in rows) PptxRow.WriteData(w, r);
                w.WriteEndArray();
            }),
        };
        var xfrm = frame.Transform;
        PptxShape.AddBox(props, xfrm?.Offset?.X?.Value, xfrm?.Offset?.Y?.Value, xfrm?.Extents?.Cx?.Value, xfrm?.Extents?.Cy?.Value);
        return props;
    }

    public static (OpenXmlElement Element, string[] Consumed) New(PptxDocument doc, SlidePart slide, IReadOnlyDictionary<string, string> props)
    {
        var data = props.TryGetValue("data", out var json) ? ParseRows(json) : null;
        var rows = props.TryGetValue("rows", out var r) ? int.Parse(r, CultureInfo.InvariantCulture) : data?.Count ?? 0;
        var cols = props.TryGetValue("cols", out var c) ? int.Parse(c, CultureInfo.InvariantCulture) : data?.Max(x => x.Count) ?? 0;
        if (data is not null)
        {
            rows = Math.Max(rows, data.Count);
            cols = Math.Max(cols, data.Count == 0 ? 0 : data.Max(x => x.Count));
        }
        if (rows < 1 || cols < 1)
            throw new WriterException(ErrorCode.Validation, "A table needs rows and cols, or data", "Example: --prop rows=3 --prop cols=4, or --prop data='[[\"a\",\"b\"]]'");
        var (slideWidth, _) = doc.SlideSize;
        var x = PptxShape.Emu(props, "x", 914400L);
        var y = PptxShape.Emu(props, "y", 1371600L);
        var w = PptxShape.Emu(props, "w", slideWidth - 2 * 914400L);
        var rowHeight = 370840L;
        var h = PptxShape.Emu(props, "h", rowHeight * rows);
        var table = new A.Table(
            new A.TableProperties(new A.TableStyleId { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" }) { FirstRow = true, BandRow = true },
            new A.TableGrid(Enumerable.Range(0, cols).Select(_ => (OpenXmlElement)new A.GridColumn { Width = w / cols })));
        for (var i = 0; i < rows; i++)
            table.Append(new A.TableRow(Enumerable.Range(0, cols).Select(_ => (OpenXmlElement)NewCell())) { Height = h / rows });
        var id = PptxDocument.NextShapeId(slide);
        var frame = new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = "Table " + id },
                new P.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(new A.Offset { X = x, Y = y }, new A.Extents { Cx = w, Cy = h }),
            new A.Graphic(new A.GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }));
        if (data is not null) Fill(frame, slide, data);
        return (frame, ["rows", "cols", "data", "x", "y", "w", "h"]);
    }

    internal static A.TableCell NewCell() => new(
        new A.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" })),
        new A.TableCellProperties());

    internal static List<List<string>> ParseRows(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Array || parsed.RootElement.EnumerateArray().Any(x => x.ValueKind != JsonValueKind.Array))
            throw new WriterException(ErrorCode.Validation, "data must be a JSON array of rows", "Example: [[\"Name\",\"Score\"],[\"Ann\",\"90\"]]");
        return parsed.RootElement.EnumerateArray().Select(row => row.EnumerateArray().Select(CellText).ToList()).ToList();
    }

    internal static List<string> ParseCells(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        if (parsed.RootElement.ValueKind != JsonValueKind.Array)
            throw new WriterException(ErrorCode.Validation, "data must be a JSON array", "Example: [\"Ann\",\"90\"]");
        return parsed.RootElement.EnumerateArray().Select(CellText).ToList();
    }

    static string CellText(JsonElement e) => e.ValueKind switch
    {
        JsonValueKind.String => e.GetString()!,
        JsonValueKind.Null => "",
        _ => e.GetRawText(),
    };

    static void Fill(P.GraphicFrame frame, SlidePart slide, List<List<string>> data)
    {
        var table = TableOf(frame)!;
        SetRowCount(frame, Math.Max(1, data.Count));
        var rows = table.Elements<A.TableRow>().ToList();
        for (var i = 0; i < data.Count; i++)
        {
            PptxRow.SetCellCount(rows[i], Math.Max(1, data[i].Count));
            var cells = rows[i].Elements<A.TableCell>().ToList();
            for (var j = 0; j < data[i].Count; j++) PptxText.SetBody(cells[j].TextBody ??= new A.TextBody(new A.BodyProperties(), new A.ListStyle()), slide, [new RunSpec(data[i][j])]);
        }
        SyncGrid(table);
    }

    static void SetRowCount(P.GraphicFrame frame, int count)
    {
        var table = TableOf(frame)!;
        var rows = table.Elements<A.TableRow>().ToList();
        while (rows.Count < count)
        {
            var clone = (A.TableRow)rows[^1].CloneNode(true);
            foreach (var cell in clone.Elements<A.TableCell>()) ClearCell(cell);
            table.Append(clone);
            rows.Add(clone);
        }
        while (rows.Count > count)
        {
            rows[^1].Remove();
            rows.RemoveAt(rows.Count - 1);
        }
        if (frame.Transform?.Extents is { } extents && rows.Count > 0)
            extents.Cy = rows.Sum(r => r.Height?.Value ?? 370840L);
    }

    internal static void ClearCell(A.TableCell cell)
    {
        var body = cell.TextBody ??= new A.TextBody(new A.BodyProperties(), new A.ListStyle());
        foreach (var p in body.Elements<A.Paragraph>().ToList()) p.Remove();
        body.Append(new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" }));
    }

    internal static void SyncGrid(A.Table table)
    {
        var wanted = table.Elements<A.TableRow>().Max(r => r.Elements<A.TableCell>().Count());
        var grid = table.TableGrid ??= new A.TableGrid();
        var columns = grid.Elements<A.GridColumn>().ToList();
        while (columns.Count < wanted)
        {
            var clone = (A.GridColumn?)columns.LastOrDefault()?.CloneNode(true) ?? new A.GridColumn { Width = 1828800L };
            grid.Append(clone);
            columns.Add(clone);
        }
        while (columns.Count > wanted)
        {
            columns[^1].Remove();
            columns.RemoveAt(columns.Count - 1);
        }
    }

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "rows": SetRowCount(frame, int.Parse(value, CultureInfo.InvariantCulture)); break;
            case "cols":
                foreach (var row in Table.Elements<A.TableRow>()) PptxRow.SetCellCount(row, int.Parse(value, CultureInfo.InvariantCulture));
                SyncGrid(Table);
                break;
            case "data": Fill(frame, slide, ParseRows(value)); break;
            case "x" or "y" or "w" or "h":
                var xfrm = frame.Transform ??= new P.Transform(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 914400L, Cy = 914400L });
                var emu = long.Parse(value, CultureInfo.InvariantCulture);
                switch (name)
                {
                    case "x": xfrm.Offset!.X = emu; break;
                    case "y": xfrm.Offset!.Y = emu; break;
                    case "w": xfrm.Extents!.Cx = emu; break;
                    default: xfrm.Extents!.Cy = emu; break;
                }
                break;
        }
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var last = Table.Elements<A.TableRow>().LastOrDefault();
        var row = last is null ? new A.TableRow(NewCell()) { Height = 370840L } : (A.TableRow)last.CloneNode(true);
        foreach (var cell in row.Elements<A.TableCell>()) ClearCell(cell);
        if (!OoxmlTree.InsertBefore(this, Table, row, index)) Table.Append(row);
        var node = new PptxRow(doc, slide, row);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override string GetRaw() => frame.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(frame, raw, doc.Namespaces);
    public override void Remove() => frame.Remove();

    public override void MoveTo(Node newParent, int? index)
    {
        if (newParent is not PptxSlide target)
            throw new WriterException(ErrorCode.Validation, "Tables move between slides only", "Use --to /slide[n].");
        frame.Remove();
        target.Insert(frame, index);
    }

    public override Node CopyTo(Node newParent, int? index) => newParent is PptxSlide target
        ? target.Place(PptxCopy.Element(frame, slide, target.Part), index)
        : throw new WriterException(ErrorCode.Validation, "Tables are copied onto slides", "Use --to /slide[n].");
}

sealed class PptxRow(PptxDocument doc, SlidePart slide, A.TableRow row) : Node
{
    public override string Kind => "row";
    public override object Anchor => row;

    protected override IEnumerable<Node> ProjectChildren() => row.Elements<A.TableCell>().Select(c => (Node)new PptxCell(doc, slide, c));

    public override IReadOnlyDictionary<string, string> GetProps() =>
        new Dictionary<string, string> { ["data"] = NodeJson.Compact(w => WriteData(w, row)) };

    internal static void WriteData(Utf8JsonWriter w, A.TableRow row)
    {
        w.WriteStartArray();
        foreach (var cell in row.Elements<A.TableCell>()) w.WriteStringValue(PptxText.BodyText(cell.TextBody));
        w.WriteEndArray();
    }

    internal static void SetCellCount(A.TableRow row, int count)
    {
        var cells = row.Elements<A.TableCell>().ToList();
        while (cells.Count < count)
        {
            var clone = cells.Count > 0 ? (A.TableCell)cells[^1].CloneNode(true) : PptxTable.NewCell();
            PptxTable.ClearCell(clone);
            if (cells.Count > 0) row.InsertAfter(clone, cells[^1]);
            else row.PrependChild(clone);
            cells.Add(clone);
        }
        while (cells.Count > count)
        {
            cells[^1].Remove();
            cells.RemoveAt(cells.Count - 1);
        }
    }

    public override void SetProp(string name, string value)
    {
        if (name != "data") return;
        var texts = PptxTable.ParseCells(value);
        SetCellCount(row, Math.Max(1, texts.Count));
        var cells = row.Elements<A.TableCell>().ToList();
        for (var i = 0; i < texts.Count; i++)
            PptxText.SetBody(cells[i].TextBody ??= new A.TextBody(new A.BodyProperties(), new A.ListStyle()), slide, [new RunSpec(texts[i])]);
        if (row.Parent is A.Table table) PptxTable.SyncGrid(table);
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var last = row.Elements<A.TableCell>().LastOrDefault();
        var cell = last is null ? PptxTable.NewCell() : (A.TableCell)last.CloneNode(true);
        PptxTable.ClearCell(cell);
        if (!OoxmlTree.InsertBefore(this, row, cell, index))
        {
            if (last is null) row.PrependChild(cell);
            else row.InsertAfter(cell, last);
        }
        if (row.Parent is A.Table table) PptxTable.SyncGrid(table);
        var node = new PptxCell(doc, slide, cell);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override void Remove()
    {
        if (row.Parent is A.Table t && t.Elements<A.TableRow>().Count() == 1)
            throw new WriterException(ErrorCode.Validation, "A table cannot lose its last row", "Remove the table instead.");
        row.Remove();
    }

    public override string GetRaw() => row.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(row, raw, doc.Namespaces);
}

sealed class PptxCell(PptxDocument doc, SlidePart slide, A.TableCell cell) : Node
{
    public override string Kind => "cell";
    public override object Anchor => cell;

    protected override IEnumerable<Node> ProjectChildren() =>
        PptxText.Paragraphs(cell.TextBody).Select(p => (Node)new PptxParagraph(doc, slide, p));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["text"] = PptxText.BodyText(cell.TextBody), ["html"] = string.Join("<br>", Children.Where(c => c.Kind == "paragraph").Select(Exporter.HtmlOf)) };
        if (PptxText.AlignOf(PptxText.Paragraphs(cell.TextBody).FirstOrDefault()?.ParagraphProperties) is { } align) props["align"] = align;
        if (cell.TableCellProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value is { } fill) props["fill"] = fill.ToUpperInvariant();
        return props;
    }

    public override void SetProp(string name, string value)
    {
        var body = cell.TextBody ??= new A.TextBody(new A.BodyProperties(), new A.ListStyle());
        switch (name)
        {
            case "text": PptxText.SetBody(body, slide, [new RunSpec(value)]); break;
            case "md": PptxText.SetBody(body, slide, InlineMarkdown.Parse(value)); break;
            case "html": PptxText.SetBody(body, slide, InlineHtml.Parse(value)); break;
            case "align":
                foreach (var p in body.Elements<A.Paragraph>()) (p.ParagraphProperties ??= new A.ParagraphProperties()).Alignment = PptxText.AlignValue(value);
                break;
            case "fill":
                var tcPr = cell.TableCellProperties ??= new A.TableCellProperties();
                foreach (var f in tcPr.ChildElements.Where(e => e is A.SolidFill or A.NoFill or A.GradientFill or A.PatternFill or A.BlipFill or A.GroupFill).ToList()) f.Remove();
                if (value != "none") PptxText.InsertBeforeAny(tcPr, new A.SolidFill(new A.RgbColorModelHex { Val = value }), e => e is A.Cell3DProperties or A.ExtensionList);
                break;
        }
    }

    public override Node Add(string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var body = cell.TextBody ??= new A.TextBody(new A.BodyProperties(), new A.ListStyle());
        var paragraph = new A.Paragraph();
        if (!OoxmlTree.InsertBefore(this, body, paragraph, index)) body.Append(paragraph);
        var node = new PptxParagraph(doc, slide, paragraph);
        foreach (var (name, value) in props) node.SetProp(name, value);
        return node;
    }

    public override void Remove()
    {
        if (cell.Parent is A.TableRow r && r.Elements<A.TableCell>().Count() == 1)
            throw new WriterException(ErrorCode.Validation, "A row cannot lose its last cell", "Remove the row instead.");
        var table = cell.Parent?.Parent as A.Table;
        cell.Remove();
        if (table is not null) PptxTable.SyncGrid(table);
    }

    public override string GetRaw() => cell.OuterXml;
    public override void SetRaw(string raw) => RawXml.Replace(cell, raw, doc.Namespaces);
}
