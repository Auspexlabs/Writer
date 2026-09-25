using System.Globalization;
using DocumentFormat.OpenXml;
using Writer.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WPS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>A text box or shape floating in a paragraph: Word's wps:wsp in a wp:anchor (also inside the mc:AlternateContent Word writes
/// around it), placed as a floating picture is (x, y from the column and the paragraph, a wrap), with a preset geometry, a fill, an
/// outline and plain text. Address one as //shape[@id=n] (its wp:docPr id).</summary>
sealed class DocxShape(W.Drawing drawing) : Node
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
    const string WpsUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    public override string Kind => "shape";
    public override object Anchor => drawing;
    OpenXmlCompositeElement Holder => (OpenXmlCompositeElement?)drawing.GetFirstChild<DW.Anchor>() ?? drawing.GetFirstChild<DW.Inline>()!;
    WPS.WordprocessingShape Wsp => drawing.Descendants<WPS.WordprocessingShape>().First();
    WPS.ShapeProperties SpPr => Wsp.GetFirstChild<WPS.ShapeProperties>() ?? Wsp.InsertAfter(new WPS.ShapeProperties(), Wsp.GetFirstChild<WPS.NonVisualDrawingShapeProperties>());

    /// <summary>The shapes in the paragraph's own runs (not inside text boxes), in reading order.</summary>
    public static IEnumerable<Node> In(W.Paragraph p) =>
        p.Descendants<W.Drawing>().Where(d => d.Parent is W.Run or AlternateContentChoice && d.Ancestors<W.Paragraph>().First() == p && d.Descendants<WPS.WordprocessingShape>().Any())
            .Select(d => (Node)new DocxShape(d));

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string>();
        var spPr = Wsp.GetFirstChild<WPS.ShapeProperties>();
        props["geometry"] = spPr?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText ?? (spPr?.GetFirstChild<A.CustomGeometry>() is null ? "rect" : "custom");
        if (Holder.GetFirstChild<DW.Extent>() is { } extent)
        {
            props["width"] = (extent.Cx?.Value ?? 0).ToString(Inv);
            props["height"] = (extent.Cy?.Value ?? 0).ToString(Inv);
        }
        var style = Wsp.GetFirstChild<WPS.ShapeStyle>();
        if ((FillOf(spPr) ?? (style?.FillReference is { } fr && fr.Index?.Value > 0 ? ColorOf(fr) : null)) is { } fill) props["fill"] = fill;
        if (((spPr?.GetFirstChild<A.Outline>() is { } ln ? FillOf(ln) : null) ?? (style?.LineReference is { } lr && lr.Index?.Value > 0 ? ColorOf(lr) : null)) is { } line) props["line"] = line;
        props["text"] = BoxText();
        if (Holder.GetFirstChild<DW.DocProperties>()?.Id?.Value is { } id) props["id"] = id.ToString(Inv);
        if (drawing.GetFirstChild<DW.Anchor>() is { } a) DocxImage.ReadPlace(a, props);
        return props;
    }

    /// <summary>A solid fill's colour, none for no fill; null when the shape leaves it to its style.</summary>
    static string? FillOf(OpenXmlElement? e) =>
        e?.GetFirstChild<A.NoFill>() is not null ? "none" : e?.GetFirstChild<A.SolidFill>() is { } solid ? ColorOf(solid) : e?.GetFirstChild<A.GradientFill>()?.Descendants<A.GradientStop>().FirstOrDefault() is { } stop ? ColorOf(stop) : null;

    /// <summary>A colour as hex: its own RGB, Word's preset black and white, or the default Office theme's colour for a scheme name.</summary>
    static string? ColorOf(OpenXmlElement e) =>
        e.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value?.ToUpperInvariant()
        ?? (e.GetFirstChild<A.PresetColor>()?.Val?.InnerText switch { "black" => "000000", "white" => "FFFFFF", _ => null })
        ?? (e.GetFirstChild<A.SchemeColor>()?.Val?.InnerText switch
        {
            "accent1" => "4472C4", "accent2" => "ED7D31", "accent3" => "A5A5A5", "accent4" => "FFC000", "accent5" => "5B9BD5", "accent6" => "70AD47",
            "bg1" or "lt1" => "FFFFFF", "tx1" or "dk1" => "000000", "bg2" or "lt2" => "E7E6E6", "tx2" or "dk2" => "44546A", _ => null,
        });

    string BoxText() => Wsp.GetFirstChild<WPS.TextBoxInfo2>()?.TextBoxContent is { } content
        ? string.Join("\n", content.Elements<W.Paragraph>().Select(DocxRuns.ParagraphText)) : "";

    public override string GetRaw() => drawing.OuterXml;

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "geometry":
                if (value is "custom" or "textbox") value = "rect";
                SpPr.RemoveAllChildren<A.CustomGeometry>();
                if (SpPr.GetFirstChild<A.PresetGeometry>() is { } geometry) geometry.Preset = new A.ShapeTypeValues(value);
                else SpPr.InsertAfter(new A.PresetGeometry(new A.AdjustValueList()) { Preset = new A.ShapeTypeValues(value) }, SpPr.GetFirstChild<A.Transform2D>());
                break;
            case "width" or "height":
                var emu = long.Parse(value, Inv);
                var extent = Holder.GetFirstChild<DW.Extent>()!;
                if (name == "width") extent.Cx = emu; else extent.Cy = emu;
                if (SpPr.GetFirstChild<A.Transform2D>()?.Extents is { } ext) { if (name == "width") ext.Cx = emu; else ext.Cy = emu; }
                break;
            case "fill":
                foreach (var old in SpPr.ChildElements.Where(e => e is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill).ToList()) old.Remove();
                SpPr.InsertAfter(Fill(value), (OpenXmlElement?)SpPr.GetFirstChild<A.PresetGeometry>() ?? SpPr.GetFirstChild<A.CustomGeometry>() ?? (OpenXmlElement?)SpPr.GetFirstChild<A.Transform2D>());
                break;
            case "line":
                var ln = SpPr.GetFirstChild<A.Outline>() ?? SpPr.InsertAfter(new A.Outline { Width = 9525 },
                    SpPr.ChildElements.LastOrDefault(e => e is A.Transform2D or A.PresetGeometry or A.CustomGeometry or A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill));
                foreach (var old in ln.ChildElements.Where(e => e is A.NoFill or A.SolidFill or A.GradientFill or A.PatternFill).ToList()) old.Remove();
                ln.PrependChild(Fill(value));
                break;
            case "text":
                SetText(value);
                break;
            case "wrap" or "x" or "y" or "xFrom" or "yFrom" or "xAlign" or "yAlign" when drawing.GetFirstChild<DW.Anchor>() is { } a:
                DocxImage.SetPlace(a, name, value == "inline" ? "square" : value);
                break;
        }
    }

    static OpenXmlElement Fill(string value) => value == "none" ? new A.NoFill() : new A.SolidFill(new A.RgbColorModelHex { Val = value.ToUpperInvariant() });

    /// <summary>The text box's paragraphs anew, one per line.</summary>
    void SetText(string text)
    {
        var wsp = Wsp;
        var box = wsp.GetFirstChild<WPS.TextBoxInfo2>();
        if (text.Length == 0) { box?.Remove(); return; }
        if (box is null)
        {
            box = new WPS.TextBoxInfo2();
            var body = wsp.GetFirstChild<WPS.TextBodyProperties>();
            if (body is null) wsp.Append(box); else body.InsertBeforeSelf(box);
        }
        box.TextBoxContent = new W.TextBoxContent(text.ReplaceLineEndings("\n").Split('\n').Select(line => line.Length == 0 ? new W.Paragraph() : new W.Paragraph(new W.Run(DocxRuns.TextElements(line)))));
    }

    public override void Remove()
    {
        OpenXmlElement own = drawing.Parent is AlternateContentChoice choice && choice.Parent is AlternateContent alt ? alt : drawing;
        var run = own.Parent as W.Run;
        own.Remove();
        if (run is not null && !run.ChildElements.Any(e => e is not W.RunProperties))
        {
            var (previous, next) = (run.PreviousSibling(), run.NextSibling());
            DocxImage.Cut(run);
            DocxRuns.Rejoin(previous, next);
        }
    }

    /// <summary>A new shape floating in the paragraph: at the top left of the column where the paragraph starts unless props place it.</summary>
    public static Node Add(DocxDocument doc, W.Paragraph p, IReadOnlyDictionary<string, string> props)
    {
        var id = DocxImage.NextId(doc);
        long Emu(string name, long fallback) => props.TryGetValue(name, out var v) ? long.Parse(v, Inv) : fallback;
        var (cx, cy) = (Emu("width", 1440000L), Emu("height", 720000L)); // 4cm × 2cm
        var geometry = props.GetValueOrDefault("geometry") ?? "rect";
        var textBox = geometry is "textbox" || (geometry == "rect" && props.ContainsKey("text"));
        if (geometry is "textbox" or "custom") geometry = "rect";
        var wsp = new WPS.WordprocessingShape(
            new WPS.NonVisualDrawingShapeProperties { TextBox = textBox ? true : null },
            new WPS.ShapeProperties(
                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = new A.ShapeTypeValues(geometry) },
                Fill(props.GetValueOrDefault("fill") ?? (textBox ? "FFFFFF" : "4472C4")),
                new A.Outline(Fill(props.GetValueOrDefault("line") ?? (textBox ? "000000" : "2F528F"))) { Width = textBox ? 6350 : 12700 }),
            new WPS.TextBodyProperties(new A.NoAutoFit())
            {
                Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square,
                LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = textBox ? A.TextAnchoringTypeValues.Top : A.TextAnchoringTypeValues.Center,
            });
        var anchor = new DW.Anchor(
            new DW.SimplePosition { X = 0L, Y = 0L },
            new DW.HorizontalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.HorizontalRelativePositionValues.Column },
            new DW.VerticalPosition(new DW.PositionOffset("0")) { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph },
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides },
            new DW.DocProperties { Id = id, Name = (textBox ? "Text Box " : "Shape ") + id.ToString(Inv) },
            new DW.NonVisualGraphicFrameDrawingProperties(),
            new A.Graphic(new A.GraphicData(wsp) { Uri = WpsUri }))
        {
            DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 114300U, DistanceFromRight = 114300U, SimplePos = false,
            RelativeHeight = DocxImage.NextZ(doc), BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true,
        };
        var drawing = new W.Drawing(anchor);
        var run = new W.Run(drawing);
        // at the paragraph's start, after the shapes already floating there, where Word anchors what floats in it
        var first = p.ChildElements.FirstOrDefault(c => c is not (W.ParagraphProperties or W.BookmarkStart) && !(c is W.Run r && r.ChildElements.All(e => e is W.RunProperties or W.Drawing or AlternateContent)));
        if (first is null) p.Append(run); else first.InsertBeforeSelf(run);
        var node = new DocxShape(drawing);
        if (props.GetValueOrDefault("text") is { Length: > 0 } text) node.SetProp("text", text);
        foreach (var (name, value) in props)
            if (name is "wrap" or "x" or "y" or "xFrom" or "yFrom" or "xAlign" or "yAlign") node.SetProp(name, value);
        return node;
    }
}
