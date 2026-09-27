using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

static class PptxFields
{
    public static string? Read(P.Shape shape)
    {
        var fields = shape.TextBody?.Descendants<Field>().ToArray() ?? [];
        if (fields.Length == 1 && !(shape.TextBody?.Descendants<Run>().Any(r => r.Text?.Text.Length > 0) ?? false))
            return fields[0].Type?.Value switch { "slidenum" => "slideNumber", { } type when type.StartsWith("datetime", StringComparison.Ordinal) => "date", _ => null };
        return shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.Footer ? "footer" : null;
    }

    public static void Set(P.Shape shape, PptxDocument doc, SlidePart slide, string kind)
    {
        var body = shape.TextBody ??= new P.TextBody(new BodyProperties(), new ListStyle());
        var old = body.Elements<Paragraph>().FirstOrDefault();
        var properties = old?.GetFirstChild<Run>()?.RunProperties ?? old?.GetFirstChild<Field>()?.RunProperties;
        if (kind != "footer")
        {
            foreach (var p in body.Elements<Paragraph>().ToList()) p.Remove();
            var field = new Field { Id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}", Type = kind == "slideNumber" ? "slidenum" : "datetime1" };
            if (properties is not null) field.Append(properties.CloneNode(true));
            field.Append(new Text(kind == "slideNumber" ? (doc.Slides.IndexOf(slide) + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) : DateTime.Today.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)));
            var para = new Paragraph(); if (old?.ParagraphProperties is { } pp) para.Append(pp.CloneNode(true)); para.Append(field); body.Append(para);
        }
        var nv = shape.NonVisualShapeProperties ??= new P.NonVisualShapeProperties();
        var app = nv.ApplicationNonVisualDrawingProperties ??= new P.ApplicationNonVisualDrawingProperties();
        app.PlaceholderShape = new P.PlaceholderShape { Type = kind switch { "slideNumber" => P.PlaceholderValues.SlideNumber, "date" => P.PlaceholderValues.DateAndTime, _ => P.PlaceholderValues.Footer } };
    }
}
