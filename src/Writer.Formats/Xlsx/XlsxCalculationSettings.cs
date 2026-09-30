using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Writer.Formats.Xlsx;

static class XlsxCalculationSettings
{
    internal static JsonObject Read(XlsxDocument doc)
    {
        var settings = doc.Workbook.Workbook!.CalculationProperties;
        return new JsonObject { ["enabled"] = settings?.Iterate?.Value ?? false,
            ["count"] = settings?.IterateCount?.Value ?? 100U, ["delta"] = settings?.IterateDelta?.Value ?? .001 };
    }

    internal static void Write(XlsxDocument doc, string json)
    {
        var value = JsonNode.Parse(json) as JsonObject ?? new JsonObject();
        var enabled = value["enabled"]?.GetValue<bool>() ?? false;
        var count = value["count"]?.GetValue<int>() ?? 100;
        var delta = value["delta"]?.GetValue<double>() ?? .001;
        if (count is < 1 or > 10000 || !double.IsFinite(delta) || delta < 0)
            throw new ArgumentException("Iteration count must be 1–10000 and maximum change must be a finite nonnegative number.");
        var settings = doc.Workbook.Workbook!.CalculationProperties ??= new CalculationProperties();
        settings.Iterate = enabled; settings.IterateCount = (uint)count; settings.IterateDelta = delta;
        doc.RecalculateOnLoad();
    }
}
