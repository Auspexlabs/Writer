using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

public sealed class XlsxAdapter : IFormatAdapter
{
    public string Format => "xlsx";
    public IReadOnlyList<string> Extensions { get; } = [".xlsx"];
    public bool CanWrite => true;

    public Document Create()
    {
        var ms = new MemoryStream();
        using (var package = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbook = package.AddWorkbookPart();
            workbook.Workbook = new Workbook(new Sheets());
            var sheet = workbook.AddNewPart<WorksheetPart>();
            sheet.Worksheet = new Worksheet(new SheetData());
            workbook.Workbook.Sheets!.Append(new Sheet { Id = workbook.GetIdOfPart(sheet), SheetId = 1U, Name = "Sheet1" });
            workbook.AddNewPart<WorkbookStylesPart>().Stylesheet = new Stylesheet(XlsxStyles.MinimalXml);
        }
        ms.Position = 0;
        return Open(ms);
    }

    public Document Open(Stream stream)
    {
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        ms.Position = 0;
        SpreadsheetDocument package;
        try
        {
            package = SpreadsheetDocument.Open(ms, true);
        }
        catch (Exception ex) when (ex is not WriterException)
        {
            throw new WriterException(ErrorCode.FormatError, $"Not a valid .xlsx file: {ex.Message}", "Check that the file opens in Excel.");
        }
        if (package.WorkbookPart?.Workbook is null)
        {
            package.Dispose();
            throw new WriterException(ErrorCode.FormatError, "The .xlsx has no workbook part", "Check that the file opens in Excel.");
        }
        return new XlsxDocument(package);
    }
}

/// <summary>An open Excel package.</summary>
public sealed class XlsxDocument : Document
{
    public SpreadsheetDocument Package { get; }
    public WorkbookPart Workbook => Package.WorkbookPart!;
    internal XlsxStyles Styles { get; }
    internal XlsxStrings Strings { get; }
    internal IEnumerable<KeyValuePair<string, string>> Namespaces => Workbook.Workbook!.NamespaceDeclarations;

    public XlsxDocument(SpreadsheetDocument package)
    {
        Package = package;
        Styles = new XlsxStyles(Workbook);
        Strings = new XlsxStrings(Workbook);
    }

    public override string Format => "xlsx";
    public override Node Root => new XlsxRoot(this);

    public override void Save(Stream stream)
    {
        using var clone = Package.Clone(stream);
    }

    public override void Dispose() => Package.Dispose();

    Sheets SheetList => Workbook.Workbook!.Sheets ??= new Sheets();

    internal List<(Sheet Sheet, WorksheetPart Part)> Sheets =>
        SheetList.Elements<Sheet>()
            .Select(s => (s, s.Id?.Value is { } id && Workbook.TryGetPartById(id, out var part) ? part as WorksheetPart : null))
            .Where(x => x.Item2 is not null).Select(x => (x.s, x.Item2!)).ToList();

    internal WorksheetPart AddSheet(string? name, int? index)
    {
        var sheets = SheetList.Elements<Sheet>().ToList();
        name = name is { Length: > 0 } ? ValidName(name) : NextName(sheets);
        if (sheets.Any(s => string.Equals(s.Name?.Value, name, StringComparison.OrdinalIgnoreCase)))
            throw new WriterException(ErrorCode.Validation, $"A sheet called '{name}' already exists", "Pick another name.");
        var part = Workbook.AddNewPart<WorksheetPart>();
        part.Worksheet = new Worksheet(new SheetData());
        var sheet = new Sheet
        {
            Id = Workbook.GetIdOfPart(part),
            SheetId = (sheets.Count == 0 ? 0u : sheets.Max(s => s.SheetId?.Value ?? 0u)) + 1,
            Name = name,
        };
        if (index is { } i && i >= 1 && i <= sheets.Count) SheetList.InsertBefore(sheet, sheets[i - 1]);
        else SheetList.Append(sheet);
        Reindex(sheets);
        return part;
    }

    static string NextName(List<Sheet> sheets)
    {
        for (var n = sheets.Count + 1; ; n++)
        {
            var candidate = $"Sheet{n}";
            if (!sheets.Any(s => string.Equals(s.Name?.Value, candidate, StringComparison.OrdinalIgnoreCase))) return candidate;
        }
    }

    internal static string ValidName(string name)
    {
        var trimmed = name.Trim();
        if (trimmed.Length is 0 or > 31 || trimmed.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0)
            throw new WriterException(ErrorCode.Validation, $"'{name}' is not a valid sheet name", "Use 1-31 characters without [ ] : * ? / \\.");
        return trimmed;
    }

    internal void RemoveSheet(Sheet sheet, WorksheetPart part)
    {
        var before = SheetList.Elements<Sheet>().ToList();
        if (before.Count == 1)
            throw new WriterException(ErrorCode.Validation, "A workbook needs at least one sheet", "Add another sheet first.");
        var removed = (uint)before.IndexOf(sheet);
        foreach (var name in LocalNames())
            if (name.LocalSheetId!.Value == removed) name.Remove();
        sheet.Remove();
        Workbook.DeletePart(part);
        Reindex(before);
    }

    internal void MoveSheet(Sheet sheet, int? index)
    {
        var before = SheetList.Elements<Sheet>().ToList();
        sheet.Remove();
        var sheets = SheetList.Elements<Sheet>().ToList();
        if (index is { } i && i >= 1 && i <= sheets.Count) SheetList.InsertBefore(sheet, sheets[i - 1]);
        else SheetList.Append(sheet);
        Reindex(before);
    }

    /// <summary>After sheets were inserted, moved or removed: what refers to a sheet by its position follows the sheet. That is the
    /// sheet-scoped defined names (filters, print areas) and the workbook view's active and first tab; an active tab left pointing
    /// at another sheet than the selected one opens in Excel as a group. A position whose sheet is gone keeps its place, within range.</summary>
    void Reindex(List<Sheet> before)
    {
        var after = SheetList.Elements<Sheet>().ToList();
        uint At(uint i) => i >= before.Count ? i : after.IndexOf(before[(int)i]) is var j and >= 0 ? (uint)j : (uint)Math.Min(i, after.Count - 1);
        foreach (var name in LocalNames()) name.LocalSheetId = At(name.LocalSheetId!.Value);
        foreach (var view in Workbook.Workbook!.BookViews?.Elements<WorkbookView>() ?? [])
        {
            if (view.ActiveTab?.Value is { } active) view.ActiveTab = At(active);
            if (view.FirstSheet?.Value is { } first) view.FirstSheet = At(first);
        }
    }

    /// <summary>Defined names scoped to one sheet (filters, print areas); their localSheetId is a position in the sheet list.</summary>
    List<DefinedName> LocalNames() =>
        Workbook.Workbook!.DefinedNames?.Elements<DefinedName>().Where(n => n.LocalSheetId?.Value is not null).ToList() ?? [];

    /// <summary>Asks Excel to recalculate on open, so formulas written here get results.</summary>
    internal void RecalculateOnLoad()
    {
        var calc = Workbook.Workbook!.CalculationProperties ??= new CalculationProperties { CalculationId = 0U };
        calc.FullCalculationOnLoad = true;
    }
}

/// <summary>Shared string table access.</summary>
sealed class XlsxStrings(WorkbookPart workbook)
{
    List<string>? _cache;

    public string Get(int index)
    {
        _cache ??= workbook.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(i => i.InnerText).ToList() ?? [];
        return index >= 0 && index < _cache.Count ? _cache[index] : "";
    }

    public int Add(string text)
    {
        _ = Get(0);
        var part = workbook.SharedStringTablePart ?? workbook.AddNewPart<SharedStringTablePart>();
        var table = part.SharedStringTable ??= new SharedStringTable();
        table.Append(new SharedStringItem(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        _cache!.Add(text);
        var count = (uint)_cache.Count;
        table.Count = count;
        table.UniqueCount = count;
        return _cache.Count - 1;
    }
}
