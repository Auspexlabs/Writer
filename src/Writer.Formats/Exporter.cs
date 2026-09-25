using System.Globalization;
using System.Text;
using Writer.Formats.Common;
using Writer.Core;
using Writer.Formats.Markdown;

namespace Writer.Formats;

/// <summary>Cross-format export: replays the source's unified tree into a fresh target document.
/// Inline formatting travels as markdown; containers the target lacks are flattened; leaf kinds it lacks become warnings.
/// Decks are built from headings (one slide each), workbooks from tables (one sheet each), mind maps from headings and list items;
/// a mind map exports as a markdown outline first.</summary>
public static class Exporter
{
    static readonly Dictionary<string, string> Aliases = new() { ["text"] = "paragraph" };

    public static (Document Target, List<string> Warnings) Export(Document source, IFormatAdapter targetAdapter, string? targetPath)
    {
        if (!targetAdapter.CanWrite)
            throw new WriterException(ErrorCode.FormatReadonly, $"{targetAdapter.Format} cannot be written", "Export to md, docx, pptx, xlsx, mm, html or json.");
        if (source is MindMap.MmDocument { FromXmind: true } && targetAdapter.Format == "mm") return (source, []); // an .xmind opens as this map already
        // a compatibility format already is a Word, Excel or PowerPoint document in memory: hand it over as it is
        if (source is Compat.CompatDocument compat && compat.Inner.Format == targetAdapter.Format)
        {
            compat.Inner.SourcePath = targetPath;
            return (compat.Inner, compat.Warnings.ToList());
        }
        if (source.Format == "mm" && targetAdapter.Format != "mm") return FromMindMap(source, targetAdapter, targetPath);
        var warnings = new List<string>();
        var target = targetAdapter.Create();
        target.SourcePath = targetPath;
        var images = new ImageSink(targetPath, target.Format);
        var blocks = TopBlocks(source);
        switch (target.Format)
        {
            case "pptx":
                Deck(source, blocks, target, images, warnings);
                break;
            case "xlsx":
                Workbook(source, target, warnings);
                break;
            case "mm":
                MindMap(source, blocks, target, warnings);
                break;
            default:
                var body = target.Root.Children.FirstOrDefault(c => c.Kind == "body")
                    ?? throw new WriterException(ErrorCode.Validation, $"Export into {target.Format} is not supported yet", "Targets: md, docx, pptx, xlsx, mm, html, json.");
                Copy(blocks, target, body, images, warnings);
                // a table of contents lists the headings, which arrive after it: build it again now they are all there
                foreach (var toc in PathResolver.Query(target.Root, "//toc"))
                    Mutations.Set(toc, new Dictionary<string, string> { ["levels"] = toc.GetProps()["levels"] });
                break;
        }
        return (target, warnings);
    }

    static IReadOnlyList<Node> TopBlocks(Document source) =>
        (source.Root.Children.FirstOrDefault(c => c.Kind == "body") ?? source.Root).Children;

    static bool Supports(string kind, Document target, Node parent) =>
        Registry.Find(kind) is { } k && k.Formats.Contains(target.Format) && k.Parents.Contains(parent.Kind);

    static void Copy(IReadOnlyList<Node> blocks, Document target, Node targetParent, ImageSink images, List<string> warnings)
    {
        foreach (var block in blocks)
        {
            var kind = Aliases.GetValueOrDefault(block.Kind, block.Kind);
            if (!Supports(kind, target, targetParent))
            {
                Flatten(block, target, targetParent, images, warnings);
                continue;
            }
            var props = Props(block, kind, target);
            if (kind == "image")
            {
                var src = images.Store(block, warnings);
                if (src is null) continue;
                props["src"] = src;
            }
            try
            {
                Mutations.Add(targetParent, kind, props, null);
            }
            catch (WriterException ex)
            {
                warnings.Add($"{block.Path}: {ex.Message}; skipped");
            }
        }
    }

    /// <summary>The block's properties the target kind accepts, with inline formatting carried as markdown.</summary>
    static Dictionary<string, string> Props(Node block, string kind, Document target)
    {
        var definition = Registry.Find(kind)!;
        var source = Registry.ToDisplay(block.Format, block.Kind, block.GetProps());
        var props = new Dictionary<string, string>();
        foreach (var p in Registry.PropsFor(definition, target.Format))
            if (!p.ReadOnly && !p.WriteOnly && p.Name != "text" && source.TryGetValue(p.Name, out var value)) props[p.Name] = value;
        var runs = block.Children.Where(c => c.Kind == "run").Select(RunSpecOf).Where(r => !r.Deleted).ToList();
        if (runs.Count > 0 && Registry.PropsFor(definition, target.Format).Any(p => p.Name == "md")) props["md"] = Plain(target, MdWriter.Inline(runs));
        else if (source.TryGetValue("text", out var text) && Registry.PropsFor(definition, target.Format).Any(p => p.Name == "text" && !p.ReadOnly))
            props["text"] = Plain(target, text); // a toc's text is generated, not written
        if (kind == "paragraph" && !props.ContainsKey("list")) props.Remove("level");
        return props;
    }

    /// <summary>Containers the target lacks (slides, shapes, pages, sheets) contribute their content instead.</summary>
    static void Flatten(Node block, Document target, Node targetParent, ImageSink images, List<string> warnings)
    {
        if (block.Kind == "decor") return;
        var props = block.GetProps();
        var inner = block.Children.Where(c => Registry.Find(c.Kind)?.Inline != true).ToList();
        if (block.Kind == "sheet" && Supports("table", target, targetParent))
        {
            if (Supports("heading", target, targetParent) && props.TryGetValue("name", out var name))
                Mutations.Add(targetParent, "heading", new Dictionary<string, string> { ["text"] = name, ["level"] = "2" }, null);
            var rows = inner.Select(r => r.GetProps().GetValueOrDefault("data") ?? "[]").ToList();
            if (rows.Count > 0)
                Mutations.Add(targetParent, "table", new Dictionary<string, string> { ["data"] = "[" + string.Join(",", rows) + "]" }, null);
            return;
        }
        if (block.Kind == "slide" && props.TryGetValue("title", out var title) && title.Length > 0 && Supports("heading", target, targetParent))
        {
            Mutations.Add(targetParent, "heading", new Dictionary<string, string> { ["text"] = title, ["level"] = "2" }, null);
            inner = inner.Where(c => c.GetProps().GetValueOrDefault("placeholder") != "title").ToList();
        }
        if (inner.Count > 0)
        {
            Copy(inner, target, targetParent, images, warnings);
            return;
        }
        if (block.Kind is not ("slide" or "page") && block.Text is { Length: > 0 } text && Supports("paragraph", target, targetParent))
        {
            Mutations.Add(targetParent, "paragraph", new Dictionary<string, string> { ["text"] = Plain(target, text) }, null);
            return;
        }
        if (block.Kind is not ("slide" or "page" or "body"))
            warnings.Add($"{block.Path}: {block.Kind} has no equivalent in {target.Format}; skipped");
    }

    /// <summary>Headings start slides; paragraphs fill the body placeholder; tables and pictures get slides of their own.</summary>
    static void Deck(Document source, IReadOnlyList<Node> blocks, Document target, ImageSink images, List<string> warnings)
    {
        var fallbackTitle = source.Root.GetProps().GetValueOrDefault("title") ?? "Untitled";
        Node? slide = null;
        Node? body = null;
        var lastTitle = fallbackTitle;

        void NewSlide(string title)
        {
            lastTitle = title;
            slide = Mutations.Add(target.Root, "slide", new Dictionary<string, string> { ["layout"] = "Content", ["title"] = title }, null);
            body = slide.Children.FirstOrDefault(c => c.GetProps().GetValueOrDefault("placeholder") == "body");
        }

        void Standalone(Node block, string kind)
        {
            var bodyIsEmpty = body is not null && body.Children.All(p => string.IsNullOrEmpty(p.Text));
            if (slide is null || !bodyIsEmpty) NewSlide(lastTitle);
            body?.Remove();
            body = null;
            var props = Props(block, kind, target);
            if (kind == "image")
            {
                var src = images.Store(block, warnings);
                if (src is null) return;
                props["src"] = src;
            }
            Mutations.Add(slide!, kind, props, null);
        }

        foreach (var block in Walk(blocks))
        {
            switch (block.Kind)
            {
                case "heading" when int.Parse(block.GetProps().GetValueOrDefault("level") ?? "1") <= 2:
                    NewSlide(Plain(target, block.Text ?? ""));
                    break;
                case "heading" or "paragraph" or "code" or "text":
                    if (block.Text is not { Length: > 0 }) break;
                    if (slide is null) NewSlide(fallbackTitle);
                    if (body is null)
                    {
                        NewSlide(lastTitle);
                        if (body is null) break;
                    }
                    var props = Props(block, "paragraph", target);
                    if (block.Kind == "heading") props["md"] = "**" + MdWriter.Escape(Plain(target, block.Text)) + "**";
                    var existing = body.Children;
                    if (existing.Count == 1 && string.IsNullOrEmpty(existing[0].Text)) Mutations.Set(existing[0], props);
                    else Mutations.Add(body, "paragraph", props, null);
                    break;
                case "table":
                    Standalone(block, "table");
                    break;
                case "image":
                    Standalone(block, "image");
                    break;
            }
        }
        if (slide is null) warnings.Add("Nothing to put on slides: the source has no headings, paragraphs, tables or pictures");
    }

    /// <summary>Blocks in reading order, descending into containers such as slides, shapes and pages.</summary>
    static IEnumerable<Node> Walk(IReadOnlyList<Node> blocks)
    {
        foreach (var block in blocks)
        {
            if (block.Kind is "heading" or "paragraph" or "code" or "text" or "table" or "image")
            {
                yield return block;
                continue;
            }
            if (block.Kind == "slide" && block.GetProps().TryGetValue("title", out var title) && title.Length > 0)
                yield return new TitleNode(title);
            var inner = block.Children.Where(c => Registry.Find(c.Kind)?.Inline != true && c.GetProps().GetValueOrDefault("placeholder") != "title").ToList();
            foreach (var n in Walk(inner)) yield return n;
        }
    }

    /// <summary>A slide title presented as a level-2 heading so decks re-export slide by slide.</summary>
    sealed class TitleNode(string title) : Node
    {
        public override string Kind => "heading";
        public override object Anchor => this;
        public override string Format => "pptx";
        public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string> { ["text"] = title, ["level"] = "2" };
    }

    /// <summary>Every table in the source becomes a sheet.</summary>
    static void Workbook(Document source, Document target, List<string> warnings)
    {
        var tables = PathResolver.Query(source.Root, "//table");
        if (tables.Count == 0)
        {
            warnings.Add("No tables found; the workbook is empty");
            return;
        }
        for (var i = 0; i < tables.Count; i++)
        {
            var name = $"Table{i + 1}";
            var sheet = i == 0 ? target.Root.Children[0] : Mutations.Add(target.Root, "sheet", new Dictionary<string, string> { ["name"] = name }, null);
            if (i == 0) Mutations.Set(sheet, new Dictionary<string, string> { ["name"] = name });
            foreach (var row in tables[i].Children)
                if (row.GetProps().TryGetValue("data", out var data))
                    Mutations.Add(sheet, "row", new Dictionary<string, string> { ["data"] = data }, null);
        }
    }

    /// <summary>A mind map as a markdown outline (root heading, nested bullets), which then travels on to any other target.</summary>
    static (Document Target, List<string> Warnings) FromMindMap(Document source, IFormatAdapter targetAdapter, string? targetPath)
    {
        var warnings = new List<string>();
        var sb = new StringBuilder();
        var notes = 0;
        foreach (var root in source.Root.Children) Write(root, -1);
        if (notes > 0) warnings.Add($"{notes} note(s) dropped: markdown has no equivalent for topic notes");
        var md = new MdAdapter().Open(new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString())));
        if (targetAdapter.Format == "md")
        {
            md.SourcePath = targetPath;
            return (md, warnings);
        }
        using (md)
        {
            var (target, more) = Export(md, targetAdapter, targetPath);
            warnings.AddRange(more);
            return (target, warnings);
        }

        void Write(Node topic, int depth)
        {
            if (topic.GetProps().ContainsKey("note")) notes++;
            var text = MdWriter.Inline([new RunSpec((topic.Text ?? "").ReplaceLineEndings(" "))]);
            if (depth < 0) sb.Append("# ").Append(text).Append("\n\n");
            else sb.Append(' ', depth * 2).Append("- ").Append(text).Append('\n');
            foreach (var child in topic.Children) Write(child, depth + 1);
        }
    }

    /// <summary>Headings and list items become topics: a level-L heading sits at depth L, a list item under its heading at its list level.
    /// The centre is the title, else the single level-1 heading, else "Mind map". A source with no structure at all contributes its paragraphs.</summary>
    static void MindMap(Document source, IReadOnlyList<Node> blocks, Document target, List<string> warnings)
    {
        if (source.Format == "mm")
            throw new WriterException(ErrorCode.Validation, "Copy a mind map with create --from instead of export", "Example: writer create copy.mm --from map.mm");
        var all = Walk(blocks).ToList();
        var title = source.Root.GetProps().GetValueOrDefault("title");
        var consumed = title is { Length: > 0 } ? null : all.Where(b => b.Kind == "heading" && Level(b) == 1).ToList() is [var only] ? only : null;
        var rootText = title is { Length: > 0 } ? title : consumed?.Text is { Length: > 0 } t ? t : "Mind map";
        var root = Mutations.Set(target.Root.Children[0], new Dictionary<string, string> { ["text"] = rootText });
        var shift = consumed is null ? 0 : 1;
        var structured = all.Any(b => !ReferenceEquals(b, consumed) && (b.Kind == "heading" || b.GetProps().ContainsKey("list")));
        var stack = new List<Node> { root };
        var headingDepth = 0;
        var skipped = 0;
        foreach (var block in all)
        {
            if (ReferenceEquals(block, consumed)) continue;
            var props = block.GetProps();
            int depth;
            switch (block.Kind)
            {
                case "heading":
                    depth = Math.Max(1, Level(block) - shift);
                    break;
                case "paragraph" or "text" when props.ContainsKey("list"):
                    depth = headingDepth + int.Parse(props.GetValueOrDefault("level") ?? "0", CultureInfo.InvariantCulture) + 1;
                    break;
                case "paragraph" or "text" when !structured:
                    depth = 1;
                    break;
                case "paragraph" or "text":
                    skipped++;
                    continue;
                default:
                    warnings.Add($"{block.Path}: {block.Kind} has no equivalent in mm; skipped");
                    continue;
            }
            while (stack.Count > depth) stack.RemoveAt(stack.Count - 1);
            var text = Plain(target, block.Text ?? "");
            stack.Add(Mutations.Add(stack[^1], "topic", new Dictionary<string, string> { ["text"] = structured ? text : NodeJson.Preview(text, 80) }, null));
            if (block.Kind == "heading") headingDepth = stack.Count - 1;
        }
        if (skipped > 0) warnings.Add($"{skipped} paragraph(s) skipped: only headings and list items become topics");
    }

    static int Level(Node heading) => int.Parse(heading.GetProps().GetValueOrDefault("level") ?? "1", CultureInfo.InvariantCulture);

    /// <summary>Text for the target: Word's page break (a form feed) is a line break in every other format, whose XML cannot hold it.</summary>
    static string Plain(Document target, string text) => target.Format == "docx" ? text : text.Replace('\f', '\n');

    public static RunSpec RunSpecOf(Node run) => RunSpec.FromProps(run.GetProps());

    /// <summary>The inline HTML of a block's runs, for the html property.</summary>
    public static string HtmlOf(Node block) => InlineHtml.Render(block.Children.Where(c => c.Kind == "run").Select(RunSpecOf));

    /// <summary>Writes exported images next to the target file, or to a temp folder for formats that embed them.</summary>
    sealed class ImageSink(string? targetPath, string targetFormat)
    {
        int _n;

        public string? Store(Node image, List<string> warnings)
        {
            if (image.GetBinary() is not { } binary)
            {
                var src = image.GetProps().GetValueOrDefault("src") ?? "";
                if (targetFormat == "md" && src.Length > 0 && !src.StartsWith('/')) return src;
                warnings.Add($"{image.Path}: image data not available; skipped");
                return null;
            }
            var ext = binary.ContentType switch { "image/png" => "png", "image/jpeg" => "jpg", "image/gif" => "gif", "image/bmp" => "bmp", _ => "bin" };
            _n++;
            if (targetFormat == "md" && targetPath is not null)
            {
                var stem = System.IO.Path.GetFileNameWithoutExtension(targetPath);
                var dir = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(targetPath))!, stem + "_files");
                Directory.CreateDirectory(dir);
                var name = $"image{_n}.{ext}";
                File.WriteAllBytes(System.IO.Path.Combine(dir, name), binary.Data);
                return $"{stem}_files/{name}";
            }
            var temp = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"writer-export-{Guid.NewGuid():N}.{ext}");
            File.WriteAllBytes(temp, binary.Data);
            return temp;
        }
    }
}
