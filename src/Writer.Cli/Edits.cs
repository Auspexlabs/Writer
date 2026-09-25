using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Writer.Core;
using Writer.Formats;
using Writer.Formats.Markdown;

namespace Writer.Cli;

/// <summary>The verbs beyond get/set/add/remove: a compact structure view, text search, whole-section edits, find and replace,
/// formulas with a check, and atomic batches. Each is one command line, so the assistant, the MCP tool and the terminal share them,
/// and each is built on the public node tree, so every format that has headings, cells or text gets it.</summary>
static class Edits
{
    /// <summary>`view <file> structure`: the document's shape on a few lines — headings with their paths and what follows each,
    /// tables by size, a workbook's sheets with their used range, header row and a sample row — where the outline lists every block.</summary>
    public static string Structure(Document doc)
    {
        var root = doc.Root;
        if (doc.Format == "xlsx") return Workbook(root);
        var body = root.Children.FirstOrDefault(c => c.Kind == "body");
        if (body is null) return Views.Outline(root);
        var sb = new StringBuilder();
        var blocks = body.Children;
        sb.Append(body.Path).Append("  ").AppendJoin("  ", blocks.GroupBy(b => b.Kind).Select(g => $"{g.Key}s={g.Count()}"));
        var styles = blocks.Select(b => b.GetProps().GetValueOrDefault("style")).Where(s => s is not null).GroupBy(s => s!).Select(g => $"{g.Key}({g.Count()})").ToList();
        if (styles.Count > 0) sb.Append("  styles=").AppendJoin(",", styles);
        sb.Append('\n');
        var run = new List<Node>();
        foreach (var block in blocks)
        {
            if (block.Kind != "heading")
            {
                run.Add(block);
                continue;
            }
            Flush(sb, run);
            var props = block.GetProps();
            sb.Append(block.Path).Append("  level=").Append(props.GetValueOrDefault("level")).Append("  \"").Append(NodeJson.Preview(props.GetValueOrDefault("text") ?? "", 80)).Append("\"\n");
        }
        Flush(sb, run);
        return sb.ToString();
    }

    /// <summary>The blocks between two headings: plain paragraphs as one count with the first and last path, anything else on its own line.</summary>
    static void Flush(StringBuilder sb, List<Node> run)
    {
        if (run.Count == 0) return;
        var plain = run.Where(b => b.Kind == "paragraph").ToList();
        if (plain.Count > 0)
        {
            sb.Append("  ").Append(plain.Count).Append(plain.Count == 1 ? " paragraph  " : " paragraphs  ").Append(plain[0].Path);
            if (plain.Count > 1) sb.Append(" … ").Append(plain[^1].Path);
            sb.Append("  \"").Append(NodeJson.Preview(plain[0].Text ?? "", 50)).Append("\"\n");
        }
        foreach (var block in run.Where(b => b.Kind != "paragraph"))
        {
            sb.Append("  ").Append(block.Path);
            var props = block.GetProps();
            if (block.Kind == "table")
            {
                sb.Append("  ").Append(props.GetValueOrDefault("rows")).Append(" rows × ").Append(props.GetValueOrDefault("cols")).Append(" cols");
                if (block.Children.FirstOrDefault(r => r.Kind == "row") is { } first)
                    sb.Append("  \"").Append(NodeJson.Preview(string.Join(" | ", first.Children.Select(c => c.Text ?? "")), 80)).Append('"');
            }
            else if (props.TryGetValue("text", out var text) && text.Length > 0) sb.Append("  \"").Append(NodeJson.Preview(text, 50)).Append('"');
            sb.Append('\n');
        }
        run.Clear();
    }

    static string Workbook(Node root)
    {
        var sb = new StringBuilder();
        foreach (var sheet in root.Children.Where(c => c.Kind == "sheet"))
        {
            var props = sheet.GetProps();
            sb.Append(sheet.Path).Append("  name=").Append(props.GetValueOrDefault("name"));
            var range = props.GetValueOrDefault("range");
            if (range is not null)
            {
                sb.Append("  range=").Append(range);
                var m = Regex.Match(range, @"^([A-Z]+)(\d+):([A-Z]+)(\d+)$");
                if (m.Success) sb.Append("  (").Append(int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture) - int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) + 1).Append(" rows × ").Append(Column(m.Groups[3].Value) - Column(m.Groups[1].Value) + 1).Append(" cols)");
            }
            else sb.Append("  (empty)");
            foreach (var name in new[] { "filter", "freeze", "merges" })
                if (props.TryGetValue(name, out var v) && v.Length <= 60) sb.Append("  ").Append(name).Append('=').Append(v);
            sb.Append('\n');
            var rows = sheet.Children.Where(c => c.Kind == "row").ToList();
            if (rows.Count > 0)
            {
                var header = Values(rows[0].GetProps().GetValueOrDefault("data") ?? "[]");
                sb.Append("  header (row ").Append(rows[0].Key).Append("): ").Append(NodeJson.Preview(string.Join(" | ", header.Select((h, i) => Letter(i + 1) + " " + h)), 300)).Append('\n');
                if (rows.Count > 1)
                    sb.Append("  row ").Append(rows[1].Key).Append(": ").Append(NodeJson.Preview(string.Join(" | ", Values(rows[1].GetProps().GetValueOrDefault("data") ?? "[]")), 200)).Append('\n');
            }
            foreach (var other in sheet.Children.Where(c => c.Kind != "row"))
            {
                sb.Append("  ").Append(other.Path);
                foreach (var (name, value) in Registry.ToDisplay(other.Format, other.Kind, other.GetProps()))
                    if (name is "type" or "title" or "series" or "categories" or "alt" && value.Length <= 80) sb.Append("  ").Append(name).Append('=').Append(value);
                sb.Append('\n');
            }
        }
        return sb.ToString();
    }

    static List<string> Values(string json)
    {
        using var parsed = JsonDocument.Parse(json);
        return parsed.RootElement.ValueKind == JsonValueKind.Array ? parsed.RootElement.EnumerateArray().Select(e => e.ValueKind == JsonValueKind.String ? e.GetString() ?? "" : e.GetRawText()).ToList() : [];
    }

    static int Column(string letters) => letters.Aggregate(0, (n, ch) => n * 26 + (char.ToUpperInvariant(ch) - 'A' + 1));

    static string Letter(int index)
    {
        var s = "";
        for (var n = index; n > 0; n = (n - 1) / 26) s = (char)('A' + (n - 1) % 26) + s;
        return s;
    }

    /// <summary>`search <file> <text>`: every block whose text contains the words, one line each with its path and the text around the match.</summary>
    public static string Search(Args a)
    {
        using var doc = Files.Open(a.Need(0, "file"));
        var needle = a.Need(1, "text");
        var comparison = a.Flag("ignore-case") ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var sb = new StringBuilder();
        var total = 0;
        foreach (var (node, text) in TextNodes(doc.Root, doc.Format))
        {
            var at = text.IndexOf(needle, comparison);
            if (at < 0) continue;
            total++;
            if (total > 50) continue;
            var from = Math.Max(0, at - 30);
            var to = Math.Min(text.Length, at + needle.Length + 30);
            sb.Append(node.Path).Append("  \"").Append(from > 0 ? "…" : "").Append(text[from..to].ReplaceLineEndings(" ")).Append(to < text.Length ? "…" : "").Append("\"\n");
        }
        sb.Append(total).Append(total == 1 ? " match" : " matches").Append(total > 50 ? " (first 50 shown)" : "").Append('\n');
        return sb.ToString();
    }

    /// <summary>The nodes that carry editable text, with it: paragraphs, headings, code, cells and topics; a match is reported once, on the
    /// outermost node that has the text, so a table cell's paragraphs are not listed again under it. Worksheet cells give their value.</summary>
    static IEnumerable<(Node Node, string Text)> TextNodes(Node node, string format)
    {
        foreach (var child in node.Children)
        {
            var kind = child.Kind;
            if (kind is "run" or "comment" or "decor" or "image" or "chart") continue;
            if (kind is "paragraph" or "heading" or "code" or "cell" or "topic" && child.GetProps().TryGetValue(format == "xlsx" ? "value" : "text", out var text) && text.Length > 0)
            {
                yield return (child, text);
                if (kind != "topic") continue;
            }
            foreach (var inner in TextNodes(child, format)) yield return inner;
        }
    }

    /// <summary>`replace <file> [path] --find a --with b [--preview] [--ignore-case]`: every occurrence in the document, or under the path.
    /// Word paragraphs keep their character formatting (only the changed words are rewritten); markdown and slides change run by run.</summary>
    public static string Replace(Args a)
    {
        var file = a.Need(0, "file");
        using var doc = Files.Open(file);
        var find = a.Opt("find") ?? throw new WriterException(ErrorCode.Usage, "--find is required", "Example: replace report.docx --find \"Q3\" --with \"Q4\"");
        if (find.Length == 0) throw new WriterException(ErrorCode.Usage, "--find must not be empty", "Give the text to look for.");
        var preview = a.Flag("preview");
        var with = a.Opt("with") ?? (preview ? "" : throw new WriterException(ErrorCode.Usage, "--with is required", "Add --with \"new text\", or --preview to count the matches first."));
        var comparison = a.Flag("ignore-case") ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var roots = a.Positional.Count > 1 ? PathResolver.Query(doc.Root, a.Positional[1]) : [doc.Root];
        var hits = new List<(Node Node, int Count, string Text)>();
        foreach (var root in roots)
        {
            var own = root.Kind is "paragraph" or "heading" or "code" or "cell" or "topic" && root.GetProps().TryGetValue(doc.Format == "xlsx" ? "value" : "text", out var t) ? [(root, t)] : Enumerable.Empty<(Node, string)>();
            foreach (var (node, text) in own.Concat(TextNodes(root, doc.Format)))
            {
                var count = Count(text, find, comparison);
                if (count > 0 && hits.All(h => !ReferenceEquals(h.Node.Anchor, node.Anchor))) hits.Add((node, count, text));
            }
        }
        if (!preview && hits.Count > 0)
        {
            Files.EnsureWritable(doc);
            foreach (var (node, _, text) in hits)
            {
                var runs = doc.Format == "docx" ? [] : node.Children.Where(c => c.Kind == "run").ToList();
                if (runs.Count == 0)
                {
                    if (doc.Format == "xlsx" && node.GetProps().ContainsKey("formula")) continue;
                    Mutations.Set(node, [new(doc.Format == "xlsx" ? "value" : "text", text.Replace(find, with, comparison))]);
                    continue;
                }
                // ponytail: a match that spans two runs is counted on the node but left as it is; the node-level text= route would drop the formatting
                foreach (var run in runs)
                    if (run.Text is { } rt && rt.Contains(find, comparison)) Mutations.Set(run, [new("text", rt.Replace(find, with, comparison))]);
            }
            Files.SaveAtomic(doc, file);
        }
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteNumber("matches", hits.Sum(h => h.Count));
            w.WriteNumber("nodes", hits.Count);
            w.WriteBoolean("replaced", !preview && hits.Count > 0);
            w.WriteStartArray("hits");
            foreach (var (node, count, text) in hits.Take(50))
            {
                w.WriteStartObject();
                w.WriteString("path", node.Path);
                w.WriteNumber("count", count);
                w.WriteString("text", NodeJson.Preview(text, 80));
                w.WriteEndObject();
            }
            w.WriteEndArray();
            w.WriteEndObject();
        }) + "\n";
    }

    static int Count(string text, string find, StringComparison comparison)
    {
        var n = 0;
        for (var at = text.IndexOf(find, comparison); at >= 0; at = text.IndexOf(find, at + find.Length, comparison)) n++;
        return n;
    }

    /// <summary>`section <file> <heading path> [--md text | --remove]`: a heading with everything up to the next heading of the same or a
    /// higher level. Bare, it prints the section with full text; --md replaces the whole section (heading included) with the blocks the
    /// markdown gives; --remove deletes it. One save, so a rewrite lands whole or not at all.</summary>
    public static string Section(Args a)
    {
        var file = a.Need(0, "file");
        using var doc = Files.Open(file);
        var heading = PathResolver.Single(doc.Root, a.Need(1, "heading path"));
        if (heading.Kind != "heading")
            throw new WriterException(ErrorCode.Validation, $"{heading.Path} is a {heading.Kind}, not a heading", "Run 'writer view <file> structure' to see the headings and their paths.");
        var blocks = SectionBlocks(heading);
        var parent = heading.Parent!;
        var markdown = a.Opt("md");
        if (markdown is null && !a.Flag("remove")) return Print(blocks);
        Files.EnsureWritable(doc);
        var index = parent.Children.TakeWhile(c => !ReferenceEquals(c.Anchor, heading.Anchor)).Count() + 1;
        var removed = blocks.Select(b => b.Path).ToList();
        int before; // the parent's children once the old section is out; the difference to afterwards is what the markdown added
        if (markdown is not null && doc.Format == "md")
        {
            // the heading's own chunk takes the new markdown as written; the rest of the section goes
            foreach (var block in blocks.Skip(1).Reverse()) block.Remove();
            before = parent.Children.Count - 1;
            heading.SetRaw(markdown);
        }
        else
        {
            foreach (var block in Enumerable.Reverse(blocks)) block.Remove();
            before = parent.Children.Count;
            if (markdown is not null) InsertMarkdown(parent, markdown, index);
        }
        Files.SaveAtomic(doc, file);
        if (markdown is null)
            return NodeJson.Write(w =>
            {
                w.WriteStartObject();
                w.WriteStartArray("removed");
                foreach (var p in removed) w.WriteStringValue(p);
                w.WriteEndArray();
                w.WriteEndObject();
            }) + "\n";
        var children = parent.Children;
        return "Now:\n" + Print(children.Skip(index - 1).Take(children.Count - before).ToList());
    }

    /// <summary>The heading and the blocks after it, up to the next heading of the same or a higher (smaller-numbered) level.</summary>
    public static List<Node> SectionBlocks(Node heading)
    {
        var level = int.Parse(heading.GetProps().GetValueOrDefault("level") ?? "1", CultureInfo.InvariantCulture);
        var siblings = heading.Parent!.Children;
        var start = siblings.TakeWhile(c => !ReferenceEquals(c.Anchor, heading.Anchor)).Count();
        var blocks = new List<Node> { siblings[start] };
        foreach (var block in siblings.Skip(start + 1))
        {
            if (block.Kind == "heading" && int.Parse(block.GetProps().GetValueOrDefault("level") ?? "1", CultureInfo.InvariantCulture) <= level) break;
            blocks.Add(block);
        }
        return blocks;
    }

    /// <summary>Blocks with their full text: paragraphs on one line, tables as tab-separated rows.</summary>
    static string Print(IReadOnlyList<Node> blocks)
    {
        var sb = new StringBuilder();
        foreach (var block in blocks)
        {
            sb.Append(block.Path);
            var props = Registry.ToDisplay(block.Format, block.Kind, block.GetProps());
            foreach (var (name, value) in props)
                if (name is "level" or "list" or "style" or "align" or "lang" or "rows" or "cols") sb.Append("  ").Append(name).Append('=').Append(value);
            if (block.Kind == "table")
            {
                sb.Append('\n');
                foreach (var row in block.Children.Where(c => c.Kind == "row"))
                    sb.Append("    ").AppendJoin('\t', row.Children.Select(c => (c.Text ?? "").ReplaceLineEndings(" "))).Append('\n');
                continue;
            }
            if (props.TryGetValue("text", out var text)) sb.Append("  \"").Append(text.ReplaceLineEndings("\\n")).Append('"');
            sb.Append('\n');
        }
        return sb.ToString();
    }

    /// <summary>Adds the blocks of a markdown text under the parent from the 1-based position on: headings, paragraphs and list items
    /// with their inline formatting, code and tables. Works for Word bodies as for markdown ones.</summary>
    public static List<Node> InsertMarkdown(Node parent, string markdown, int? index)
    {
        using var md = new MdAdapter().Open(new MemoryStream(Encoding.UTF8.GetBytes(markdown)));
        var added = new List<Node>();
        foreach (var block in md.Root.Children[0].Children)
        {
            var source = block.GetProps();
            var props = new Dictionary<string, string>();
            var kind = block.Kind == "image" ? "paragraph" : block.Kind;
            var accepted = Registry.PropsFor(Registry.Get(parent.Format, kind), parent.Format).Where(p => !p.ReadOnly).Select(p => p.Name).ToHashSet();
            foreach (var name in new[] { "level", "list", "lang", "data" })
                if (accepted.Contains(name) && source.TryGetValue(name, out var v)) props[name] = v;
            if (block.Kind == "image") props["text"] = source.GetValueOrDefault("alt") is { Length: > 0 } alt ? alt : source.GetValueOrDefault("src") ?? "";
            else if (kind is "heading" or "paragraph" && source.TryGetValue("html", out var html)) props["html"] = html;
            else if (kind == "code") props["text"] = source.GetValueOrDefault("text") ?? "";
            added.Add(Mutations.Add(parent, kind, props, index is { } i ? i + added.Count : null));
        }
        if (added.Count == 0) throw new WriterException(ErrorCode.Validation, "The markdown has no blocks", "Give at least one paragraph or heading.");
        return added;
    }

    /// <summary>`formula <file> <cell or range path> <formula> [--check]`: writes the formula (a range fills like Excel's fill handle) and
    /// reports what it refers to — each reference's values or a range's shape — with warnings for text in a numeric range, functions this
    /// editor does not compute, unbalanced parentheses and a reference to the cell itself. --check reports without writing.</summary>
    public static string Formula(Args a)
    {
        var file = a.Need(0, "file");
        using var doc = Files.Open(file);
        if (doc.Format != "xlsx") throw new WriterException(ErrorCode.Validation, "formula works on .xlsx files", "For a Word or markdown table, compute the value and write it with set.");
        var target = PathResolver.Single(doc.Root, a.Need(1, "cell or range path"));
        if (target.Kind is not ("cell" or "range")) throw new WriterException(ErrorCode.Validation, $"{target.Path} is a {target.Kind}", "Give a cell (/sheet[1]/cell[E2]) or a range (/sheet[1]/range[E2:E20]).");
        var formula = a.Need(2, "formula").Trim().TrimStart('=');
        if (formula.Length == 0) throw new WriterException(ErrorCode.Usage, "The formula is empty", "Example: formula book.xlsx /sheet[1]/cell[E2] \"C2*D2\"");
        var sheet = target.Parent!;
        var warnings = new List<string>();
        var quoted = Regex.Replace(formula, "\"[^\"]*\"", m => new string(' ', m.Length));
        if (quoted.Count(c => c == '(') != quoted.Count(c => c == ')')) warnings.Add("parentheses are not balanced");
        if (formula.Count(c => c == '"') % 2 == 1) warnings.Add("a string literal is not closed");
        foreach (Match m in Regex.Matches(quoted, @"(?<![\w.])([A-Za-z_][\w.]*)\s*\("))
        {
            var name = m.Groups[1].Value.ToUpperInvariant();
            if (name.StartsWith("_XLFN.", StringComparison.Ordinal)) name = name[6..];
            if (!Functions.Contains(name)) warnings.Add($"{name} is not a function this editor computes (Excel may; check the spelling)");
        }
        var targetBox = Box(target.Key!);
        var refs = new List<(string Text, string Sheet, (int C1, int R1, int C2, int R2) Box)>();
        foreach (Match m in Regex.Matches(quoted, @"(?:(?:'([^']+)'|([A-Za-z_][\w.]*))!)?(?<![\w.$])(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)(?![\w(])"))
        {
            var sheetName = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Success ? m.Groups[2].Value : sheet.Name ?? "";
            var text = m.Groups[3].Value.Replace("$", "").ToUpperInvariant();
            if (refs.Any(r => r.Text == text && r.Sheet == sheetName)) continue;
            refs.Add((text, sheetName, Box(text)));
        }
        var report = new Dictionary<string, string>();
        foreach (var (text, sheetName, box) in refs)
        {
            var on = doc.Root.Children.FirstOrDefault(s => s.Kind == "sheet" && string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (on is null)
            {
                warnings.Add($"no sheet named {sheetName}");
                continue;
            }
            var same = ReferenceEquals(on.Anchor, sheet.Anchor);
            if (same && box.C1 <= targetBox.C2 && box.C2 >= targetBox.C1 && box.R1 <= targetBox.R2 && box.R2 >= targetBox.R1)
                warnings.Add($"{text} overlaps the target {target.Key}: a circular reference");
            var label = same ? text : sheetName + "!" + text;
            if (box.C1 == box.C2 && box.R1 == box.R2)
            {
                var props = on.ResolveVirtual("cell", text)!.GetProps();
                report[label] = props.TryGetValue("formula", out var f) ? $"={f} → {props.GetValueOrDefault("value")}" : props.GetValueOrDefault("type") is null ? "(blank)" : props.GetValueOrDefault("value") ?? "";
                continue;
            }
            int numbers = 0, blanks = 0, cells = 0;
            var texts = new List<string>();
            foreach (var row in Enumerable.Range(box.R1, Math.Min(box.R2, box.R1 + 999) - box.R1 + 1))
                foreach (var col in Enumerable.Range(box.C1, Math.Min(box.C2, box.C1 + 49) - box.C1 + 1))
                {
                    cells++;
                    var props = on.ResolveVirtual("cell", Letter(col) + row.ToString(CultureInfo.InvariantCulture))!.GetProps();
                    switch (props.GetValueOrDefault("type"))
                    {
                        case null: blanks++; break;
                        case "string": if (texts.Count < 3) texts.Add($"{Letter(col)}{row} '{NodeJson.Preview(props.GetValueOrDefault("value") ?? "", 20)}'"); else texts.Add(""); break;
                        default: numbers++; break;
                    }
                }
            var textCount = texts.Count;
            report[label] = $"{cells} cells: {numbers} numbers, {textCount} text, {blanks} blank";
            if (cells > 0 && blanks == cells) warnings.Add($"{label} is empty");
            else if (textCount > 0) warnings.Add($"{label} has {textCount} text cell{(textCount == 1 ? "" : "s")} ({string.Join(", ", texts.Where(t => t.Length > 0))}{(textCount > 3 ? ", …" : "")}): SUM and AVERAGE skip them");
        }
        var written = false;
        if (!a.Flag("check"))
        {
            Files.EnsureWritable(doc);
            Mutations.Set(target, [new("formula", formula)]);
            Files.SaveAtomic(doc, file);
            written = true;
        }
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteString("target", target.Key);
            w.WriteString("formula", formula);
            w.WriteBoolean("written", written);
            w.WriteStartObject("refs");
            foreach (var (label, value) in report) w.WriteString(label, value);
            w.WriteEndObject();
            w.WriteStartArray("warnings");
            foreach (var warning in warnings) w.WriteStringValue(warning);
            w.WriteEndArray();
            w.WriteString("note", "Results compute when the sheet opens in the editor or in Excel.");
            w.WriteEndObject();
        }) + "\n";
    }

    static (int C1, int R1, int C2, int R2) Box(string reference)
    {
        var parts = reference.Replace("$", "").ToUpperInvariant().Split(':');
        var (c1, r1) = Cell(parts[0]);
        var (c2, r2) = parts.Length > 1 ? Cell(parts[1]) : (c1, r1);
        return (Math.Min(c1, c2), Math.Min(r1, r2), Math.Max(c1, c2), Math.Max(r1, r2));

        static (int, int) Cell(string s)
        {
            var letters = new string(s.TakeWhile(char.IsAsciiLetter).ToArray());
            return (Column(letters), int.Parse(s[letters.Length..], CultureInfo.InvariantCulture));
        }
    }

    /// <summary>The functions the editor's own calculation engine (ui/sheet-engine.js) evaluates; a formula outside this set shows
    /// its result only once Excel has opened the file.</summary>
    static readonly HashSet<string> Functions = new(("ABS ADDRESS AGGREGATE AND AREAS AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS CEILING CEILING.MATH CELL CHAR CHOOSE CLEAN CODE COLUMN COLUMNS COMBIN CONCAT CONCATENATE CONVERT CORREL COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS CUMIPMT CUMPRINC DATE DATEDIF DATEVALUE DAYS DAYS360 DB DDB DEC2BIN DEC2HEX DEC2OCT DELTA DOLLAR EDATE EFFECT EOMONTH ERROR.TYPE EVEN EXACT EXP FACT FALSE FILTER FIND FINDB FIXED FLOOR FLOOR.MATH FORECAST FORECAST.LINEAR FREQUENCY FV GCD GEOMEAN GROWTH HARMEAN HLOOKUP HOUR IF IFERROR IFNA IFS INDEX INDIRECT INFO INT INTERCEPT IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISFORMULA ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISOWEEKNUM ISREF ISTEXT LARGE LCM LEFT LEFTB LEN LENB LET LN LOG LOG10 LOOKUP LOWER MATCH MAX MAXA MAXIFS MEDIAN MID MIDB MIN MINA MINIFS MOD MODE MODE.SNGL MROUND N NA NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NORMDIST NORMINV NORMSDIST NORMSINV NOT NOW NPER NPV NUMBERVALUE ODD OFFSET OR PEARSON PERCENTILE PERCENTILE.EXC PERCENTILE.INC PI PMT POWER PPMT PRODUCT PROPER PV QUARTILE QUARTILE.EXC QUARTILE.INC QUOTIENT RAND RANDBETWEEN RANK RANK.AVG RANK.EQ RATE REPLACE REPT RIGHT RIGHTB RMB ROUND ROUNDDOWN ROUNDUP ROW ROWS SEARCH SEQUENCE SIGN SLN SLOPE SMALL SORT SORTBY SQRT STDEV STDEV.P STDEV.S STDEVP SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SWITCH SYD T TEXT TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TIME TIMEVALUE TODAY TRANSPOSE TREND TRIM TRUE TRUNC TYPE UNICHAR UNICODE UNIQUE UPPER VALUE VAR VAR.P VAR.S VARP VLOOKUP WEEKDAY WEEKNUM WORKDAY WORKDAY.INTL XIRR XLOOKUP XMATCH XNPV XOR YEAR YEARFRAC").Split(' '));

    /// <summary>`batch <file> --run "<command>" [--run ...]`: the commands in order on a copy of the file, then the copy takes the file's
    /// place — one write, and none when any command fails. Every command addresses this file; its file argument is rewritten to the copy.</summary>
    public static string Batch(Args a)
    {
        var file = a.Need(0, "file");
        if (!File.Exists(file)) throw new WriterException(ErrorCode.FileNotFound, $"{file} not found", "Check the path.");
        var runs = a.All("run");
        if (runs.Count == 0) throw new WriterException(ErrorCode.Usage, "batch needs at least one --run", "Example: batch a.docx --run \"set a.docx /body/paragraph[1] --prop text=Hi\" --run \"add a.docx /body --type paragraph --prop text=Bye\"");
        var tmp = Path.Combine(Path.GetTempPath(), $"writer-batch-{Guid.NewGuid():N}{Path.GetExtension(file)}");
        var results = new List<(string Command, string Output)>();
        try
        {
            File.Copy(file, tmp);
            for (var i = 0; i < runs.Count; i++)
            {
                var argv = Mcp.Tokenize(runs[i]);
                if (argv.Length > 0 && argv[0] == "writer") argv = argv[1..];
                if (argv.Length < 2 || argv[0] is "batch" or "create" or "export" or "mcp" or "serve" or "watch" or "app" or "help")
                    throw new WriterException(ErrorCode.Usage, $"Command {i + 1} cannot run in a batch: {runs[i]}", "A batch takes get, query, view, search, section, replace, formula, add, set, remove, move and copy on the batch's file.");
                argv[1] = tmp;
                var stdout = new StringWriter();
                var stderr = new StringWriter();
                if (Runner.Run(argv, stdout, stderr) != 0)
                {
                    var error = stderr.ToString().Replace(tmp, file, StringComparison.Ordinal);
                    string message = error.Trim(), hint = "";
                    try
                    {
                        var e = JsonDocument.Parse(error).RootElement.GetProperty("error");
                        message = e.GetProperty("message").GetString() ?? message;
                        hint = e.GetProperty("hint").GetString() ?? "";
                    }
                    catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException) { }
                    throw new WriterException(ErrorCode.Validation, $"Command {i + 1} of {runs.Count} failed, nothing was written — {runs[i]}: {message}", hint);
                }
                results.Add((runs[i], stdout.ToString().Replace(tmp, file, StringComparison.Ordinal)));
            }
            var bytes = File.ReadAllBytes(tmp);
            Files.ReplaceAtomic(file, s => s.Write(bytes));
        }
        finally
        {
            if (File.Exists(tmp)) File.Delete(tmp);
        }
        return NodeJson.Write(w =>
        {
            w.WriteStartObject();
            w.WriteNumber("commands", results.Count);
            w.WriteStartArray("results");
            foreach (var (command, output) in results)
            {
                w.WriteStartObject();
                w.WriteString("command", command);
                w.WriteString("output", NodeJson.Preview(output.TrimEnd('\n'), 300));
                w.WriteEndObject();
            }
            w.WriteEndArray();
            w.WriteEndObject();
        }) + "\n";
    }
}
