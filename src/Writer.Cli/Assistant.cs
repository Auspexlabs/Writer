using System.Text;
using System.Text.Json.Nodes;
using Writer.Core;
using Writer.Formats;

namespace Writer.Cli;

/// <summary>The assistant's side that needs no network: the system prompt for the open file, the editing tools (writer, batch,
/// plan) and running them. Chat.cs talks to the model APIs with it; the engine in a web page (src/Writer.Browser) serves the
/// embedded editor's assistant with it, the site answering the model calls (docs/embed.md).</summary>
public static class Assistant
{
    public const string ToolName = "writer";

    public const string ToolDescription =
        "Run a writer command against a document (docx, xlsx, pptx, md read/write; pdf read). "
        + "Same syntax as the CLI without the program name, e.g. \"view report.docx outline\", "
        + "\"set report.docx /body/paragraph[2] --prop text=\\\"New text\\\"\", \"help docx paragraph\". "
        + "Returns the command's JSON or text. Start with \"help\" and \"view <file> outline\".";

    public const string Instructions =
        "Workflow: `view <file> outline` to see every element with its path, `query` or `get` to inspect, "
        + "`add`/`set`/`remove`/`move`/`copy` to change, `view <file> html` to check the result, `export` to convert. "
        + "When unsure about a property, run `help <format> <element>` instead of guessing.";

    public const string BatchToolName = "batch";
    public const string PlanToolName = "plan";

    /// <summary>An outline up to this long goes into the prompt whole; a longer document is shown as its structure instead, and the
    /// model reads the parts it needs with section, search and get.</summary>
    public const int OutlineBudget = 16000;

    /// <summary>The editing tools in order: one writer command, an atomic batch of them, and the scratch plan.</summary>
    public static IEnumerable<(string Name, string Description, JsonObject Schema)> EditingTools()
    {
        yield return (ToolName, ToolDescription, Parameters());
        yield return (BatchToolName,
            "Run several writer commands on the open file as one step: they apply in order and the file is saved once; when any command fails nothing is written and the error names the command. "
            + "Use it for every change that takes more than one command (the edits of one section, a table's rows, a summary row with its formatting). "
            + "label is the step's name in the user's language, shown while it runs, e.g. 正在改第 2 节…",
            new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["commands"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "Writer command lines, without the program name, all on the open file." },
                    ["label"] = new JsonObject { ["type"] = "string", ["description"] = "What this step does, a few words in the user's language." },
                },
                ["required"] = new JsonArray("commands"),
            });
        yield return (PlanToolName,
            "Your scratch plan for a task with several parts: numbered steps, one line each, marked done as you go. Call it before the first edit of such a task and again whenever the plan changes. It writes nothing.",
            new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject { ["text"] = new JsonObject { ["type"] = "string", ["description"] = "The plan, e.g. 1. 读第 2 节 ✓ 2. 改写 3. 补表格" } },
                ["required"] = new JsonArray("text"),
            });
    }

    /// <summary>Runs an editing tool call (writer, batch or plan) and returns what the tool event shows: a display line (the command
    /// as given, a batch's label, 计划), the exit code, the output, and whether the file was written. runArgv runs one command
    /// line against the workspace (Serve.RunArgv, or the page's engine).</summary>
    public static (string Display, int Code, string Output, bool Wrote) RunEditingTool(string? tool, JsonObject? input, Func<string[], (int Code, string Text)> runArgv)
    {
        if (tool == PlanToolName)
        {
            var text = Str(input?["text"])?.Trim();
            if (string.IsNullOrWhiteSpace(text)) return (PlanToolName, 1, "The tool input must be JSON like {\"text\": \"1. ... 2. ...\"}.", false);
            return ("计划 " + text.ReplaceLineEndings(" / "), 0, "Plan noted. Work through it and call plan again when it changes.", false);
        }
        if (tool == BatchToolName)
        {
            var commands = (input?["commands"] as JsonArray)?.Select(Str).Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => c!).ToList() ?? [];
            if (commands.Count == 0) return (BatchToolName, 1, "The tool input must be JSON like {\"commands\": [\"set a.docx /body/paragraph[1] --prop text=Hi\"], \"label\": \"...\"}.", false);
            var label = Str(input?["label"])?.Trim() is { Length: > 0 } l ? l : $"批量修改 {commands.Count} 条";
            string[] first;
            try
            {
                first = Args.Tokenize(commands[0]);
            }
            catch (WriterException ex)
            {
                return (label, 1, ex.Message + " " + ex.Hint, false);
            }
            if (first.Length > 0 && first[0] == "writer") first = first[1..];
            if (first.Length < 2) return (label, 1, "Each batch command names the file, e.g. set a.docx /body/paragraph[1] --prop text=Hi.", false);
            var argv = new List<string> { "batch", first[1] };
            foreach (var command in commands)
            {
                argv.Add("--run");
                argv.Add(command);
            }
            var (code, output) = runArgv(argv.ToArray());
            return (label, code, output, code == 0);
        }
        var line = Str(input?["command"]);
        if (line is null) return ("", 1, "The tool input must be JSON like {\"command\": \"view report.docx outline\"}.", false);
        string[] tokens;
        try
        {
            tokens = Args.Tokenize(line);
        }
        catch (WriterException ex)
        {
            return (line, 1, ex.Message + " " + ex.Hint, false);
        }
        var (writerCode, writerOutput) = runArgv(tokens);
        return (line, writerCode, writerOutput, writerCode == 0 && Writes(tokens));
    }

    /// <summary>Whether a command line changes the file: the editing verbs, and section, replace and formula unless they only read.</summary>
    public static bool Writes(string[] argv)
    {
        if (argv.Length > 0 && argv[0] == "writer") argv = argv[1..];
        if (argv.Length == 0) return false;
        return argv[0] switch
        {
            "add" or "set" or "remove" or "move" or "copy" or "create" or "export" or "batch" => true,
            "section" => argv.Contains("--md") || argv.Any(t => t.StartsWith("--md=", StringComparison.Ordinal)) || argv.Contains("--remove"),
            "replace" => !argv.Contains("--preview"),
            "formula" => !argv.Contains("--check"),
            _ => false,
        };
    }

    /// <summary>The writer tool's input: one command line.</summary>
    static JsonObject Parameters() => new()
    {
        ["type"] = "object",
        ["properties"] = new JsonObject
        {
            ["command"] = new JsonObject { ["type"] = "string", ["description"] = "The writer command line, without the program name." },
        },
        ["required"] = new JsonArray("command"),
    };

    /// <summary>Rules, the rules for the open file's format with examples, the command reference for that format, and its current
    /// outline — or, for a long document and for workbooks, its structure.</summary>
    public static string SystemPrompt(string? file, string workspace, bool outline = true, bool web = false)
    {
        var sb = new StringBuilder();
        sb.Append("You are the assistant built into a document editor. The user sees the document on the left and talks to you on the right. ");
        sb.Append("You change documents only through the tools, which edit the file in place; the editor shows each change as it lands, and at the end of your turn the user gets one card to keep or undo everything you changed.\n\n");
        sb.Append("## How to work\n\n");
        sb.Append("- Look before you change: the outline or structure below gives every path; read the exact part with `section <file> <heading path>`, `search <file> <text>`, `get <file> <path>` or `view <file> text` instead of guessing what it says.\n");
        sb.Append("- Smallest edit that does the job: change the paragraph, cell or section asked for and nothing else. Formatting, styles, numbering and wording elsewhere stay as they are; never rewrite a paragraph to change one word.\n");
        sb.Append("- Text goes into an existing paragraph with `set … --prop text=` (plain), `md=` or `html=` (with inline formatting): the paragraph keeps its style and, in Word, the character formatting of the words that stay. New blocks go in with `add … --after <path>`; a whole section with `section … --md`.\n");
        sb.Append("- Two or more related commands go in one `batch` call with a short label in the user's language (正在改第 2 节…): they land together or not at all. One command: the writer tool. Paths count per kind from the top, so after adding or removing blocks read the structure again before addressing what follows.\n");
        sb.Append("- A task with several parts: call `plan` first with numbered steps, then work through them and update it as they finish. A failed command comes back with its reason — fix that one command and carry on; do not start over, and do not repeat a command that just succeeded.\n");
        sb.Append("- Before removing a table, a sheet, a whole section or more than a few paragraphs, say so in one line, then do it. Never delete or rewrite content the user did not ask about; `replace … --preview` first when a find and replace may touch many places, and say the count.\n");
        sb.Append("- Never tell the user to run commands. When the user asks a question, answer it from the document without editing.\n");
        sb.Append("- Reply in the user's language. While working, at most one short sentence between steps; finish with one line saying what changed and where (段落、表格、单元格 by their content or position, not paths), or the answer. No command output, no paths unless the user asks.\n");
        if (web) sb.Append("- Use web_search and web_fetch when the task needs current or outside facts the document doesn't have, and say which URLs you used.\n");
        sb.Append('\n').Append(Instructions).Append("\n\n");
        var format = file is null ? null : Adapters.CanonicalFormat(Path.GetExtension(file));
        if (file is not null)
        {
            var relative = Path.IsPathRooted(file) && file.StartsWith(workspace + Path.DirectorySeparatorChar, StringComparison.Ordinal)
                ? Path.GetRelativePath(workspace, file) : file;
            sb.Append("Open file: ").Append(relative).Append(" (use exactly this path in commands).\n");
            if (format is "pdf") sb.Append("PDF files are read-only: read them with `view <file> text`; to edit, `export --to <name>.docx` first and tell the user.\n");
            if (format is "docx") sb.Append('\n').Append(WordRules);
            if (format is "xlsx") sb.Append('\n').Append(ExcelRules);
            if (format is "md") sb.Append('\n').Append(MarkdownRules);
            if (format is "pptx") sb.Append('\n').Append(SlideDesign);
        }
        sb.Append("\n## Command reference\n\n").Append(Run(["help"]));
        if (format is not null and not "pdf")
            foreach (var kind in Registry.ForFormat(format))
                sb.Append('\n').Append(Run(["help", format, kind.Name]));
        if (file is not null && outline)
        {
            var text = format == "xlsx" ? "" : Run(["view", file, "outline"]);
            if (format != "xlsx" && text.Length <= OutlineBudget) sb.Append("\n## Current outline\n\n").Append(text);
            else sb.Append("\n## Structure\n\nThe document is shown in short; read the parts you need with section, search, get or view text.\n\n").Append(Trim(Run(["view", file, "structure"]), OutlineBudget));
        }
        return sb.ToString();
    }

    /// <summary>What a Word document asks for: edits that keep its formatting, structure through properties, Chinese typography.</summary>
    public const string WordRules =
        "## Word (docx)\n\n"
        + "- Paragraph text: `set f.docx /body/paragraph[3] --prop text=\"…\"` rewrites the words and keeps the bold, colour and size of the words that stay; `md=` adds inline formatting (**bold**, *italic*, [link](url)). Never remove and re-add a paragraph to change its text.\n"
        + "- Structure through properties: headings are `heading` nodes with level; list items are paragraphs with list=bullet|number and level; style=Quote, align=center. Keep the document's own heading levels, numbering and styles; do not hand-format runs one by one to imitate a style.\n"
        + "- New content: `add f.docx /body --type paragraph --prop md=\"…\" --after /body/paragraph[3]`; a whole section from markdown: `section f.docx /body/heading[2] --md \"## 标题\\n\\n正文…\\n\\n- 要点\"` (the heading line included; headings, lists and tables come out as Word blocks).\n"
        + "- Tables: cells are `/body/table[1]/row[2]/cell[3]` with text/md; `add f.docx /body/table[1] --type row --prop data='[\"…\",\"…\"]'`; `--prop data=` on the table rewrites every cell. Sums and averages in a Word table are numbers you compute from the cells and write.\n"
        + "- Chinese text: 全角标点（，。：；！？「」“”）, no space between CJK and punctuation, one space between CJK and Latin letters or digits, 半角 digits and units; match the document's quote style and number style. English text: English typography.\n"
        + "- Tracked changes: when the document has track=true, text you write is recorded as revisions; do not accept or reject revisions unless asked.\n\n"
        + "Example — 把第二段写得委婉些，再在最后加一段总结:\n"
        + "  batch label=\"改写第二段并加总结\" commands=[\"set f.docx /body/paragraph[2] --prop text=\\\"…\\\"\", \"add f.docx /body --type heading --prop text=总结 --prop level=2\", \"add f.docx /body --type paragraph --prop md=\\\"…\\\"\"]\n"
        + "Example — 全文把「甲方」改成「委托方」: `replace f.docx --find 甲方 --with 委托方 --preview`, then without --preview; reply with the count.\n";

    /// <summary>What a workbook asks for: read the data first, formulas through the formula verb and its check, ranges for formatting.</summary>
    public const string ExcelRules =
        "## Excel (xlsx)\n\n"
        + "- Address cells as /sheet[1]/cell[B3], blocks as /sheet[1]/range[A2:D20], rows as /sheet[1]/row[3]; the structure below gives each sheet's used range, header row and first data row. Read the data with `get f.xlsx /sheet[1]/range[A1:D20]` before computing or describing anything.\n"
        + "- Formulas: `formula f.xlsx /sheet[1]/cell[E2] \"C2*D2\"` writes one; on a range it fills like Excel's fill handle, shifting relative references per row and column (`formula f.xlsx /sheet[1]/range[E2:E20] \"C2*D2\"`), $ pins a reference. The result lists what the formula refers to and warns about text in a numeric range, functions this editor cannot compute, unbalanced parentheses and circular references: when it warns, fix the formula and run it once more. Results compute when the sheet opens; never write a computed number where a formula belongs, and never overwrite a formula cell with a value.\n"
        + "- Summary row: fill the formulas across (`formula f.xlsx /sheet[1]/range[B21:E21] \"SUM(B2:B20)\"`), label it (`set f.xlsx /sheet[1]/cell[A21] --prop value=合计`), then format the row (`set f.xlsx /sheet[1]/range[A21:E21] --prop bold=true --prop border=thin`) — one batch.\n"
        + "- Formatting goes on ranges: bold, fill, color, size, format (a number format: 0.00, #,##0, 0%, yyyy-mm-dd), align, wrap, border=thin, borderColor; column widths, frozen panes and the filter on the sheet (`set f.xlsx /sheet[1] --prop freeze=A2 --prop filter=A1:E20 --prop widths='{\"A\":18}'`).\n"
        + "- Sort the data rows without the header: `set f.xlsx /sheet[1]/range[A2:E20] --prop sort=C:desc` (cells move with their formatting and formulas).\n"
        + "- Charts: `add f.xlsx /sheet[1] --type chart --prop type=column --prop title=\"…\" --prop categories=A2:A13 --prop series='[{\"name\":\"B1\",\"values\":\"B2:B13\"}]'`; line for a trend, pie for shares, bar for a ranking.\n"
        + "- Values: numbers as plain digits (1234.5 — units and thousands separators belong to the number format or the header), dates as yyyy-mm-dd, text as text; `--prop values=` on a range writes a block of cells at once. Never clear a range the user did not name.\n\n"
        + "Example — 在最后加一行合计并加粗:\n"
        + "  batch label=\"添加合计行\" commands=[\"set f.xlsx /sheet[1]/cell[A21] --prop value=合计\", \"formula f.xlsx /sheet[1]/range[B21:E21] \\\"SUM(B2:B20)\\\"\", \"set f.xlsx /sheet[1]/range[A21:E21] --prop bold=true\"]\n";

    /// <summary>What a markdown file asks for: sections rewritten as markdown, inline formatting kept, the file's conventions respected.</summary>
    public const string MarkdownRules =
        "## Markdown (md)\n\n"
        + "- Blocks are /body/heading[n], /body/paragraph[n] (list items are paragraphs with list= and level=), /body/table[n], /body/code[n]; paths count per kind from the top.\n"
        + "- A section is a heading and everything up to the next heading of the same or a higher level: `section f.md /body/heading[3]` reads it, `section f.md /body/heading[3] --md \"### 标题\\n\\n…\"` rewrites it exactly as written (heading line included), `--remove` deletes it. This is the way to rewrite, expand or restructure a part.\n"
        + "- Inline formatting: `set f.md /body/paragraph[2] --prop md=\"**要点**：…\"`; text= drops the formatting. Code blocks: text= and lang=.\n"
        + "- Tables: `/body/table[1]/row[2]/cell[1]` with text/md; `add f.md /body/table[1] --type row --prop data='[\"…\",\"…\"]'`; `--prop data='[[…],[…]]'` rewrites the whole table (the first row is the header).\n"
        + "- Lists: `add f.md /body --type paragraph --prop md=\"…\" --prop list=bullet --prop level=0 --after /body/paragraph[4]` continues a list; numbering renders itself.\n"
        + "- Keep the file's conventions: heading levels, bullet marker, blank lines, front matter, raw HTML and quote blocks stay as they are.\n\n"
        + "Example — 把「安装」一节改成三步: `section f.md /body/heading[2] --md \"## 安装\\n\\n1. …\\n2. …\\n3. …\"`\n";

    /// <summary>The design rules the assistant works by on a slide deck, and the properties that carry them out (the slide editor's
    /// ✦ 美化 asks for exactly this; a user's own "make it nicer" gets the same treatment).</summary>
    public const string SlideDesign =
        "## Slide design\n\n"
        + "When asked to design, polish or tidy slides, work by these rules and say in one line what you applied:\n"
        + "- One idea per slide: a title that states the point (a sentence, not a topic), at most five bullets of at most ten words; a crowded slide is split (a new slide of the same layout takes the rest) rather than shrunk below 18pt.\n"
        + "- Type scale: title 36–44pt, section title 40–48pt, body 20–24pt, captions and sources 12–14pt; at most two text sizes besides the title on a slide; the same size for the same role on every slide.\n"
        + "- Contrast: dark text on a light background or light on dark, never mid-grey on grey; one accent colour, on at most two elements per slide.\n"
        + "- Alignment and margins: nothing closer to an edge than a sixteenth of the slide's width; text left-aligned; neighbouring shapes share an edge or a centre line; equal gaps between siblings.\n"
        + "- Spacing: body line spacing 1.1–1.3, clear room between title and body, no shape touching another.\n"
        + "- Layout choice: title for the opening and the close, section for a chapter break, content for a list, two or comparison for two things side by side, picture or caption for one image with a short text, quote for a quotation, titleOnly for a chart or one big number, blank for a full-bleed picture.\n"
        + "- Palette: one palette for the whole deck (paper or mist for reports, ink or night for a keynote, sea for technology, clay or sand for warm subjects, rose for consumer topics) and one fonts=heading,body; never colour shapes one by one to fake a theme.\n\n"
        + "The tools: the document's palette= and fonts=; a slide's layout= (moves its placeholders as PowerPoint does), align=edge:paths, distribute=axis:paths and background=; a shape's fit=shrink when it reads overflow=true (or shorten its text, or enlarge the box), size=, and x= y= w= h=. "
        + "Read `view <file> outline` first — it shows every shape's box, size and overflow — then make the fewest edits that satisfy the rules, one command per change, and never rewrite the user's words unless asked.\n";

    // ---------- the autocomplete (POST /complete): the words the user is about to type, shown after the caret ----------

    /// <summary>How much of the text around the caret a completion sees.</summary>
    public const int CompleteBefore = 2000, CompleteAfter = 400;

    public const string CompleteSystem =
        "You are the autocomplete of a document editor. The user is typing; predict the words they will type next at the cursor.\n"
        + "Reply with the continuation only: no quotes, no explanations, no markdown, nothing before or after it.\n"
        + "Keep it short: finish the current sentence, or add one short sentence (at most about 20 words, or 30 Chinese characters).\n"
        + "Match the language, tone and style of the text, and what is typed where the <where> tag says. Start exactly at the cursor: "
        + "finish an unfinished word, and begin with a space only when the text needs one there.\n"
        + "Never repeat what is already before the cursor, and do not run into the text after it.\n"
        + "If nothing useful fits, reply with nothing.";

    /// <summary>The user message of a completion request: where the text is typed (<paramref name="hint"/>, from the editor: "a
    /// spreadsheet cell …", "the speaker notes …"), the text before the caret, and after it when there is any.</summary>
    public static string CompleteMessage(string before, string after, string? hint = null)
    {
        before = before.Length > CompleteBefore ? before[^CompleteBefore..] : before;
        after = after.Length > CompleteAfter ? after[..CompleteAfter] : after;
        var sb = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(hint)) sb.Append("<where>").Append(hint.Length > 1200 ? hint[..1200] : hint).Append("</where>\n");
        sb.Append("<text_before_cursor>\n").Append(before).Append("\n</text_before_cursor>\n");
        if (after.Trim().Length > 0) sb.Append("<text_after_cursor>\n").Append(after).Append("\n</text_after_cursor>\n");
        return sb.Append("Continue the text at the cursor.").ToString();
    }

    /// <summary>A model's completion as the editor shows it: one line, without a code fence or wrapping quotes, without the end of
    /// the text before the caret when the model repeated it, and with no second space at the caret. "" when nothing is left.</summary>
    public static string CleanCompletion(string? text, string before)
    {
        if (string.IsNullOrEmpty(text)) return "";
        var t = text.Replace("\r", "");
        if (t.StartsWith("```", StringComparison.Ordinal)) t = t[(t.IndexOf('\n') is var n and >= 0 ? n + 1 : t.Length)..];
        var nl = t.IndexOf('\n');
        if (nl == 0) return ""; // a new paragraph: the editor suggests within the one being typed
        if (nl > 0) t = t[..nl];
        t = t.TrimEnd();
        if (t.Length >= 2 && ((t[0] == '"' && t[^1] == '"') || (t[0] == '“' && t[^1] == '”')) && t.IndexOf(t[0], 1) == t.Length - 1) t = t[1..^1];
        for (var k = Math.Min(Math.Min(t.Length, before.Length), 80); k >= 4; k--)
            if (before.EndsWith(t[..k], StringComparison.Ordinal)) { t = t[k..]; break; }
        if (before.Length > 0 && (char.IsWhiteSpace(before[^1]) || IsWide(before[^1]))) t = t.TrimStart();
        if (t.Length > 240) t = t[..240];
        return t.Trim().Length == 0 ? "" : t;
    }

    /// <summary>A CJK character or full-width punctuation: no space follows it.</summary>
    static bool IsWide(char c) => c is >= '⺀' and <= '鿿' or >= '豈' and <= '﫿' or >= '＀' and <= '￯' or >= '　' and <= '〿';

    static string Run(string[] argv)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var code = Runner.Run(argv, stdout, stderr);
        return code == 0 ? stdout.ToString() : stderr.ToString();
    }

    static string Trim(string s, int max) => s.Length <= max ? s : s[..max] + "\n…(truncated)";

    static string? Str(JsonNode? node) => node is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;
}
