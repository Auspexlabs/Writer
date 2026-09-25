using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Cli;
using Writer.Core;
using Writer.Formats;
using Writer.Formats.Compat;
using Writer.Formats.Docx;
using Writer.Formats.Xlsx;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Formats Writer opens but never writes: they become a Word, Excel or PowerPoint document in memory, the
/// tree reads that, and export to that format hands it over unchanged.</summary>
public class CompatTests : IDisposable
{
    readonly string _dir = Directory.CreateTempSubdirectory("writer-compat").FullName;
    public void Dispose() => Directory.Delete(_dir, true);
    string In(string name) => Path.Combine(_dir, name);

    static (int Code, string Out, string Err) Run(params string[] argv)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var code = Runner.Run(argv, stdout, stderr);
        return (code, stdout.ToString(), stderr.ToString());
    }

    static Document Open(string name, byte[] bytes) => Adapters.ForPath(name).Open(new MemoryStream(bytes));
    static Document Open(string name, string text) => Open(name, Encoding.UTF8.GetBytes(text));

    static string Cell(Document doc, string reference) => PathResolver.Single(doc.Root, $"/sheet[1]/cell[{reference}]").GetProps().GetValueOrDefault("text") ?? "";

    [Fact]
    public void Every_readable_extension_names_its_editor()
    {
        var open = Adapters.OpensAs;
        Assert.Equal("docx", open["doc"]);
        Assert.Equal("docx", open["wps"]);
        Assert.Equal("docx", open["docx"]);
        Assert.Equal("xlsx", open["et"]);
        Assert.Equal("xlsx", open["csv"]);
        Assert.Equal("pptx", open["dps"]);
        Assert.Equal("pptx", open["potx"]);
        Assert.Equal("md", open["txt"]);
        Assert.Equal("ppt", Adapters.CanonicalFormat(".ppt"));
        Assert.Equal("pptx", Adapters.CanonicalFormat("pptx"));
        Assert.False(Adapters.ForPath("a.ppt").CanWrite);
    }

    [Fact]
    public void Csv_opens_as_a_typed_sheet_and_exports_to_xlsx_unchanged()
    {
        using var doc = Open("data.csv", "Name,Score,Date,Code\nAnn,90,2026-01-05,007\n\"Li, Wei\",\"3,5\",,x\n");
        Assert.Equal("csv", doc.Format);
        Assert.Equal("Ann", Cell(doc, "A2"));
        Assert.Equal("90", Cell(doc, "B2"));
        Assert.Equal("007", Cell(doc, "D2"));
        Assert.Equal("Li, Wei", Cell(doc, "A3"));
        Assert.Equal("date", PathResolver.Single(doc.Root, "/sheet[1]/cell[C2]").GetProps()["type"]);

        var (target, warnings) = Exporter.Export(doc, new XlsxAdapter(), In("data.xlsx"));
        Assert.Empty(warnings);
        Assert.Equal("xlsx", target.Format);
        var ms = new MemoryStream();
        target.Save(ms);
        target.Dispose();
        using var again = new XlsxAdapter().Open(new MemoryStream(ms.ToArray()));
        Assert.Equal("Li, Wei", Cell(again, "A3"));
    }

    [Fact]
    public void Tsv_and_gbk_text_are_decoded()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var gbk = Encoding.GetEncoding(936).GetBytes("姓名\t分数\n张三\t88\n");
        using var doc = Open("表.tsv", gbk);
        Assert.Equal("张三", Cell(doc, "A2"));
        Assert.Equal("88", Cell(doc, "B2"));
    }

    [Fact]
    public void Text_file_is_a_native_markdown_document()
    {
        Assert.Equal("md", Adapters.ForPath("notes.txt").Format);
        Assert.True(Adapters.ForPath("notes.txt").CanWrite);
    }

    [Fact]
    public void Html_keeps_headings_runs_lists_tables_and_pictures()
    {
        var png = Convert.ToBase64String(FakePng(4, 2));
        var html = $$"""
            <!DOCTYPE html><html><head><title>t</title><style>p{color:red}</style></head><body>
            <h2 align="center">Title &amp; more</h2>
            <p>Plain <strong>bold</strong> and <em>italic</em> with <a href="https://example.com">a link</a>.</p>
            <ul><li>one</li><li>two<ul><li>nested</li></ul></li></ul>
            <ol><li>first</li></ol>
            <table><tr><th colspan="2">Head</th></tr><tr><td bgcolor="#ffff00">a</td><td>b<br>c</td></tr></table>
            <img src="data:image/png;base64,{{png}}" width="96">
            <blockquote>quoted</blockquote>
            <pre>line 1
            line 2</pre>
            </body></html>
            """;
        using var doc = Open("page.html", html);
        var body = doc.Root.Children.First(c => c.Kind == "body").Children;
        var kinds = body.Select(b => b.Kind).ToList();
        Assert.Equal(["heading", "paragraph", "paragraph", "paragraph", "paragraph", "paragraph", "table", "image", "paragraph", "paragraph"], kinds);
        Assert.Equal("Title & more", body[0].Text);
        Assert.Equal("2", body[0].GetProps()["level"]);
        Assert.Equal("center", body[0].GetProps()["align"]);
        var runs = body[1].Children.Where(c => c.Kind == "run").Select(r => r.GetProps()).ToList();
        Assert.Contains(runs, r => r["text"] == "bold" && r.GetValueOrDefault("bold") == "true");
        Assert.Contains(runs, r => r["text"] == "italic" && r.GetValueOrDefault("italic") == "true");
        Assert.Contains(runs, r => r["text"] == "a link" && r.GetValueOrDefault("link") == "https://example.com");
        Assert.Equal("bullet", body[2].GetProps()["list"]);
        Assert.Equal("1", body[4].GetProps()["level"]);
        Assert.Equal("number", body[5].GetProps()["list"]);
        var rows = body[6].Children.Where(c => c.Kind == "row").ToList();
        Assert.Equal(2, rows.Count);
        var head = rows[0].Children.First(c => c.Kind == "cell");
        Assert.Equal("2", head.GetProps()["colspan"]);
        Assert.Equal("Head", head.Text);
        var cells = rows[1].Children.Where(c => c.Kind == "cell").ToList();
        Assert.Equal("FFFF00", cells[0].GetProps()["fill"]);
        Assert.Equal("a", cells[0].Text);
        Assert.Contains("b", cells[1].Text);
        Assert.Equal("Quote", body[8].GetProps()["style"]);
        Assert.Equal("quoted", body[8].Text);
    }

    [Fact]
    public void An_html_table_saved_as_xls_opens_as_a_sheet()
    {
        var fake = "<html><body><table><tr><td>Item</td><td>Qty</td></tr><tr><td>Pen</td><td>12</td></tr></table></body></html>";
        using var doc = Open("export.xls", fake);
        Assert.Equal("xls", doc.Format);
        Assert.Equal("Pen", Cell(doc, "A2"));
        Assert.Equal("12", Cell(doc, "B2"));
    }

    [Fact]
    public void A_doc_that_is_plain_text_opens_as_paragraphs()
    {
        using var doc = Open("readme.doc", "first line\nsecond line");
        var body = doc.Root.Children.First(c => c.Kind == "body").Children;
        Assert.Equal(["first line", "second line"], body.Select(b => b.Text).ToList());
    }

    [Theory]
    [InlineData("macro.docm", WordprocessingDocumentType.MacroEnabledDocument)]
    [InlineData("template.dotx", WordprocessingDocumentType.Template)]
    public void Macro_enabled_and_template_word_files_open_as_a_plain_docx_without_the_vba(string name, WordprocessingDocumentType type)
    {
        var ms = new MemoryStream();
        ms.Write(Docx(P("Hello from " + name)));
        ms.Position = 0;
        using (var package = WordprocessingDocument.Open(ms, true))
        {
            package.ChangeDocumentType(type);
            if (type == WordprocessingDocumentType.MacroEnabledDocument)
            {
                var vba = package.MainDocumentPart!.AddNewPart<VbaProjectPart>();
                using var s = vba.GetStream(FileMode.Create);
                s.Write(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 });
            }
        }
        using var doc = Open(name, ms.ToArray());
        Assert.Equal("docm", doc.Format);
        Assert.Equal("Hello from " + name, doc.Root.Children.First(c => c.Kind == "body").Children[0].Text);
        var inner = ((CompatDocument)doc).Inner;
        Assert.Equal(WordprocessingDocumentType.Document, ((DocxDocument)inner).Package.DocumentType);
        Assert.Null(((DocxDocument)inner).Package.MainDocumentPart!.VbaProjectPart);
    }

    [Fact]
    public void Cli_refuses_to_write_a_compatibility_file_and_exports_it_instead()
    {
        var csv = In("sales.csv");
        File.WriteAllText(csv, "a,b\n1,2\n");
        var set = Run("set", csv, "/sheet[1]/cell[A1]", "--prop", "value=9");
        Assert.NotEqual(0, set.Code);
        var error = JsonDocument.Parse(set.Err).RootElement.GetProperty("error");
        Assert.Equal("FORMAT_READONLY", error.GetProperty("code").GetString());
        Assert.Contains(".xlsx", error.GetProperty("hint").GetString());
        Assert.Equal("a,b\n1,2\n", File.ReadAllText(csv));

        var export = Run("export", csv, "--to", In("sales.xlsx"));
        Assert.Equal(0, export.Code);
        Assert.True(File.Exists(In("sales.xlsx")));
        var get = Run("get", In("sales.xlsx"), "/sheet[1]/cell[B2]");
        Assert.Equal("2", JsonDocument.Parse(get.Out).RootElement.GetProperty("props").GetProperty("text").GetString());
    }

    static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "compat", name);

    [Fact]
    public void Doc_opens_as_a_word_document_and_exports_to_docx_unchanged()
    {
        using var doc = Adapters.ForPath("sample.doc").Open(File.OpenRead(Fixture("sample.doc")));
        Assert.Equal("doc", doc.Format);
        var body = Body(doc);
        Assert.Contains(body, b => b.Kind == "heading");
        Assert.Contains(body, b => b.Kind == "paragraph" && b.GetProps().GetValueOrDefault("list") == "bullet");
        Assert.Contains(body, b => b.Kind == "table");
        Assert.Contains(body.SelectMany(b => b.Children), r => r.Kind == "run" && r.GetProps().GetValueOrDefault("bold") == "true");
        var (target, _) = Exporter.Export(doc, new DocxAdapter(), In("sample.docx"));
        Assert.Same(((CompatDocument)doc).Inner, target);
        using (var fs = File.Create(In("sample.docx"))) target.Save(fs);
        using var again = new DocxAdapter().Open(File.OpenRead(In("sample.docx")));
        Assert.Equal(body.Count, Body(again).Count);
    }

    [Fact]
    public void Sheet_writer_applies_every_cell_style()
    {
        var warnings = new List<string>();
        var cell = new CellModel { Row = 2, Col = 3, Value = 1.5, SizePt = 14, Bold = true, Italic = true, Underline = true, Strike = true, Wrap = true, Color = "FF0000", Fill = "FFFF00", Format = "0.00", Align = "center", VAlign = "middle", Font = "Arial" };
        using var doc = Writers.Xlsx([new SheetModel { Name = "S", Cells = [cell], ColWidths = { [3] = 20 }, RowHeights = { [2] = 30 } }], warnings);
        Assert.Empty(warnings);
        var props = PathResolver.Single(doc.Root, "/sheet[1]/cell[C2]").GetProps();
        Assert.Equal("14", props["size"]);
        Assert.Equal("true", props["bold"]);
        Assert.Equal("FF0000", props["color"]);
        Assert.Equal("FFFF00", props["fill"]);
        Assert.Equal("0.00", props["format"]);
        Assert.Equal("center", props["align"]);
        Assert.Equal("middle", props["valign"]);
        Assert.Equal("Arial", props["font"]);
        Assert.Equal("1.5", props["text"]);
    }

    [Fact]
    public void Xls_opens_as_a_workbook_and_exports_to_xlsx()
    {
        using var doc = Adapters.ForPath("sample.xls").Open(File.OpenRead(Fixture("sample.xls")));
        Assert.Equal("xls", doc.Format);
        var sheets = doc.Root.Children.Where(c => c.Kind == "sheet").ToList();
        Assert.Equal(2, sheets.Count);
        Assert.Equal("Data", sheets[0].GetProps()["name"]);
        Assert.Equal("SUM(A1:A2)", PathResolver.Single(doc.Root, "/sheet[1]/cell[A3]").GetProps()["formula"]);
        Assert.Equal("3", Cell(doc, "A3"));
        Assert.Contains("A5:B6", sheets[0].GetProps()["merges"]);
        var (target, _) = Exporter.Export(doc, new XlsxAdapter(), In("sample.xlsx"));
        Assert.Equal("xlsx", target.Format);
        Assert.Same(((CompatDocument)doc).Inner, target);
    }

    [Fact]
    public void Ppt_opens_as_a_deck()
    {
        using var doc = Adapters.ForPath("sample.ppt").Open(File.OpenRead(Fixture("sample.ppt")));
        Assert.Equal("ppt", doc.Format);
        var slides = doc.Root.Children.Where(c => c.Kind == "slide").ToList();
        Assert.Single(slides);
        var shapes = slides[0].Children.Where(c => c.Kind != "decor").ToList();
        Assert.Contains(shapes, s => s.Kind == "shape" && s.Text!.Contains("Hello"));
        Assert.Contains(shapes, s => s.Kind == "image");
        Assert.Contains(shapes, s => s.Kind == "shape" && s.Children.Any(p => p.Kind == "paragraph" && p.GetProps().GetValueOrDefault("list") == "bullet"));
    }

    static IReadOnlyList<Node> Body(Document doc) => doc.Root.Children.First(c => c.Kind == "body").Children;

    [Fact]
    public void Rtf_from_textutil_keeps_runs_lists_and_tables()
    {
        using var doc = Adapters.ForPath("sample.rtf").Open(File.OpenRead(Fixture("sample.rtf")));
        Assert.Equal("rtf", doc.Format);
        var body = Body(doc);
        Assert.Equal("Compat 样例", body[0].Text);
        Assert.Equal("true", body[0].Children.First(c => c.Kind == "run").GetProps().GetValueOrDefault("bold"));
        var runs = body[1].Children.Where(c => c.Kind == "run").Select(r => r.GetProps()).ToList();
        Assert.Contains(runs, r => r["text"] == "bold" && r.GetValueOrDefault("bold") == "true");
        Assert.Contains(runs, r => r["text"] == "italic" && r.GetValueOrDefault("italic") == "true");
        Assert.Contains(runs, r => r["text"] == "red" && r.GetValueOrDefault("color") == "FB0007");
        Assert.Contains("中文段落。", body[1].Text);
        Assert.Equal(["bullet", "bullet", "number", "number"], body.Skip(2).Take(4).Select(b => b.GetProps().GetValueOrDefault("list")).ToList());
        Assert.Equal("first bullet", body[2].Text);
        var table = body.First(b => b.Kind == "table");
        var rows = table.Children.Where(c => c.Kind == "row").ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal(["A2", "B2"], rows[1].Children.Where(c => c.Kind == "cell").Select(c => c.Text).ToList());
        Assert.Equal("Last paragraph.", body[^1].Text);
    }

    [Fact]
    public void Rtf_headings_links_pictures_merges_and_gbk_bytes()
    {
        var png = Convert.ToHexString(FakePng(3, 3));
        var rtf = @"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fcharset0 Arial;}{\f1\fnil\fcharset134 SimSun;}}{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}
{\stylesheet{\s0 Normal;}{\s1\outlinelevel0 heading 1;}}
\pard\s1\b Chapter\par
\pard\plain\f1 \'d6\'d0\'ce\'c4\par
\pard{\field{\*\fldinst HYPERLINK ""https://example.com""}{\fldrslt Example}} after\par
\pard\qc centered\par
\pard\page after break\par
{\pict\pngblip\picwgoal1134\pichgoal1134 " + png + @"}
\trowd\clmgf\cellx3000\clmrg\cellx6000\cellx9000 \pard\intbl merged\cell\cell last\cell\row
\trowd\cellx3000\cellx6000\cellx9000 \pard\intbl a\cell b\cell c\cell\row
\pard tail\par}";
        using var doc = Open("note.rtf", rtf);
        var body = Body(doc);
        Assert.Equal("heading", body[0].Kind);
        Assert.Equal("Chapter", body[0].Text);
        Assert.Equal("中文", body[1].Text);
        Assert.Equal("https://example.com", body[2].Children.First(c => c.Kind == "run").GetProps().GetValueOrDefault("link"));
        Assert.Equal("Example after", body[2].Text);
        Assert.Equal("center", body[3].GetProps()["align"]);
        Assert.Equal("pagebreak", body[4].Kind);
        Assert.Equal("after break", body[5].Text);
        Assert.Equal("image", body[6].Kind);
        var table = body[7];
        var first = table.Children.First(c => c.Kind == "row").Children.Where(c => c.Kind == "cell").ToList();
        Assert.Equal(2, first.Count);
        Assert.Equal("2", first[0].GetProps()["colspan"]);
        Assert.Equal("merged", first[0].Text);
        Assert.Equal("tail", body[8].Text);
    }

    static byte[] OdfZip(string mime, string content, string? styles = null, (string Name, byte[] Bytes)[]? pictures = null)
    {
        var ms = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, true))
        {
            using (var w = new StreamWriter(zip.CreateEntry("mimetype").Open())) w.Write(mime);
            using (var w = new StreamWriter(zip.CreateEntry("content.xml").Open())) w.Write(content);
            if (styles is not null) using (var w = new StreamWriter(zip.CreateEntry("styles.xml").Open())) w.Write(styles);
            foreach (var (name, bytes) in pictures ?? []) using (var s = zip.CreateEntry(name).Open()) s.Write(bytes);
        }
        return ms.ToArray();
    }

    const string OdfNs = "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\"";

    [Fact]
    public void Odt_from_textutil_opens()
    {
        using var doc = Adapters.ForPath("sample.odt").Open(File.OpenRead(Fixture("sample.odt")));
        var body = Body(doc);
        Assert.Equal("Compat 样例", body[0].Text);
        Assert.Contains(body, b => b.Text == "first bullet" && b.GetProps().GetValueOrDefault("list") == "bullet");
        Assert.Contains(body, b => b.Text == "step two" && b.GetProps().GetValueOrDefault("list") == "number");
        var table = body.First(b => b.Kind == "table");
        Assert.Equal(["A1", "B1"], table.Children.First(c => c.Kind == "row").Children.Where(c => c.Kind == "cell").Select(c => c.Text).ToList());
    }

    [Fact]
    public void Odt_keeps_headings_styles_spans_tables_pictures_and_page_breaks()
    {
        var content = $$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document-content {{OdfNs}}>
            <office:font-face-decls><style:font-face style:name="F1" svg:font-family="'Noto Sans'"/></office:font-face-decls>
            <office:automatic-styles>
              <style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:text-align="center" fo:break-before="page"/></style:style>
              <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:color="#008000" fo:font-size="14pt" style:font-name="F1"/></style:style>
              <style:style style:name="C1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ffcc00"/></style:style>
              <text:list-style style:name="L1"><text:list-level-style-number text:level="1"/></text:list-style>
            </office:automatic-styles>
            <office:body><office:text>
              <text:h text:outline-level="2">Second <text:span text:style-name="T1">level</text:span></text:h>
              <text:p text:style-name="P1">Centered<text:s text:c="2"/>after<text:line-break/>break <text:a xlink:href="https://example.com">link</text:a></text:p>
              <text:list text:style-name="L1"><text:list-item><text:p>one</text:p><text:list><text:list-item><text:p>nested</text:p></text:list-item></text:list></text:list-item></text:list>
              <table:table><table:table-row><table:table-cell table:number-columns-spanned="2" table:style-name="C1"><text:p>wide</text:p></table:table-cell><table:covered-table-cell/></table:table-row>
              <table:table-row><table:table-cell><text:p>a</text:p><text:p>b</text:p></table:table-cell><table:table-cell><text:p>c</text:p></table:table-cell></table:table-row></table:table>
              <text:p><draw:frame svg:width="2cm" svg:height="1cm"><draw:image xlink:href="Pictures/p.png"/></draw:frame></text:p>
            </office:text></office:body></office:document-content>
            """;
        using var doc = Open("a.odt", OdfZip("application/vnd.oasis.opendocument.text", content, null, [("Pictures/p.png", FakePng(6, 3))]));
        var body = Body(doc);
        Assert.Equal("heading", body[0].Kind);
        Assert.Equal("2", body[0].GetProps()["level"]);
        var bold = body[0].Children.First(c => c.Kind == "run" && c.Text == "level").GetProps();
        Assert.Equal("true", bold["bold"]);
        Assert.Equal("008000", bold["color"]);
        Assert.Equal("14", bold["size"]);
        Assert.Equal("Noto Sans", bold["font"]);
        Assert.Equal("pagebreak", body[1].Kind);
        Assert.Equal("center", body[2].GetProps()["align"]);
        Assert.StartsWith("Centered  after", body[2].Text);
        Assert.Equal("https://example.com", body[2].Children.First(c => c.Kind == "run" && c.Text == "link").GetProps()["link"]);
        Assert.Equal("number", body[3].GetProps()["list"]);
        Assert.Equal("0", body[3].GetProps()["level"]);
        Assert.Equal("1", body[4].GetProps()["level"]);
        var rows = body[5].Children.Where(c => c.Kind == "row").ToList();
        var wide = rows[0].Children.First(c => c.Kind == "cell");
        Assert.Equal("2", wide.GetProps()["colspan"]);
        Assert.Equal("FFCC00", wide.GetProps()["fill"]);
        Assert.Contains("a", rows[1].Children.First(c => c.Kind == "cell").Text);
        Assert.Equal("image", body[^1].Kind);
        Assert.Equal("720000", body[^1].GetProps()["width"]);
    }

    [Fact]
    public void Ods_keeps_typed_cells_formulas_styles_merges_and_sizes()
    {
        var content = $$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document-content {{OdfNs}}>
            <office:automatic-styles>
              <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="4cm"/></style:style>
              <style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="1cm"/></style:style>
              <number:percentage-style style:name="N1"><number:number number:decimal-places="1"/><number:text>%</number:text></number:percentage-style>
              <style:style style:name="ce1" style:family="table-cell"><style:table-cell-properties fo:background-color="#ff0000"/><style:text-properties fo:font-weight="bold"/></style:style>
              <style:style style:name="ce2" style:family="table-cell" style:data-style-name="N1"/>
            </office:automatic-styles>
            <office:body><office:spreadsheet>
              <table:table table:name="Data">
                <table:table-column table:style-name="co1"/><table:table-column table:number-columns-repeated="2"/>
                <table:table-row table:style-name="ro1">
                  <table:table-cell office:value-type="string" table:style-name="ce1"><text:p>Name</text:p></table:table-cell>
                  <table:table-cell office:value-type="float" office:value="1.5"><text:p>1.5</text:p></table:table-cell>
                  <table:table-cell office:value-type="date" office:date-value="2026-02-03"><text:p>2026-02-03</text:p></table:table-cell>
                </table:table-row>
                <table:table-row>
                  <table:table-cell table:formula="of:=SUM([.B1:.B1];[.B1])" office:value-type="float" office:value="3"><text:p>3</text:p></table:table-cell>
                  <table:table-cell office:value-type="percentage" office:value="0.25" table:style-name="ce2"><text:p>25.0%</text:p></table:table-cell>
                  <table:table-cell office:value-type="boolean" office:boolean-value="true"><text:p>TRUE</text:p></table:table-cell>
                </table:table-row>
                <table:table-row>
                  <table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="1" office:value-type="string"><text:p>merged</text:p></table:table-cell><table:covered-table-cell/>
                  <table:table-cell office:value-type="string" table:number-columns-repeated="2"><text:p>rep</text:p></table:table-cell>
                </table:table-row>
                <table:table-row table:number-rows-repeated="1048000"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>
              </table:table>
              <table:table table:name="第二"><table:table-row><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row></table:table>
            </office:spreadsheet></office:body></office:document-content>
            """;
        using var doc = Open("a.ods", OdfZip("application/vnd.oasis.opendocument.spreadsheet", content));
        var sheets = doc.Root.Children.Where(c => c.Kind == "sheet").ToList();
        Assert.Equal(["Data", "第二"], sheets.Select(s => s.GetProps()["name"]).ToList());
        Assert.Equal("Name", Cell(doc, "A1"));
        Assert.Equal("true", PathResolver.Single(doc.Root, "/sheet[1]/cell[A1]").GetProps()["bold"]);
        Assert.Equal("FF0000", PathResolver.Single(doc.Root, "/sheet[1]/cell[A1]").GetProps()["fill"]);
        Assert.Equal("1.5", Cell(doc, "B1"));
        Assert.Equal("date", PathResolver.Single(doc.Root, "/sheet[1]/cell[C1]").GetProps()["type"]);
        Assert.Equal("SUM(B1:B1,B1)", PathResolver.Single(doc.Root, "/sheet[1]/cell[A2]").GetProps()["formula"]);
        Assert.Equal("0.0%", PathResolver.Single(doc.Root, "/sheet[1]/cell[B2]").GetProps()["format"]);
        Assert.Equal("TRUE", Cell(doc, "C2").ToUpperInvariant());
        Assert.Equal("rep", Cell(doc, "D3"));
        var sheet = sheets[0].GetProps();
        Assert.Contains("A3:B3", sheet["merges"]);
        Assert.Contains("\"A\":", sheet["widths"]);
        Assert.Contains("\"1\":", sheet["heights"]);
    }

    [Fact]
    public void Odp_keeps_slide_size_titles_bullets_pictures_and_shapes()
    {
        var styles = $$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document-styles {{OdfNs}}>
            <office:automatic-styles><style:page-layout style:name="PL1"><style:page-layout-properties fo:page-width="28cm" fo:page-height="15.75cm"/></style:page-layout></office:automatic-styles>
            <office:master-styles><style:master-page style:name="Default" style:page-layout-name="PL1"/></office:master-styles>
            </office:document-styles>
            """;
        var content = $$"""
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document-content {{OdfNs}}>
            <office:automatic-styles>
              <style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#336699" draw:stroke="none"/></style:style>
              <style:style style:name="P1" style:family="paragraph"><style:text-properties fo:font-size="32pt"/></style:style>
              <text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" text:bullet-char="•"/></text:list-style>
            </office:automatic-styles>
            <office:body><office:presentation>
              <draw:page draw:name="page1" draw:master-page-name="Default">
                <draw:frame presentation:class="title" svg:x="1cm" svg:y="1cm" svg:width="26cm" svg:height="3cm"><draw:text-box><text:p text:style-name="P1">Hello 演示</text:p></draw:text-box></draw:frame>
                <draw:frame presentation:class="outline" svg:x="1cm" svg:y="5cm" svg:width="12cm" svg:height="8cm"><draw:text-box><text:list text:style-name="L1"><text:list-item><text:p>first</text:p></text:list-item><text:list-item><text:p>second</text:p></text:list-item></text:list></draw:text-box></draw:frame>
                <draw:frame svg:x="14cm" svg:y="5cm" svg:width="4cm" svg:height="3cm"><draw:image xlink:href="Pictures/p.png"/></draw:frame>
                <draw:custom-shape draw:style-name="gr1" svg:x="20cm" svg:y="10cm" svg:width="5cm" svg:height="2cm"><text:p>box</text:p><draw:enhanced-geometry draw:type="round-rectangle"/></draw:custom-shape>
                <presentation:notes><draw:frame presentation:class="notes"><draw:text-box><text:p>speaker note</text:p></draw:text-box></draw:frame></presentation:notes>
              </draw:page>
              <draw:page draw:name="page2" draw:master-page-name="Default"><draw:frame svg:x="1cm" svg:y="1cm" svg:width="5cm" svg:height="1cm"><draw:text-box><text:p>second slide</text:p></draw:text-box></draw:frame></draw:page>
            </office:presentation></office:body></office:document-content>
            """;
        using var doc = Open("a.odp", OdfZip("application/vnd.oasis.opendocument.presentation", content, styles, [("Pictures/p.png", FakePng(6, 3))]));
        var root = doc.Root.GetProps();
        Assert.Equal("10080000", root["width"]);
        Assert.Equal("5670000", root["height"]);
        var slides = doc.Root.Children.Where(c => c.Kind == "slide").ToList();
        Assert.Equal(2, slides.Count);
        var shapes = slides[0].Children.Where(c => c.Kind != "decor").ToList();
        Assert.Equal(["shape", "shape", "image", "shape"], shapes.Select(s => s.Kind).ToList());
        Assert.Equal("Hello 演示", shapes[0].Text);
        Assert.Equal("360000", shapes[0].GetProps()["x"]);
        Assert.Equal("9360000", shapes[0].GetProps()["w"]);
        Assert.Equal("32", shapes[0].GetProps()["size"]);
        var bullets = shapes[1].Children.Where(c => c.Kind == "paragraph").ToList();
        Assert.Equal(["first", "second"], bullets.Select(b => b.Text).ToList());
        Assert.Equal("bullet", bullets[0].GetProps()["list"]);
        Assert.Equal("1440000", shapes[2].GetProps()["w"]);
        Assert.Equal("roundRect", shapes[3].GetProps()["geometry"]);
        Assert.Equal("336699", shapes[3].GetProps()["fill"]);
        Assert.Equal("box", shapes[3].Text);
        Assert.Equal("speaker note", slides[0].GetProps()["notes"]);
        Assert.Equal("second slide", slides[1].Children.First(c => c.Kind == "shape").Text);
    }
}
