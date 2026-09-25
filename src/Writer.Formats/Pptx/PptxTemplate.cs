using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>The blank 16:9 deck `create` starts from: one master, Office's theme, and the layouts PowerPoint users expect
/// (<see cref="Layouts"/>). A standard layout a deck lacks is made from the same specs on demand (<see cref="LayoutXml"/>).</summary>
public static class PptxTemplate
{
    public const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public const string RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string PNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
    public const string Ns = $"xmlns:a=\"{ANs}\" xmlns:r=\"{RNs}\" xmlns:p=\"{PNs}\"";

    public const long SlideWidth = 12192000;
    public const long SlideHeight = 6858000;

    public static void WriteBlank(Stream stream)
    {
        using var package = PresentationDocument.Create(stream, PresentationDocumentType.Presentation);
        var presentation = package.AddPresentationPart();
        var master = presentation.AddNewPart<SlideMasterPart>("rIdMaster");
        var theme = master.AddNewPart<ThemePart>("rIdTheme");
        theme.Theme = new A.Theme(ThemeXml);
        presentation.AddPart(theme);
        var ids = new StringBuilder();
        for (var i = 0; i < Layouts.Length; i++)
        {
            var layout = master.AddNewPart<SlideLayoutPart>("rIdLayout" + (i + 1));
            layout.SlideLayout = new P.SlideLayout(LayoutXml(Layouts[i], SlideWidth, SlideHeight));
            layout.AddPart(master);
            ids.Append($"<p:sldLayoutId id=\"{2147483649L + i}\" r:id=\"rIdLayout{i + 1}\"/>");
        }
        master.SlideMaster = new P.SlideMaster(MasterXml(ids.ToString()));
        presentation.Presentation = new P.Presentation(PresentationXml);
    }

    const string PresentationXml =
        $"""
        <p:presentation {Ns} saveSubsetFonts="1">
        <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdMaster"/></p:sldMasterIdLst>
        <p:sldIdLst/>
        <p:sldSz cx="12192000" cy="6858000"/>
        <p:notesSz cx="6858000" cy="9144000"/>
        <p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>
        </p:presentation>
        """;

    const string GroupHeader =
        """
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
        """;

    // ---------- layouts ----------

    /// <summary>A box on the 1600 × 900 grid the slide editor works in, scaled to the deck's size when written.</summary>
    public readonly record struct Box(int X, int Y, int W, int H);

    /// <summary>A layout placeholder. Type null is a content placeholder (&lt;p:ph idx/&gt;, PowerPoint's obj); Size in points,
    /// Color a scheme colour; Bullets false writes buNone and no indent.</summary>
    public sealed record Ph(string? Type, int? Idx, Box Box, string Anchor = "t", int? Size = null, bool? Bold = null, bool Bullets = true, string? Color = null, string? Sz = null);

    /// <summary>A standard layout: the key FindLayout takes (the slide editor's name for it), its OOXML type and PowerPoint's
    /// English name, what a Chinese PowerPoint calls it, its placeholders in drawing order and its own graphics.
    /// ui/office-io.js LAYOUT_SPECS draws the same boxes: keep the two in step.</summary>
    public sealed record LayoutSpec(string Key, string Type, string Name, string Chinese, Ph[] Placeholders, Decor[]? Graphics = null);

    /// <summary>A layout's own graphic, drawn under the placeholders on every slide of the layout: a short accent rule (bar)
    /// or a large opening quotation mark (quote), both in the first accent colour.</summary>
    public sealed record Decor(string Kind, Box Box);

    static readonly Box TitleBox = new(128, 56, 1344, 140), BodyBox = new(128, 236, 1344, 584);
    static readonly Ph Title = new("title", null, TitleBox, "b");

    public static readonly LayoutSpec[] Layouts =
    [
        new("title", "title", "Title Slide", "标题幻灯片",
            [new("ctrTitle", null, new(128, 200, 1344, 260), "b", 48), new("subTitle", 1, new(128, 508, 1344, 120), "t", 24, Bullets: false, Color: "tx2")],
            [new("bar", new(128, 482, 96, 6))]),
        new("content", "obj", "Title and Content", "标题和内容", [Title, new(null, 1, BodyBox)]),
        new("section", "secHead", "Section Header", "节标题",
            [new("title", null, new(128, 250, 1344, 250), "b", 44), new("body", 1, new(128, 548, 1344, 120), "t", 20, Bullets: false, Color: "tx2")],
            [new("bar", new(128, 522, 96, 6))]),
        new("two", "twoObj", "Two Content", "两栏内容", [Title, new(null, 1, new(128, 236, 640, 584), Size: 20, Sz: "half"), new(null, 2, new(832, 236, 640, 584), Size: 20, Sz: "half")]),
        new("comparison", "twoTxTwoObj", "Comparison", "比较",
        [
            Title,
            new("body", 1, new(128, 236, 640, 72), "b", 24, true, false),
            new(null, 2, new(128, 324, 640, 496), Size: 20, Sz: "half"),
            new("body", 3, new(832, 236, 640, 72), "b", 24, true, false, Sz: "quarter"),
            new(null, 4, new(832, 324, 640, 496), Size: 20, Sz: "quarter"),
        ]),
        new("titleOnly", "titleOnly", "Title Only", "仅标题", [Title]),
        new("blank", "blank", "Blank", "空白", []),
        new("caption", "objTx", "Content with Caption", "内容与标题",
            [new("title", null, new(128, 96, 480, 220), "b", 28), new(null, 1, new(672, 96, 800, 724)), new("body", 2, new(128, 348, 480, 472), "t", 16, Bullets: false, Sz: "half")]),
        new("picture", "picTx", "Picture with Caption", "图片与标题",
            [new("title", null, new(128, 96, 480, 220), "b", 28), new("pic", 1, new(672, 96, 800, 724), Bullets: false), new("body", 2, new(128, 348, 480, 472), "t", 16, Bullets: false, Sz: "half")]),
        new("quote", "cust", "Quote", "引用",
            [new("title", null, new(192, 300, 1216, 260), "t", 36, false), new("body", 1, new(192, 600, 1216, 80), "t", 20, Bullets: false, Color: "tx2")],
            [new("quote", new(168, 130, 240, 200))]),
    ];

    /// <summary>The standard layout a name stands for: its key, the key's aliases (the OOXML type), PowerPoint's English name or
    /// the Chinese one, ignoring case. Null for any other name.</summary>
    public static LayoutSpec? Standard(string name)
    {
        var n = name.Trim();
        if (n.Equals("compare", StringComparison.OrdinalIgnoreCase)) n = "comparison";
        return Layouts.FirstOrDefault(l => new[] { l.Key, l.Name, l.Chinese }.Any(x => x.Equals(n, StringComparison.OrdinalIgnoreCase))
            || l.Type != "cust" && l.Type.Equals(n, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>The p:sldLayout of a standard layout for a deck of this size.</summary>
    public static string LayoutXml(LayoutSpec spec, long width, long height)
    {
        var g = new Grid(width, height);
        var shapes = new StringBuilder();
        var id = 2;
        foreach (var ph in spec.Placeholders) shapes.Append(Placeholder(id++, g, ph));
        foreach (var decor in spec.Graphics ?? []) shapes.Append(decor.Kind == "quote" ? QuoteMark(id++, g, decor.Box) : Bar(id++, g, decor.Box));
        var xml = shapes.ToString();
        var type = spec.Type == "cust" ? "" : $" type=\"{spec.Type}\"";
        return $"<p:sldLayout {Ns}{type} preserve=\"1\"><p:cSld name=\"{spec.Name}\"><p:spTree>{GroupHeader}{xml}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>";
    }

    readonly record struct Grid(long Width, long Height)
    {
        public long X(int units) => (long)Math.Round(units * (Width / 1600.0));
        public long Y(int units) => (long)Math.Round(units * (Height / 900.0));
    }

    static string Xfrm(Grid g, Box b) => $"<a:xfrm><a:off x=\"{g.X(b.X)}\" y=\"{g.Y(b.Y)}\"/><a:ext cx=\"{g.X(b.W)}\" cy=\"{g.Y(b.H)}\"/></a:xfrm>";

    /// <summary>A short accent rule (layout graphics are drawn on every slide of the layout, under its placeholders).</summary>
    static string Bar(int id, Grid g, Box at) =>
        $"<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"Accent Bar\"/><p:cNvSpPr/><p:nvPr userDrawn=\"1\"/></p:nvSpPr>"
        + $"<p:spPr>{Xfrm(g, at)}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>"
        + "<p:txBody><a:bodyPr rtlCol=\"0\" anchor=\"ctr\"/><a:lstStyle/><a:p><a:endParaRPr lang=\"zh-CN\"/></a:p></p:txBody></p:sp>";

    /// <summary>A large opening quotation mark in the accent colour and the heading font.</summary>
    static string QuoteMark(int id, Grid g, Box at) =>
        $"<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"Quote Mark\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr userDrawn=\"1\"/></p:nvSpPr>"
        + $"<p:spPr>{Xfrm(g, at)}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        + "<p:txBody><a:bodyPr wrap=\"none\" lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\" rtlCol=\"0\" anchor=\"t\"/><a:lstStyle/>"
        + "<a:p><a:r><a:rPr lang=\"zh-CN\" sz=\"12000\"><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill><a:latin typeface=\"+mj-lt\"/><a:ea typeface=\"+mj-ea\"/></a:rPr><a:t>“</a:t></a:r></a:p></p:txBody></p:sp>";

    static readonly Dictionary<string, string> Names = new()
    {
        ["ctrTitle"] = "Title", ["title"] = "Title", ["subTitle"] = "Subtitle", ["body"] = "Text Placeholder", ["pic"] = "Picture Placeholder",
    };

    static readonly Dictionary<string, string> Prompts = new()
    {
        ["ctrTitle"] = "Click to edit Master title style", ["title"] = "Click to edit Master title style", ["subTitle"] = "Click to edit Master subtitle style",
        ["pic"] = "Click icon to add picture",
    };

    static string Placeholder(int id, Grid g, Ph ph)
    {
        var attrs = (ph.Type is null ? "" : $" type=\"{ph.Type}\"") + (ph.Sz is null ? "" : $" sz=\"{ph.Sz}\"") + (ph.Idx is null ? "" : $" idx=\"{ph.Idx}\"");
        var name = (ph.Type is null ? "Content Placeholder" : Names[ph.Type]) + " " + (id - 1);
        var defRPr = ph.Size is null && ph.Bold is null && ph.Color is null ? "" :
            "<a:defRPr" + (ph.Size is { } size ? $" sz=\"{size * 100}\"" : "") + (ph.Bold is { } bold ? $" b=\"{(bold ? 1 : 0)}\"" : "") + ">"
            + (ph.Color is { } color ? $"<a:solidFill><a:schemeClr val=\"{color}\"/></a:solidFill>" : "") + "</a:defRPr>";
        var lvl1 = (ph.Bullets ? "<a:lvl1pPr>" : "<a:lvl1pPr marL=\"0\" indent=\"0\"><a:buNone/>") + defRPr + "</a:lvl1pPr>";
        var list = lvl1 == "<a:lvl1pPr></a:lvl1pPr>" ? "<a:lstStyle/>" : $"<a:lstStyle>{lvl1}</a:lstStyle>";
        var prompt = ph.Type is not null && Prompts.TryGetValue(ph.Type, out var p) ? p : "Click to edit Master text styles";
        return $"<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"{name}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr>"
            + $"<p:spPr>{Xfrm(g, ph.Box)}</p:spPr>"
            + $"<p:txBody><a:bodyPr anchor=\"{ph.Anchor}\"><a:normAutofit/></a:bodyPr>{list}<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>{prompt}</a:t></a:r></a:p></p:txBody></p:sp>";
    }

    // ---------- master ----------

    static string MasterXml(string layoutIds)
    {
        var g = new Grid(SlideWidth, SlideHeight);
        return $"""
        <p:sldMaster {Ns}>
        <p:cSld>
        <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
        <p:spTree>
        {GroupHeader}
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr>{Xfrm(g, TitleBox)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="b"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master title style</a:t></a:r></a:p></p:txBody></p:sp>
        <p:sp><p:nvSpPr><p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr>{Xfrm(g, BodyBox)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>
        </p:spTree>
        </p:cSld>
        <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
        <p:sldLayoutIdLst>{layoutIds}</p:sldLayoutIdLst>
        <p:txStyles>
        <p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="3600" b="1" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>
        <p:bodyStyle>{BodyLevels()}</p:bodyStyle>
        <p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:otherStyle>
        </p:txStyles>
        </p:sldMaster>
        """;
    }

    static string BodyLevels()
    {
        var sb = new StringBuilder();
        int[] sizes = [2400, 2000, 1800, 1600, 1600];
        for (var level = 1; level <= 5; level++)
        {
            var marL = 228600 + 457200 * (level - 1);
            sb.Append($"<a:lvl{level}pPr marL=\"{marL}\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">")
              .Append("<a:lnSpc><a:spcPct val=\"100000\"/></a:lnSpc><a:spcBef><a:spcPts val=\"1000\"/></a:spcBef>")
              .Append("<a:buFont typeface=\"Arial\" panose=\"020B0604020202020204\" pitchFamily=\"34\" charset=\"0\"/><a:buChar char=\"•\"/>")
              .Append($"<a:defRPr sz=\"{sizes[level - 1]}\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl{level}pPr>");
        }
        return sb.ToString();
    }

    // ---------- theme and palettes ----------

    const string ThemeXml =
        $"""
        <a:theme xmlns:a="{ANs}" name="Writer">
        <a:themeElements>
        <a:clrScheme name="Writer">
        <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
        <a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
        <a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
        <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Writer">
        <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface=""/></a:majorFont>
        <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface=""/></a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Writer">
        <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
        <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
        <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
        <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
        </a:fmtScheme>
        </a:themeElements>
        <a:objectDefaults/><a:extraClrSchemeLst/>
        </a:theme>
        """;

    /// <summary>A colour and font scheme the slide editor offers under 设计 (ui/office-io.js THEMES, same keys and colours):
    /// text on background, secondary text, a card tone, six accents (the first is the one shapes and accent rules take),
    /// the heading and body fonts. Dark palettes keep dk1 dark and swap the master's colour map.</summary>
    public sealed record Palette(string Key, string Name, bool Dark, string Bg, string Fg, string Sub, string Card, string[] Accents, string Heading, string Body);

    public static readonly Palette[] Palettes =
    [
        new("paper", "Paper", false, "FFFFFF", "1D1D1F", "6E6E73", "F5F5F7", ["1D1D1F", "B5563A", "2F5D8A", "C9A227", "7A5A45", "8E8E93"], "Noto Serif SC", "Noto Sans SC"),
        new("ink", "Ink", true, "1D1D1F", "FFFFFF", "BDB7AA", "2C2C2E", ["E3B25A", "D98C5F", "8FB3D9", "C9C2B2", "A36F4B", "6E6E73"], "Noto Serif SC", "Noto Sans SC"),
        new("sea", "Sea", true, "1F3550", "FFFFFF", "C5D2E0", "2A4666", ["6CC6D9", "F2C14E", "9DB4CF", "E88D67", "5B8DB8", "A3B8CC"], "Noto Sans SC", "Noto Sans SC"),
        new("clay", "Clay", false, "F3E6DA", "3A2618", "7A5A45", "EAD5C3", ["B5563A", "D9A441", "6D8BA6", "8C6A4F", "C98C6B", "A99985"], "Noto Serif SC", "Noto Sans SC"),
        new("mist", "Mist", false, "F4F6F8", "1F2A37", "5B6573", "E4E9EF", ["3E6FB0", "E0913A", "7A8CA3", "C4574B", "5AA3C8", "9AA5B1"], "Noto Sans SC", "Noto Sans SC"),
        new("sand", "Sand", false, "F7F3EC", "2B2620", "7A6E5F", "ECE4D6", ["C2833A", "4F6D8F", "A65D4E", "8C7B5E", "D6A85C", "6F665A"], "Noto Serif SC", "Noto Sans SC"),
        new("rose", "Rose", false, "FFFFFF", "1D1D1F", "6E6E73", "F7ECEC", ["C4383C", "1D1D1F", "E08A6D", "8E5C6B", "D9A6A0", "8E8E93"], "Noto Sans SC", "Noto Sans SC"),
        new("night", "Night", true, "0F1B2D", "F5F7FA", "9FB0C7", "1B2A40", ["5AC8FA", "FFB547", "8E9CFF", "FF7A7A", "C792EA", "B8C4D6"], "Noto Sans SC", "Noto Sans SC"),
    ];

    public static Palette? FindPalette(string key) => Palettes.FirstOrDefault(p => p.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

    /// <summary>The palette's a:clrScheme, named "Writer &lt;Name&gt;" so the deck says which palette it wears.</summary>
    internal static string ColorSchemeXml(Palette p)
    {
        var (dk1, lt1, dk2, lt2) = p.Dark ? (p.Bg, p.Fg, p.Card, p.Sub) : (p.Fg, p.Bg, p.Sub, p.Card);
        var sb = new StringBuilder($"<a:clrScheme xmlns:a=\"{ANs}\" name=\"Writer {p.Name}\">");
        foreach (var (slot, color) in new[] { ("dk1", dk1), ("lt1", lt1), ("dk2", dk2), ("lt2", lt2) }) sb.Append($"<a:{slot}><a:srgbClr val=\"{color}\"/></a:{slot}>");
        for (var i = 0; i < 6; i++) sb.Append($"<a:accent{i + 1}><a:srgbClr val=\"{p.Accents[i]}\"/></a:accent{i + 1}>");
        return sb.Append($"<a:hlink><a:srgbClr val=\"{p.Accents[0]}\"/></a:hlink><a:folHlink><a:srgbClr val=\"{p.Sub}\"/></a:folHlink></a:clrScheme>").ToString();
    }

    internal static string FontSchemeXml(string name, string heading, string body) =>
        $"<a:fontScheme xmlns:a=\"{ANs}\" name=\"{System.Security.SecurityElement.Escape(name)}\">"
        + $"<a:majorFont><a:latin typeface=\"{System.Security.SecurityElement.Escape(heading)}\"/><a:ea typeface=\"{System.Security.SecurityElement.Escape(heading)}\"/><a:cs typeface=\"\"/></a:majorFont>"
        + $"<a:minorFont><a:latin typeface=\"{System.Security.SecurityElement.Escape(body)}\"/><a:ea typeface=\"{System.Security.SecurityElement.Escape(body)}\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>";

    // ---------- notes ----------

    static string NotesPlaceholder(int id, string name, int index, string xfrm) =>
        $"<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"{name}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" idx=\"{index}\"/></p:nvPr></p:nvSpPr><p:spPr>{xfrm}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>"
        + "<p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>";

    static string EmuXfrm(long x, long y, long w, long h) => $"<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{w}\" cy=\"{h}\"/></a:xfrm>";

    /// <summary>The notes master PowerPoint creates on first use: slide image on top, notes text below, on a portrait page.</summary>
    public static string NotesMasterXml =>
        $"""
        <p:notesMaster {Ns}>
        <p:cSld>
        <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
        <p:spTree>
        {GroupHeader}
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr>{EmuXfrm(381000, 685800, 6096000, 3429000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>
        {NotesPlaceholder(3, "Notes Placeholder 2", 3, EmuXfrm(685800, 4343400, 5486400, 4114800))}
        </p:spTree>
        </p:cSld>
        <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
        <p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle>
        </p:notesMaster>
        """;

    /// <summary>An empty notes page for one slide: the slide image placeholder and the notes text placeholder.</summary>
    public static P.NotesSlide NewNotesSlide() => new(
        $"""
        <p:notes {Ns}>
        <p:cSld><p:spTree>
        {GroupHeader}
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
        <p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>
        </p:spTree></p:cSld>
        <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
        </p:notes>
        """);

    // ---------- slides ----------

    /// <summary>Placeholders PowerPoint leaves off a new slide unless 页眉页脚 turns them on.</summary>
    internal static bool IsFooter(P.PlaceholderShape ph) => ph.Type?.InnerText is "dt" or "ftr" or "sldNum" or "hdr" or "sldImg";

    /// <summary>An empty slide placeholder standing in for a layout's: the same p:ph (type, size, index), no text of its own, so
    /// PowerPoint shows its prompt and the layout's position and look apply.</summary>
    internal static P.Shape EmptyPlaceholder(uint id, string name, P.PlaceholderShape layoutPh)
    {
        var ph = (P.PlaceholderShape)layoutPh.CloneNode(true);
        ph.HasCustomPrompt = null;
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(ph)),
            new P.ShapeProperties(),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" })));
    }

    /// <summary>The placeholders of a layout a slide stands in for, in drawing order (footer ones left out).</summary>
    internal static IEnumerable<(P.PlaceholderShape Ph, string Name)> LayoutPlaceholders(SlideLayoutPart layout) =>
        (layout.SlideLayout?.CommonSlideData?.ShapeTree?.Elements<P.Shape>() ?? [])
            .Select(s => (Ph: s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape, Name: s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Placeholder"))
            .Where(x => x.Ph is not null && !IsFooter(x.Ph))
            .Select(x => (x.Ph!, x.Name));

    /// <summary>An empty slide bound to a layout, with the layout's placeholders copied over as empty shapes.</summary>
    public static P.Slide NewSlide(SlideLayoutPart layout)
    {
        var slide = new P.Slide($"<p:sld {Ns}><p:cSld><p:spTree>{GroupHeader}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
        var tree = slide.CommonSlideData!.ShapeTree!;
        uint id = 2;
        foreach (var (ph, name) in LayoutPlaceholders(layout)) tree.Append(EmptyPlaceholder(id++, name, ph));
        return slide;
    }
}
