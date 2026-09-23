using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>The blank 16:9 deck `create` starts from: one master, one theme, three layouts (Title, Content, Blank).</summary>
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
        foreach (var (id, xml) in new[] { ("rIdLayout1", TitleLayoutXml), ("rIdLayout2", ContentLayoutXml), ("rIdLayout3", BlankLayoutXml) })
        {
            var layout = master.AddNewPart<SlideLayoutPart>(id);
            layout.SlideLayout = new P.SlideLayout(xml);
            layout.AddPart(master);
        }
        master.SlideMaster = new P.SlideMaster(MasterXml);
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

    static string Placeholder(int id, string name, string type, int? index, string? xfrm, string bodyPr, string listStyle, string text)
    {
        var ph = index is null ? $"<p:ph type=\"{type}\"/>" : type == "body" ? $"<p:ph idx=\"{index}\"/>" : $"<p:ph type=\"{type}\" idx=\"{index}\"/>";
        var spPr = xfrm is null ? "<p:spPr/>" : $"<p:spPr>{xfrm}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>";
        return $"<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"{name}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>{ph}</p:nvPr></p:nvSpPr>{spPr}"
            + $"<p:txBody>{bodyPr}{listStyle}<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>";
    }

    static string Xfrm(long x, long y, long w, long h) => $"<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{w}\" cy=\"{h}\"/></a:xfrm>";

    static string MasterXml =>
        $"""
        <p:sldMaster {Ns}>
        <p:cSld>
        <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
        <p:spTree>
        {GroupHeader}
        {Placeholder(2, "Title Placeholder 1", "title", null, Xfrm(838200, 365125, 10515600, 1325563),
            "<a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"ctr\"><a:normAutofit/></a:bodyPr>", "<a:lstStyle/>", "Click to edit Master title style")}
        {Placeholder(3, "Text Placeholder 2", "body", 1, Xfrm(838200, 1825625, 10515600, 4351338),
            "<a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"><a:normAutofit/></a:bodyPr>", "<a:lstStyle/>", "Click to edit Master text styles")}
        </p:spTree>
        </p:cSld>
        <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
        <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdLayout1"/><p:sldLayoutId id="2147483650" r:id="rIdLayout2"/><p:sldLayoutId id="2147483651" r:id="rIdLayout3"/></p:sldLayoutIdLst>
        <p:txStyles>
        <p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>
        <p:bodyStyle>{BodyLevels()}</p:bodyStyle>
        <p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:otherStyle>
        </p:txStyles>
        </p:sldMaster>
        """;

    static string BodyLevels()
    {
        var sb = new StringBuilder();
        int[] sizes = [2800, 2400, 2000, 1800, 1800];
        for (var level = 1; level <= 5; level++)
        {
            var marL = 228600 + 457200 * (level - 1);
            sb.Append($"<a:lvl{level}pPr marL=\"{marL}\" indent=\"-228600\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">")
              .Append("<a:lnSpc><a:spcPct val=\"90000\"/></a:lnSpc><a:spcBef><a:spcPts val=\"1000\"/></a:spcBef>")
              .Append("<a:buFont typeface=\"Arial\" panose=\"020B0604020202020204\" pitchFamily=\"34\" charset=\"0\"/><a:buChar char=\"•\"/>")
              .Append($"<a:defRPr sz=\"{sizes[level - 1]}\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl{level}pPr>");
        }
        return sb.ToString();
    }

    static string Layout(string type, string name, string shapes) =>
        $"<p:sldLayout {Ns} type=\"{type}\" preserve=\"1\"><p:cSld name=\"{name}\"><p:spTree>{GroupHeader}{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>";

    static string TitleLayoutXml => Layout("title", "Title",
        Placeholder(2, "Title 1", "ctrTitle", null, Xfrm(1524000, 1122363, 9144000, 2387600),
            "<a:bodyPr anchor=\"b\"><a:normAutofit/></a:bodyPr>", "<a:lstStyle><a:lvl1pPr algn=\"ctr\"><a:defRPr sz=\"6000\"/></a:lvl1pPr></a:lstStyle>", "Click to edit Master title style")
        + Placeholder(3, "Subtitle 2", "subTitle", 1, Xfrm(1524000, 3602038, 9144000, 1655762),
            "<a:bodyPr><a:normAutofit/></a:bodyPr>", "<a:lstStyle><a:lvl1pPr marL=\"0\" indent=\"0\" algn=\"ctr\"><a:buNone/><a:defRPr sz=\"2400\"/></a:lvl1pPr></a:lstStyle>", "Click to edit Master subtitle style"));

    static string ContentLayoutXml => Layout("obj", "Content",
        Placeholder(2, "Title 1", "title", null, null, "<a:bodyPr/>", "<a:lstStyle/>", "Click to edit Master title style")
        + Placeholder(3, "Content Placeholder 2", "body", 1, null, "<a:bodyPr/>", "<a:lstStyle/>", "Click to edit Master text styles"));

    static string BlankLayoutXml => Layout("blank", "Blank", "");

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

    /// <summary>The notes master PowerPoint creates on first use: slide image on top, notes text below, on a portrait page.</summary>
    public static string NotesMasterXml =>
        $"""
        <p:notesMaster {Ns}>
        <p:cSld>
        <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
        <p:spTree>
        {GroupHeader}
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr>{Xfrm(381000, 685800, 6096000, 3429000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>
        {Placeholder(3, "Notes Placeholder 2", "body", 3, Xfrm(685800, 4343400, 5486400, 4114800),
            "<a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/>", "<a:lstStyle/>", "Click to edit Master text styles")}
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

    /// <summary>An empty slide bound to a layout, with the layout's placeholders copied over as empty shapes.</summary>
    public static P.Slide NewSlide(SlideLayoutPart layout)
    {
        var slide = new P.Slide($"<p:sld {Ns}><p:cSld><p:spTree>{GroupHeader}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
        var tree = slide.CommonSlideData!.ShapeTree!;
        uint id = 2;
        foreach (var source in layout.SlideLayout?.CommonSlideData?.ShapeTree?.Elements<P.Shape>() ?? [])
        {
            var ph = source.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
            if (ph is null) continue;
            var name = source.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "Placeholder";
            tree.Append(new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = id++, Name = name },
                    new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new P.ApplicationNonVisualDrawingProperties((P.PlaceholderShape)ph.CloneNode(true))),
                new P.ShapeProperties(),
                new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" }))));
        }
        return slide;
    }
}
