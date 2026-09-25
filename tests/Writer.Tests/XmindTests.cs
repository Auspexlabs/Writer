using System.IO.Compression;
using System.Text;
using Writer.Cli;
using Writer.Core;
using Writer.Formats;
using Writer.Formats.MindMap;

namespace Writer.Tests;

public class XmindTests
{
    const string Json = """
        [{"id":"s1","class":"sheet","title":"Sheet 1","rootTopic":{"id":"r1","class":"topic","title":"Plan","structureClass":"org.xmind.ui.logic.right",
          "children":{"attached":[
            {"id":"a1","title":"Alpha","href":"https://x.y","labels":["urgent"," Q4 "],"markers":[{"markerId":"priority-1"},{"markerId":"flag-red"},{"markerId":"task-half"},{"markerId":"people-blue"}],
             "notes":{"plain":{"content":"note here"}},"style":{"id":"st1","properties":{"svg:fill":"#ffff00","fo:font-weight":"bold","fo:font-size":"18pt"}},
             "children":{"attached":[{"id":"a11","title":"A one"},{"id":"a12","title":"A two"}],"summary":[{"id":"s1t","title":"Sum"}]},
             "boundaries":[{"id":"b1","range":"(0,1)"}],"summaries":[{"id":"su1","range":"(0,1)","topicId":"s1t"}]},
            {"id":"b1","title":"Beta","branch":"folded","image":{"src":"xap:resources/p.png","width":10,"height":5},"children":{"attached":[{"id":"b11","title":"deep","href":"xmind:#a1"}]}}],
          "detached":[{"id":"f1","title":"Float","position":{"x":300.4,"y":-120}}]}},
          "relationships":[{"id":"rel1","end1Id":"a1","end2Id":"b1","title":"leads"},{"id":"rel2","end1Id":"a1","end2Id":"nope"}]},
         {"id":"s2","class":"sheet","title":"Sheet 2","rootTopic":{"id":"r2","title":"Other"}}]
        """;

    const string Xml = """
        <?xml version="1.0" encoding="UTF-8" standalone="no"?>
        <xmap-content xmlns="urn:xmind:xmap:xmlns:content:2.0" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:svg="http://www.w3.org/2000/svg" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:xlink="http://www.w3.org/1999/xlink" version="2.0">
        <sheet id="sh1"><topic id="r1" structure-class="org.xmind.ui.map.unbalanced"><title>Root</title>
          <children><topics type="attached">
            <topic id="a1" style-id="st1" branch="folded"><title>Alpha</title><notes><plain>xml note</plain></notes><labels><label>tag</label></labels><marker-refs><marker-ref marker-id="star-blue"/></marker-refs>
              <xhtml:img xhtml:src="xap:attachments/i.jpg" svg:width="20" svg:height="10"/>
              <children><topics type="attached"><topic id="a11"><title>A one</title></topic><topic id="a12"><title>A two</title></topic></topics><topics type="summary"><topic id="sx"><title>Sum</title></topic></topics></children>
              <summaries><summary range="(1,1)" topic-id="sx"/></summaries></topic>
            <topic id="b1" xlink:href="https://b.example"><title>Beta</title></topic>
          </topics><topics type="detached"><topic id="f1"><title>Free</title><position svg:x="-200" svg:y="80.6"/></topic></topics></children>
        </topic>
        <relationships><relationship id="rl" end1="b1" end2="a1"><title>back</title></relationship></relationships>
        </sheet></xmap-content>
        """;

    const string Styles = """
        <?xml version="1.0" encoding="UTF-8" standalone="no"?>
        <xmap-styles xmlns="urn:xmind:xmap:xmlns:style:2.0" xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:svg="http://www.w3.org/2000/svg" version="2.0">
        <styles><style id="st1" type="topic"><topic-properties svg:fill="#112233" fo:color="#ABCDEF" fo:font-style="italic"/></style></styles>
        </xmap-styles>
        """;

    static MemoryStream Zip(params (string Name, byte[] Bytes)[] entries)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
            foreach (var (name, bytes) in entries)
            {
                using var w = zip.CreateEntry(name).Open();
                w.Write(bytes);
            }
        ms.Position = 0;
        return ms;
    }

    static byte[] U(string s) => Encoding.UTF8.GetBytes(s);

    static Node At(Document doc, string path) => PathResolver.Single(doc.Root, path);

    [Fact]
    public void Content_json_opens_as_a_map_with_everything_xmind_puts_on_a_topic()
    {
        using var doc = new XmindAdapter().Open(Zip(("content.json", U(Json)), ("resources/p.png", U("PNG"))));
        Assert.Equal("Plan\n- Alpha\n  - A one\n  - A two\n  - Sum\n- Beta\n  - deep\n- Float\n", Views.Text(doc.Root));
        Assert.Equal("logic", At(doc, "/topic[1]").GetProps()["structure"]);
        var alpha = At(doc, "//topic[@id=a1]").GetProps();
        Assert.Equal(("https://x.y", "urgent,Q4", "full-1,flag,50%", "note here"), (alpha["link"], alpha["labels"], alpha["icon"], alpha["note"]));
        Assert.Equal(("FFFF00", "true", "18"), (alpha["fill"], alpha["bold"], alpha["size"]));
        Assert.Equal("[{\"to\":\"b1\",\"label\":\"leads\",\"color\":\"\",\"arrows\":\"end\"}]", alpha["rels"]);
        Assert.Equal("CFE2F3", At(doc, "//topic[@id=a11]").GetProps()["cloud"]);
        Assert.Equal("CFE2F3", At(doc, "//topic[@id=a12]").GetProps()["cloud"]);
        Assert.Equal("a11:a12", At(doc, "//topic[@id=s1t]").GetProps()["summary"]);
        var beta = At(doc, "//topic[@id=b1]").GetProps();
        Assert.Equal(("true", "data:image/png;base64,UE5H", "10,5"), (beta["collapsed"], beta["image"], beta["imageSize"]));
        Assert.Equal("[{\"to\":\"a1\",\"label\":\"\",\"color\":\"\",\"arrows\":\"end\"}]", At(doc, "//topic[@id=b11]").GetProps()["rels"]);
        Assert.False(At(doc, "//topic[@id=b11]").GetProps().ContainsKey("link"));
        Assert.Equal("300,-120", At(doc, "//topic[@id=f1]").GetProps()["free"]);

        var (target, warnings) = Exporter.Export(doc, new MmAdapter(), null);
        Assert.Same(doc, target);
        Assert.Empty(warnings);
        Assert.Equal("mm", Adapters.ForPath("x.xmind").Open(Zip(("content.json", U(Json)))).Format);
        Assert.Equal(ErrorCode.FormatReadonly, Assert.Throws<WriterException>(() => new XmindAdapter().Create()).Code);
        Assert.Equal(ErrorCode.FormatError, Assert.Throws<WriterException>(() => new XmindAdapter().Open(new MemoryStream(U("nope")))).Code);
        Assert.Equal(ErrorCode.FormatError, Assert.Throws<WriterException>(() => new XmindAdapter().Open(Zip(("other.txt", U("x"))))).Code);
    }

    [Fact]
    public void Content_xml_with_styles_opens_the_same_way()
    {
        using var doc = new XmindAdapter().Open(Zip(("content.xml", U(Xml)), ("styles.xml", U(Styles)), ("attachments/i.jpg", U("JPG"))));
        Assert.Equal("Root\n- Alpha\n  - A one\n  - A two\n  - Sum\n- Beta\n- Free\n", Views.Text(doc.Root));
        Assert.False(At(doc, "/topic[1]").GetProps().ContainsKey("structure"));
        var alpha = At(doc, "//topic[@id=a1]").GetProps();
        Assert.Equal(("xml note", "tag", "star-blue", "true"), (alpha["note"], alpha["labels"], alpha["icon"], alpha["collapsed"]));
        Assert.Equal(("112233", "ABCDEF", "true", "data:image/jpeg;base64,SlBH", "20,10"), (alpha["fill"], alpha["color"], alpha["italic"], alpha["image"], alpha["imageSize"]));
        Assert.Equal("a12:a12", At(doc, "//topic[@id=sx]").GetProps()["summary"]);
        Assert.Equal("https://b.example", At(doc, "//topic[@id=b1]").GetProps()["link"]);
        Assert.Contains("\"to\":\"a1\",\"label\":\"back\"", At(doc, "//topic[@id=b1]").GetProps()["rels"]);
        Assert.Equal("-200,81", At(doc, "//topic[@id=f1]").GetProps()["free"]);
    }

    [Fact]
    public void Cli_exports_an_xmind_file_to_a_freemind_map_that_reopens()
    {
        var dir = Directory.CreateTempSubdirectory("writer-xmind").FullName;
        try
        {
            var src = Path.Combine(dir, "plan.xmind");
            File.WriteAllBytes(src, Zip(("content.json", U(Json))).ToArray());
            var to = Path.Combine(dir, "plan.mm");
            Assert.Equal(0, Runner.Run(["export", src, "--to", to], new StringWriter(), new StringWriter()));
            var xml = File.ReadAllText(to);
            Assert.StartsWith("<map version=\"1.0.1\">", xml);
            Assert.Contains("<node TEXT=\"Alpha\" ID=\"a1\"", xml);
            Assert.Contains("LINK=\"https://x.y\"", xml);
            Assert.Contains("<arrowlink DESTINATION=\"b1\"", xml);
            using var stream = File.OpenRead(to);
            using var again = Adapters.ForPath(to).Open(stream);
            Assert.Equal("Plan\n- Alpha\n  - A one\n  - A two\n  - Sum\n- Beta\n  - deep\n- Float\n", Views.Text(again.Root));
            Assert.Equal("urgent,Q4", At(again, "//topic[@id=a1]").GetProps()["labels"]);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }
}
