using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Docx;
using static Writer.Tests.TestDocs;

namespace Writer.Tests;

/// <summary>Word's 插入 and 引用 on the engine: text boxes and shapes floating in a paragraph, bookmarks, numbered captions and drop caps.</summary>
public class DocxInsertTests
{
    static Dictionary<string, string> Props(params (string Name, string Value)[] pairs) => pairs.ToDictionary(p => p.Name, p => p.Value);

    static Document Reopen(Document doc)
    {
        var ms = new MemoryStream();
        doc.Save(ms);
        return OpenDocx(ms.ToArray());
    }

    static void AssertValid(Document doc) =>
        Assert.Empty(new OpenXmlValidator().Validate(((DocxDocument)doc).Package).Where(e => e.Part is MainDocumentPart).Select(e => e.Path?.XPath + ": " + e.Description));

    [Fact]
    public void Text_boxes_and_shapes_float_in_a_paragraph_and_survive_reopen()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "正文在旁边")), null);
        var box = Mutations.Add(p, "shape", Props(("geometry", "textbox"), ("text", "第一行\n第二行"), ("width", "5cm"), ("height", "2.5cm"), ("x", "1cm"), ("y", "0.5cm")), null);
        Mutations.Add(p, "shape", Props(("geometry", "ellipse"), ("fill", "ED7D31"), ("line", "none"), ("wrap", "front")), null);
        Assert.Equal("正文在旁边", p.GetProps()["text"]);
        AssertValid(doc);
        Assert.Contains("wps:txbx", box.GetRaw());

        using var reopened = Reopen(doc);
        var shapes = PathResolver.Query(reopened.Root, "//shape").Select(n => Registry.ToDisplay("docx", "shape", n.GetProps())).ToList();
        Assert.Equal(("rect", "第一行\n第二行", "FFFFFF", "000000", "5cm", "2.5cm", "1cm", "0.5cm", "square"),
            (shapes[0]["geometry"], shapes[0]["text"], shapes[0]["fill"], shapes[0]["line"], shapes[0]["width"], shapes[0]["height"], shapes[0]["x"], shapes[0]["y"], shapes[0]["wrap"]));
        Assert.Equal(("ellipse", "ED7D31", "none", "front", ""), (shapes[1]["geometry"], shapes[1]["fill"], shapes[1]["line"], shapes[1]["wrap"], shapes[1]["text"]));
        var ellipse = PathResolver.Single(reopened.Root, $"//shape[@id={shapes[1]["id"]}]");
        ellipse = Mutations.Set(ellipse, Props(("text", "圆"), ("fill", "none"), ("geometry", "star5")));
        Assert.Equal(("圆", "none", "star5"), (ellipse.GetProps()["text"], ellipse.GetProps()["fill"], ellipse.GetProps()["geometry"]));
        AssertValid(reopened);
        ellipse.Remove();
        Assert.Single(PathResolver.Query(reopened.Root, "//shape"));
        Assert.Equal("正文在旁边", PathResolver.Single(reopened.Root, "/body/paragraph[1]").GetProps()["text"]);
    }

    [Fact]
    public void Captions_number_by_label_and_keep_their_field_when_the_text_changes()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var first = Mutations.Add(body, "paragraph", Props(("caption", "图"), ("html", "图 1 流程")), null);
        var second = Mutations.Add(body, "paragraph", Props(("caption", "图"), ("html", "图 2 结构")), null);
        var table = Mutations.Add(body, "paragraph", Props(("caption", "表"), ("html", "表 1 数据")), null);
        Assert.Contains("SEQ 图", second.GetRaw());
        Assert.Contains(">2<", second.GetRaw()); // the second figure counts 2
        AssertValid(doc);

        using var reopened = Reopen(doc);
        var props = PathResolver.Query(reopened.Root, "/body/paragraph").Select(n => n.GetProps()).ToList();
        Assert.Equal(("图", "图 1 流程", "Caption"), (props[0]["caption"], props[0]["text"], props[0]["style"]));
        Assert.Equal(("图", "图 2 结构"), (props[1]["caption"], props[1]["text"]));
        Assert.Equal(("表", "表 1 数据"), (props[2]["caption"], props[2]["text"]));
        var p = Mutations.Set(PathResolver.Single(reopened.Root, "/body/paragraph[2]"), Props(("html", "图 3 新结构")));
        Assert.Equal(("图", "图 3 新结构"), (p.GetProps()["caption"], p.GetProps()["text"]));
        Assert.Contains("SEQ 图", p.GetRaw());
        p = Mutations.Set(p, Props(("caption", "none")));
        Assert.False(p.GetProps().ContainsKey("caption"));
        Assert.Equal("图 3 新结构", p.GetProps()["text"]);
    }

    [Fact]
    public void Bookmarks_wrap_a_paragraph_and_links_go_to_them()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var target = Mutations.Add(body, "paragraph", Props(("text", "结果"), ("bookmark", "Results")), null);
        Mutations.Add(body, "paragraph", Props(("html", "见 <a href=\"#Results\">结果</a>")), null);
        Assert.Throws<WriterException>(() => Mutations.Set(target, Props(("bookmark", "has space"))));
        AssertValid(doc);
        using var reopened = Reopen(doc);
        Assert.Equal("Results", PathResolver.Single(reopened.Root, "/body/paragraph[1]").GetProps()["bookmark"]);
        Assert.Contains("href=\"#Results\"", PathResolver.Single(reopened.Root, "/body/paragraph[2]").GetProps()["html"]);
        var moved = Mutations.Set(PathResolver.Single(reopened.Root, "/body/paragraph[2]"), Props(("bookmark", "Results"))); // a name moves to its new paragraph
        Assert.Equal("Results", moved.GetProps()["bookmark"]);
        Assert.False(PathResolver.Single(reopened.Root, "/body/paragraph[1]").GetProps().ContainsKey("bookmark"));
        moved = Mutations.Set(moved, Props(("bookmark", "none")));
        Assert.False(moved.GetProps().ContainsKey("bookmark"));
    }

    [Fact]
    public void A_drop_cap_is_a_frame_three_lines_high()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var cap = Mutations.Add(body, "paragraph", Props(("text", "春"), ("dropCap", "drop")), null);
        Mutations.Add(body, "paragraph", Props(("text", "眠不觉晓，处处闻啼鸟。")), null);
        Assert.Contains("w:dropCap=\"drop\"", cap.GetRaw());
        Assert.Contains("w:lines=\"3\"", cap.GetRaw());
        AssertValid(doc);
        using var reopened = Reopen(doc);
        var p = PathResolver.Single(reopened.Root, "/body/paragraph[1]");
        Assert.Equal(("drop", "春"), (p.GetProps()["dropCap"], p.GetProps()["text"]));
        p = Mutations.Set(p, Props(("dropCap", "none")));
        Assert.False(p.GetProps().ContainsKey("dropCap"));
    }

    [Fact]
    public void Comment_replies_share_their_comment_place_and_go_with_it()
    {
        using var doc = new DocxAdapter().Create();
        var body = doc.Root.Children.Single();
        var p = Mutations.Add(body, "paragraph", Props(("text", "增长了百分之二十")), null);
        var first = Mutations.Add(p, "comment", Props(("text", "数据来源？"), ("quote", "百分之二十")), null);
        var id = first.GetProps()["id"];
        Mutations.Add(p, "comment", Props(("text", "见附表"), ("parent", id)), null);
        AssertValid(doc);
        using var reopened = Reopen(doc);
        var comments = PathResolver.Query(reopened.Root, "//comment").Select(n => n.GetProps()).ToList();
        Assert.Equal(2, comments.Count);
        Assert.Equal(("见附表", id, "百分之二十"), (comments[1]["text"], comments[1]["parent"], comments[1]["quote"]));
        Assert.False(comments[0].ContainsKey("parent"));
        var parent = Mutations.Set(PathResolver.Single(reopened.Root, $"//comment[@id={id}]"), Props(("resolved", "true")));
        Assert.Equal("true", parent.GetProps()["resolved"]);
        parent = Mutations.Set(parent, Props(("resolved", "false")));
        Assert.False(parent.GetProps().ContainsKey("resolved"));
        Assert.Equal(id, PathResolver.Query(reopened.Root, "//comment").Last().GetProps()["parent"]); // un-resolving keeps the thread
        parent.Remove();
        Assert.Empty(PathResolver.Query(reopened.Root, "//comment"));
    }
}
