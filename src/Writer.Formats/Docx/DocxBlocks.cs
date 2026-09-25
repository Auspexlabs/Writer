using System.Globalization;
using DocumentFormat.OpenXml;
using Writer.Core;
using Writer.Formats.Common;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

/// <summary>A node whose native element holds children: body, cell, table, row, paragraph.</summary>
interface IDocxContainer
{
    OpenXmlElement Container { get; }
}

/// <summary>Block-level projection and editing shared by the body and table cells. Block content controls are transparent.</summary>
static class DocxBlocks
{
    public static IEnumerable<Node> Project(DocxDocument doc, OpenXmlElement container)
    {
        OpenXmlElement? skipTo = null; // the last paragraph of a bare TOC field already projected as one block
        foreach (var child in container.ChildElements)
        {
            if (skipTo is not null)
            {
                if (ReferenceEquals(child, skipTo)) skipTo = null;
                continue;
            }
            switch (child)
            {
                case W.Paragraph p when DocxToc.FieldAt(doc, p) is { } field:
                    var group = DocxToc.Group(doc, field);
                    if (group.Count > 1) skipTo = group[^1];
                    yield return new DocxToc(doc, group);
                    break;
                case W.Paragraph or W.Table:
                    yield return Wrap(doc, child);
                    break;
                case W.SdtBlock sdt when DocxToc.IsToc(sdt):
                    yield return new DocxToc(doc, [sdt]);
                    break;
                case W.SdtBlock sdt when sdt.SdtContentBlock is { } content:
                    foreach (var n in Project(doc, content)) yield return n;
                    break;
            }
        }
    }

    public static Node Wrap(DocxDocument doc, OpenXmlElement element) => element switch
    {
        W.Paragraph p when DocxToc.HasTocField(p) => new DocxToc(doc, DocxToc.BareField(p)),
        W.Paragraph p when DocxImage.IsPictureParagraph(p) => DocxImage.Block(doc, p),
        W.Paragraph p when DocxPageBreak.IsPageBreak(p) => new DocxPageBreak(doc, p),
        W.Paragraph p => new DocxParagraph(doc, p),
        W.Table t => new DocxTable(doc, t),
        W.SdtBlock sdt when DocxToc.IsToc(sdt) => new DocxToc(doc, [sdt]),
        _ => throw new InvalidOperationException($"Not a block: {element.LocalName}"),
    };

    /// <summary>Creates a block, inserts it and applies the remaining properties.</summary>
    public static Node Add(DocxDocument doc, Node parent, OpenXmlElement container, string kind, IReadOnlyDictionary<string, string> props, int? index)
    {
        var (element, consumed) = kind switch
        {
            "paragraph" => ((OpenXmlElement)new W.Paragraph(), Array.Empty<string>()),
            "heading" => (Heading(doc, props), []),
            "code" => (new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = doc.Styles.ResolveStyle("Code", "paragraph") })), []),
            "table" => DocxTable.New(doc, props),
            "image" => DocxImage.New(doc, props),
            "pagebreak" => (new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })), []),
            "toc" => DocxToc.New(doc, props),
            _ => throw new WriterException(ErrorCode.UnsupportedKind, $"Cannot add {kind} under {parent.Kind}", "Blocks: paragraph, heading, code, table, image, pagebreak, toc."),
        };
        InsertAt(parent, container, element, index);
        if (kind is "paragraph" or "heading" or "code" && DocxRevisions.Tracking(doc)) DocxRevisions.MarkInserted(doc, (W.Paragraph)element);
        var node = Wrap(doc, element);
        foreach (var (name, value) in props)
            if (!consumed.Contains(name)) node.SetProp(name, value);
        return node;
    }

    static W.Paragraph Heading(DocxDocument doc, IReadOnlyDictionary<string, string> props)
    {
        var level = props.TryGetValue("level", out var l) ? int.Parse(l, CultureInfo.InvariantCulture) : 1;
        return new W.Paragraph(new W.ParagraphProperties(new W.ParagraphStyleId { Val = doc.Styles.HeadingStyleId(level) }));
    }

    /// <summary>Inserts before the child at the 1-based position among the parent's projected children, or appends
    /// (before the section properties when the container is the body).</summary>
    public static void InsertAt(Node parent, OpenXmlElement container, OpenXmlElement element, int? index)
    {
        if (OoxmlTree.InsertBefore(parent, container, element, index)) return;
        if (container is W.Body body && body.GetFirstChild<W.SectionProperties>() is { } section) body.InsertBefore(element, section);
        else container.Append(element);
    }

    /// <summary>Removes an element, keeping table cells valid (a cell always ends with a paragraph).</summary>
    public static void Detach(OpenXmlElement element)
    {
        var parent = element.Parent;
        element.Remove();
        if (parent is W.TableCell cell && cell.LastChild is not W.Paragraph) cell.Append(new W.Paragraph());
    }

    public static void Move(Node target, OpenXmlElement element, int? index)
    {
        var container = ContainerOf(target);
        if (container.Ancestors().Contains(element) || ReferenceEquals(container, element))
            throw new WriterException(ErrorCode.Validation, "Cannot move an element into itself", "Pick another target.");
        Detach(element);
        InsertAt(target, container, element, index);
    }

    public static OpenXmlElement ContainerOf(Node node) =>
        node is IDocxContainer c ? c.Container
        : throw new WriterException(ErrorCode.Validation, $"{node.Kind} cannot hold children", "Targets: body, cell, table, row, paragraph.");

    public static void ReplaceRaw(DocxDocument doc, OpenXmlElement element, string xml) => RawXml.Replace(element, xml, doc.Namespaces);
}
