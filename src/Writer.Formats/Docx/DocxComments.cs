using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace Writer.Formats.Docx;

/// <summary>A comment, projected under the paragraph holding its reference mark. Its text lives in word/comments.xml,
/// its anchor is the commentRangeStart / commentRangeEnd pair around runs of the paragraph, and "done" in commentsExtended.xml.</summary>
sealed class DocxComment(DocxDocument doc, W.Comment comment) : Node
{
    public override string Kind => "comment";
    public override object Anchor => comment;

    public override IReadOnlyDictionary<string, string> GetProps()
    {
        var props = new Dictionary<string, string> { ["id"] = comment.Id?.Value ?? "" };
        if (comment.Author?.Value is { } author) props["author"] = author;
        if (comment.Initials?.Value is { Length: > 0 } initials) props["initials"] = initials;
        if (comment.Date?.InnerText is { } date) props["date"] = date;
        props["text"] = DocxComments.Text(comment);
        props["quote"] = DocxComments.Quote(doc, comment.Id?.Value ?? "");
        if (DocxComments.Resolved(doc, comment)) props["resolved"] = "true";
        if (DocxComments.ParentOf(doc, comment)?.Id?.Value is { } parent) props["parent"] = parent;
        return props;
    }

    public override string GetRaw() => comment.OuterXml;

    public override void SetProp(string name, string value)
    {
        switch (name)
        {
            case "author": comment.Author = value; break;
            case "initials": comment.Initials = value.Length > 0 ? new StringValue(value) : null; break;
            case "date": comment.Date = new DateTimeValue { InnerText = value }; break;
            case "text": DocxComments.SetText(comment, value); break;
            case "quote":
                var paragraph = DocxComments.ReferenceRun(doc, comment.Id!.Value!)?.Ancestors<W.Paragraph>().FirstOrDefault()
                    ?? throw new WriterException(ErrorCode.Validation, "The comment has no anchor paragraph", "Remove the comment and add it again.");
                DocxComments.Unanchor(doc, comment.Id.Value!);
                DocxComments.Anchor(paragraph, comment.Id.Value!, value);
                break;
            case "resolved": DocxComments.SetResolved(doc, comment, value == "true"); break;
            case "parent": DocxComments.SetParent(doc, comment, value); break;
        }
    }

    public override void Remove() => DocxComments.Remove(doc, comment);
}

static class DocxComments
{
    static IEnumerable<W.Comment> All(DocxDocument doc) => doc.Main.WordprocessingCommentsPart?.Comments?.Elements<W.Comment>() ?? [];

    public static int Count(DocxDocument doc) => All(doc).Count();

    /// <summary>The comments whose reference mark sits in the paragraph, in reading order.</summary>
    public static IEnumerable<Node> In(DocxDocument doc, W.Paragraph p)
    {
        var byId = All(doc).Where(c => c.Id?.Value is not null).ToDictionary(c => c.Id!.Value!, c => c);
        foreach (var reference in p.Descendants<W.CommentReference>())
            if (reference.Id?.Value is { } id && byId.TryGetValue(id, out var comment)) yield return new DocxComment(doc, comment);
    }

    public static Node Add(DocxDocument doc, W.Paragraph p, IReadOnlyDictionary<string, string> props)
    {
        var id = (1 + All(doc).Select(c => int.TryParse(c.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : -1).DefaultIfEmpty(-1).Max())
            .ToString(CultureInfo.InvariantCulture);
        var parent = props.GetValueOrDefault("parent") is { Length: > 0 } parentId
            ? All(doc).FirstOrDefault(c => c.Id?.Value == parentId) ?? throw new WriterException(ErrorCode.Validation, $"No comment {parentId} to reply to", "Give the id of a comment in the document.")
            : null;
        if (parent is not null) AnchorBeside(doc, parent.Id!.Value!, id); // a reply shares its comment's place, as Word writes it
        else Anchor(p, id, props.GetValueOrDefault("quote")); // validates the quote before anything is added
        var part = doc.Main.WordprocessingCommentsPart ?? doc.Main.AddNewPart<WordprocessingCommentsPart>();
        var comments = part.Comments ??= new W.Comments();
        var comment = new W.Comment
        {
            Id = id,
            Author = props.GetValueOrDefault("author") ?? DocxRevisions.Author(doc),
            Initials = props.GetValueOrDefault("initials") is { Length: > 0 } initials ? new StringValue(initials) : null,
            Date = new DateTimeValue { InnerText = props.GetValueOrDefault("date") ?? DocxRevisions.Now() },
        };
        SetText(comment, props.GetValueOrDefault("text") ?? "");
        comments.Append(comment);
        var node = new DocxComment(doc, comment);
        if (parent is not null) Extended(doc, comment, create: true)!.ParaIdParent = Extended(doc, parent, create: true)!.ParaId;
        if (props.GetValueOrDefault("resolved") == "true") node.SetProp("resolved", "true");
        return node;
    }

    public static string Text(W.Comment comment) => string.Join("\n", comment.Elements<W.Paragraph>().Select(DocxRuns.ParagraphText));

    public static void SetText(W.Comment comment, string text)
    {
        var paraIds = comment.Elements<W.Paragraph>().Select(p => p.ParagraphId?.Value).ToList();
        comment.RemoveAllChildren<W.Paragraph>();
        var lines = text.ReplaceLineEndings("\n").Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            var p = new W.Paragraph();
            if (i == 0) p.Append(new W.Run(new W.AnnotationReferenceMark()));
            if (lines[i].Length > 0) p.Append(new W.Run(DocxRuns.TextElements(lines[i])));
            if (i == lines.Length - 1 && paraIds.Count > 0 && paraIds[^1] is { } keep) p.ParagraphId = keep; // keeps the commentsExtended link
            comment.Append(p);
        }
    }

    /// <summary>The visible text between the comment's range markers; empty when it has none.</summary>
    public static string Quote(DocxDocument doc, string id)
    {
        var body = doc.Main.Document!.Body!;
        var start = body.Descendants<W.CommentRangeStart>().FirstOrDefault(s => s.Id?.Value == id);
        if (start is null) return "";
        var inside = false;
        var text = new System.Text.StringBuilder();
        foreach (var e in body.Descendants())
        {
            if (ReferenceEquals(e, start)) { inside = true; continue; }
            if (!inside) continue;
            if (e is W.CommentRangeEnd end && end.Id?.Value == id) break;
            if (e is W.Run run && DocxRuns.HasText(run) && DocxRuns.Revision(run) is not W.DeletedRun) text.Append(DocxRuns.RunText(run));
        }
        return text.ToString();
    }

    /// <summary>A reply's range and reference mark, right after its comment's.</summary>
    static void AnchorBeside(DocxDocument doc, string parent, string id)
    {
        var body = doc.Main.Document!.Body!;
        var start = body.Descendants<W.CommentRangeStart>().FirstOrDefault(x => x.Id?.Value == parent);
        var end = body.Descendants<W.CommentRangeEnd>().FirstOrDefault(x => x.Id?.Value == parent);
        var reference = ReferenceRun(doc, parent) ?? throw new WriterException(ErrorCode.Validation, $"Comment {parent} has no place in the text", "Reply to a comment that is anchored.");
        if (start is not null && end is not null) { start.InsertAfterSelf(new W.CommentRangeStart { Id = id }); end.InsertAfterSelf(new W.CommentRangeEnd { Id = id }); }
        reference.InsertAfterSelf(new W.Run(new W.CommentReference { Id = id }));
    }

    /// <summary>Makes the comment a reply to another (by id), or a comment of its own again (none).</summary>
    public static void SetParent(DocxDocument doc, W.Comment comment, string parentId)
    {
        if (parentId is "none" or "") { if (Extended(doc, comment) is { } ex) ex.ParaIdParent = null; return; }
        var parent = All(doc).FirstOrDefault(c => c.Id?.Value == parentId) ?? throw new WriterException(ErrorCode.Validation, $"No comment {parentId} to reply to", "Give the id of a comment in the document.");
        Extended(doc, comment, create: true)!.ParaIdParent = Extended(doc, parent, create: true)!.ParaId;
    }

    /// <summary>The comment a reply answers (commentsExtended's paraIdParent).</summary>
    public static W.Comment? ParentOf(DocxDocument doc, W.Comment comment) =>
        Extended(doc, comment)?.ParaIdParent?.Value is { } parentPara ? All(doc).FirstOrDefault(c => c.Elements<W.Paragraph>().LastOrDefault()?.ParagraphId?.Value == parentPara) : null;

    public static W.Run? ReferenceRun(DocxDocument doc, string id) =>
        doc.Main.Document!.Body!.Descendants<W.CommentReference>().FirstOrDefault(r => r.Id?.Value == id)?.Parent as W.Run;

    /// <summary>Puts the range markers around the quoted text of the paragraph (the whole paragraph without a quote) and the
    /// reference mark after them, splitting runs at the quote's ends.</summary>
    public static void Anchor(W.Paragraph p, string id, string? quote)
    {
        var reference = new W.Run(new W.CommentReference { Id = id });
        if (quote is null or "")
        {
            var first = p.ChildElements.FirstOrDefault(c => c is not W.ParagraphProperties);
            if (first is null) p.Append(new W.CommentRangeStart { Id = id });
            else p.InsertBefore(new W.CommentRangeStart { Id = id }, first);
            p.Append(new W.CommentRangeEnd { Id = id });
            p.Append(reference);
            return;
        }
        var text = DocxRuns.ParagraphText(p);
        var at = text.IndexOf(quote, StringComparison.Ordinal);
        if (at < 0)
            throw new WriterException(ErrorCode.Validation, $"The paragraph does not contain '{NodeJson.Preview(quote, 40)}'",
                $"Quote text from the paragraph: \"{NodeJson.Preview(text, 60)}\", or leave quote out to comment on the whole paragraph.");
        var firstRun = SplitAt(p, at)!;
        var lastRun = SplitAt(p, at + quote.Length, endOf: true)!;
        firstRun.InsertBeforeSelf(new W.CommentRangeStart { Id = id });
        var end = lastRun.InsertAfterSelf(new W.CommentRangeEnd { Id = id });
        end.InsertAfterSelf(reference);
    }

    /// <summary>Makes a run boundary at a visible character offset and returns the run starting there (or ending there with endOf).</summary>
    static W.Run? SplitAt(W.Paragraph p, int offset, bool endOf = false)
    {
        var pos = 0;
        foreach (var (run, _) in DocxRuns.Walk(p).ToList())
        {
            var text = DocxRuns.RunText(run);
            var end = pos + text.Length;
            if (offset == pos && !endOf) return run;
            if (offset == end && endOf) return run;
            if (offset > pos && offset < end)
            {
                DocxReplace.SplitRun(run, offset - pos);
                return endOf ? run : (W.Run)run.NextSibling()!;
            }
            pos = end;
        }
        return null;
    }

    /// <summary>Removes the range markers and reference mark, joining the runs they split so the paragraph is as it was.</summary>
    public static void Unanchor(DocxDocument doc, string id)
    {
        var body = doc.Main.Document!.Body!;
        foreach (var r in body.Descendants<W.CommentReference>().Where(r => r.Id?.Value == id).ToList())
        {
            var run = r.Parent as W.Run;
            r.Remove();
            if (run is not null && run.ChildElements.All(c => c is W.RunProperties)) run.Remove();
        }
        foreach (var e in body.Descendants().Where(e => e is W.CommentRangeStart s && s.Id?.Value == id || e is W.CommentRangeEnd n && n.Id?.Value == id).ToList())
        {
            var (previous, next) = (e.PreviousSibling(), e.NextSibling());
            e.Remove();
            DocxRuns.Rejoin(previous, next);
        }
    }

    public static void Remove(DocxDocument doc, W.Comment comment)
    {
        foreach (var reply in All(doc).Where(c => !ReferenceEquals(c, comment) && ParentOf(doc, c) == comment).ToList()) Remove(doc, reply); // its replies go with it, as in Word
        var id = comment.Id?.Value ?? "";
        Unanchor(doc, id);
        SetResolved(doc, comment, false);
        var part = doc.Main.WordprocessingCommentsPart!;
        comment.Remove();
        if (!All(doc).Any()) doc.Main.DeletePart(part);
    }

    /// <summary>The comment's commentsExtended entry (done, paraIdParent), made when asked to (its last paragraph gets a paraId).</summary>
    static W15.CommentEx? Extended(DocxDocument doc, W.Comment comment, bool create = false)
    {
        var last = comment.Elements<W.Paragraph>().LastOrDefault();
        var paraId = last?.ParagraphId?.Value;
        var found = paraId is null ? null : doc.Main.WordprocessingCommentsExPart?.CommentsEx?.Elements<W15.CommentEx>().FirstOrDefault(c => c.ParaId?.Value == paraId);
        if (found is not null || !create) return found;
        if (last is null) comment.Append(last = new W.Paragraph());
        last.ParagraphId ??= new HexBinaryValue(Random.Shared.Next(1, 0x7FFFFFFF).ToString("X8", CultureInfo.InvariantCulture));
        var part = doc.Main.WordprocessingCommentsExPart ?? doc.Main.AddNewPart<WordprocessingCommentsExPart>();
        var root = part.CommentsEx ??= new W15.CommentsEx();
        return root.AppendChild(new W15.CommentEx { ParaId = last.ParagraphId.Value, Done = false });
    }

    public static bool Resolved(DocxDocument doc, W.Comment comment) => Extended(doc, comment)?.Done?.Value == true;

    public static void SetResolved(DocxDocument doc, W.Comment comment, bool done)
    {
        var existing = Extended(doc, comment);
        if (!done)
        {
            if (existing is null) return;
            if (existing.ParaIdParent is not null || All(doc).Any(c => ParentOf(doc, c) == comment)) { existing.Done = false; return; } // a thread keeps its links
            var container = existing.Parent!;
            existing.Remove();
            if (!container.HasChildren) doc.Main.DeletePart(doc.Main.WordprocessingCommentsExPart!);
            return;
        }
        if (existing is not null) { existing.Done = true; return; }
        var last = comment.Elements<W.Paragraph>().LastOrDefault();
        if (last is null) comment.Append(last = new W.Paragraph());
        last.ParagraphId ??= new HexBinaryValue(Random.Shared.Next(1, 0x7FFFFFFF).ToString("X8", CultureInfo.InvariantCulture));
        var part = doc.Main.WordprocessingCommentsExPart ?? doc.Main.AddNewPart<WordprocessingCommentsExPart>();
        var root = part.CommentsEx ??= new W15.CommentsEx();
        root.Append(new W15.CommentEx { ParaId = last.ParagraphId.Value, Done = true });
    }
}
