using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace Writer.Formats.Pptx;

static class PptxMedia
{
    public static DataPartReferenceRelationship? Reference(OpenXmlPart owner, P.Picture picture)
    {
        var nv = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
        var id = nv?.GetFirstChild<A.VideoFromFile>()?.Link?.Value ?? nv?.GetFirstChild<A.AudioFromFile>()?.Link?.Value;
        return id is null ? null : owner.DataPartReferenceRelationships.FirstOrDefault(r => r.Id == id);
    }
    public static (string ContentType, byte[] Data)? Bytes(OpenXmlPart owner, P.Picture picture)
    {
        if (Reference(owner, picture)?.DataPart is not { } data) return null;
        using var stream = data.GetStream(FileMode.Open, FileAccess.Read); using var copy = new MemoryStream(); stream.CopyTo(copy);
        return (data.ContentType, copy.ToArray());
    }
    static string ShapeId(P.Picture picture) => picture.NonVisualPictureProperties!.NonVisualDrawingProperties!.Id!.Value.ToString(CultureInfo.InvariantCulture);
    static P14.Media? Metadata(P.Picture picture) => picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.Descendants<P14.Media>().FirstOrDefault();
    static P.CommonMediaNode? Timing(OpenXmlPart owner, P.Picture picture) => (owner as SlidePart)?.Slide?.Timing?.Descendants<P.CommonMediaNode>().FirstOrDefault(n => n.TargetElement?.GetFirstChild<P.ShapeTarget>()?.ShapeId?.Value == ShapeId(picture));
    static double Number(OpenXmlElement? e, string name, double fallback = 0) => double.TryParse(e?.GetAttributes().FirstOrDefault(a=>a.LocalName==name&&a.NamespaceUri=="").Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) && double.IsFinite(n) ? n : fallback;
    internal static string ReadPlayback(OpenXmlPart owner, P.Picture picture)
    {
        var media = Metadata(picture); var node = Timing(owner, picture); var time = node?.CommonTimeNode;
        return NodeJson.Compact(w => {
            w.WriteStartObject();
            w.WriteNumber("trimStart", Number(media?.MediaTrim, "st") / 1000); w.WriteNumber("trimEnd", Number(media?.MediaTrim, "end") / 1000);
            w.WriteNumber("fadeIn", Number(media?.MediaFade, "in") / 1000); w.WriteNumber("fadeOut", Number(media?.MediaFade, "out") / 1000);
            w.WriteNumber("volume", Number(node, "vol", 100000) / 1000); w.WriteBoolean("muted", node?.Mute?.Value ?? false);
            w.WriteBoolean("loop", time?.RepeatCount?.Value == "indefinite");
            w.WriteBoolean("autoplay", time?.StartConditionList?.Elements<P.Condition>().Any(c => c.Delay?.Value == "0" && (c.Event is null || c.Event.InnerText == "onBegin")) ?? false);
            w.WriteNumber("slideCount", node?.SlideCount?.Value ?? 1); w.WriteBoolean("hideStopped", node?.ShowWhenStopped?.Value == false);
            w.WriteStartArray("bookmarks"); foreach (var b in media?.MediaBookmarkList?.Elements<P14.MediaBookmark>() ?? []) { w.WriteStartObject(); w.WriteString("name", b.Name?.Value ?? ""); w.WriteNumber("time", Number(b, "time") / 1000); w.WriteEndObject(); } w.WriteEndArray();
            w.WriteEndObject();
        });
    }
    internal static void WritePlayback(SlidePart slide, P.Picture picture, string json)
    {
        using var parsed = JsonDocument.Parse(json); var p = parsed.RootElement;
        if (p.ValueKind != JsonValueKind.Object || Metadata(picture) is not { } media) throw Bad("Playback settings need an embedded media object");
        double N(string k, double fallback = 0) { var v = p.TryGetProperty(k, out var el) && el.TryGetDouble(out var n) ? n : fallback; if (!double.IsFinite(v) || v < 0 || v > 8640000) throw Bad("Invalid media time or volume"); return v; }
        bool B(string k) => p.TryGetProperty(k, out var v) && v.ValueKind == JsonValueKind.True;
        var start = N("trimStart"); var end = N("trimEnd"); var fadeIn = N("fadeIn"); var fadeOut = N("fadeOut"); var volume = N("volume", 100); var count = N("slideCount", 1);
        if (volume > 100 || count < 1 || count > 999 || count % 1 != 0) throw Bad("Volume must be 0–100 and slideCount 1–999");
        var bookmarks = p.TryGetProperty("bookmarks", out var bs) ? bs.EnumerateArray().Select(b => (Name: b.GetProperty("name").GetString() ?? "", Time: b.GetProperty("time").GetDouble())).ToArray() : [];
        if (bookmarks.Length > 1000 || bookmarks.Select(b => b.Name).Distinct().Count() != bookmarks.Length || bookmarks.Any(b => b.Name.Length == 0 || !double.IsFinite(b.Time) || b.Time < 0 || b.Time > 8640000)) throw Bad("Media bookmarks need unique names and nonnegative times");
        void Set(OpenXmlElement e, string key, double v) => e.SetAttribute(new OpenXmlAttribute("", key, "", Math.Round(v * 1000).ToString(CultureInfo.InvariantCulture)));
        var trim = media.MediaTrim ??= new P14.MediaTrim(); Set(trim, "st", start); Set(trim, "end", end);
        var fade = media.MediaFade ??= new P14.MediaFade(); Set(fade, "in", fadeIn); Set(fade, "out", fadeOut);
        media.MediaBookmarkList = bookmarks.Length == 0 ? null : new P14.MediaBookmarkList(bookmarks.Select(b => { var mark = new P14.MediaBookmark { Name = b.Name }; Set(mark, "time", b.Time); return mark; }));
        var node = Timing(slide, picture);
        if (node is null) {
            var timing = slide.Slide!.Timing;
            if (timing is null) { timing = new P.Timing(); slide.Slide.AddChild(timing); }
            timing.TimeNodeList ??= new P.TimeNodeList();
            var root = timing.TimeNodeList.GetFirstChild<P.ParallelTimeNode>()?.CommonTimeNode;
            var next = timing.Descendants<P.CommonTimeNode>().Select(c => c.Id?.Value ?? 0).DefaultIfEmpty(0u).Max() + 1;
            if (root is null) { root = new P.CommonTimeNode { Id = next++, Duration = "indefinite", Restart = P.TimeNodeRestartValues.Never, NodeType = P.TimeNodeValues.TmingRoot }; timing.TimeNodeList.Append(new P.ParallelTimeNode(root)); }
            var children = root.ChildTimeNodeList ??= new P.ChildTimeNodeList();
            node = new P.CommonMediaNode(new P.CommonTimeNode { Id = next, Duration = "media", Fill = P.TimeNodeFillValues.Hold }, new P.TargetElement(new P.ShapeTarget { ShapeId = ShapeId(picture) }));
            children.Append(Reference(slide, picture)?.DataPart.ContentType.StartsWith("video/", StringComparison.Ordinal) == true ? (OpenXmlElement)new P.Video(node) : new P.Audio(node));
        }
        node.Volume = (int)Math.Round(volume * 1000); node.Mute = B("muted"); node.SlideCount = (uint)count; node.ShowWhenStopped = !B("hideStopped");
        var ctn = node.CommonTimeNode ??= new P.CommonTimeNode(); ctn.RepeatCount = B("loop") ? "indefinite" : null;
        ctn.StartConditionList = B("autoplay") ? new P.StartConditionList(new P.Condition { Delay = "0" }) : new P.StartConditionList(new P.Condition(new P.TargetElement(new P.ShapeTarget { ShapeId = ShapeId(picture) })) { Event = P.TriggerEventValues.OnClick, Delay = "0" });
    }
    internal static void CopyPlayback(OpenXmlElement source, OpenXmlElement target, SlidePart from, SlidePart to)
    {
        var a = source.Descendants<P.Picture>().Prepend(source as P.Picture).Where(p => p is not null).ToArray(); var b = target.Descendants<P.Picture>().Prepend(target as P.Picture).Where(p => p is not null).ToArray();
        for (var i = 0; i < Math.Min(a.Length, b.Length); i++) if (Metadata(b[i]!) is not null && Timing(from, a[i]!) is not null) WritePlayback(to, b[i]!, ReadPlayback(from, a[i]!));
    }
    static WriterException Bad(string message) => new(ErrorCode.Validation, message, "Use seconds for trimStart, trimEnd, fadeIn, fadeOut and bookmarks; trimEnd is time removed from the end.");
    public static void Attach(PptxDocument doc, SlidePart slide, P.Picture picture, string uri)
    {
        var comma = uri.IndexOf(',');
        if (comma < 0 || !uri.StartsWith("data:", StringComparison.Ordinal) || !uri[..comma].EndsWith(";base64", StringComparison.Ordinal)) throw new WriterException(ErrorCode.Validation, "Media must be a base64 data URI", "Choose an audio or video file.");
        var mime = uri[5..(comma - 7)]; var video = mime.StartsWith("video/", StringComparison.Ordinal);
        if (!video && !mime.StartsWith("audio/", StringComparison.Ordinal)) throw new WriterException(ErrorCode.Validation, "Unsupported media type", "Choose an audio or video file.");
        var bytes = Convert.FromBase64String(uri[(comma + 1)..]);
        var extension = mime switch { "video/mp4" => ".mp4", "video/webm" => ".webm", "video/quicktime" => ".mov", "audio/mpeg" => ".mp3", "audio/mp4" or "audio/x-m4a" => ".m4a", "audio/wav" or "audio/x-wav" => ".wav", "audio/ogg" => ".ogg", _ => ".bin" };
        var part = doc.Package.CreateMediaDataPart(mime, extension); using (var stream = new MemoryStream(bytes)) part.FeedData(stream);
        var legacyId = video ? slide.AddVideoReferenceRelationship(part).Id : slide.AddAudioReferenceRelationship(part).Id;
        var id = slide.AddMediaReferenceRelationship(part).Id;
        var nv = picture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties ??= new P.ApplicationNonVisualDrawingProperties();
        nv.RemoveAllChildren<A.VideoFromFile>(); nv.RemoveAllChildren<A.AudioFromFile>();
        nv.AddChild(video ? new A.VideoFromFile { Link = legacyId } : new A.AudioFromFile { Link = legacyId });
        var list = nv.GetFirstChild<P.ApplicationNonVisualDrawingPropertiesExtensionList>() ?? nv.AppendChild(new P.ApplicationNonVisualDrawingPropertiesExtensionList());
        foreach (var old in list.Elements<P.ApplicationNonVisualDrawingPropertiesExtension>().Where(e => e.GetFirstChild<P14.Media>() is not null).ToList()) old.Remove();
        list.Append(new P.ApplicationNonVisualDrawingPropertiesExtension(new P14.Media { Embed = id }) { Uri = "{DAA4B4D4-3F8A-4F9B-91E9-EDA526CFB0B4}" });
        var props = picture.NonVisualPictureProperties.NonVisualDrawingProperties!;
        props.HyperlinkOnClick = new A.HyperlinkOnClick { Id = "", Action = "ppaction://media" };
    }
}
