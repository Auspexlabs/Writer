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
