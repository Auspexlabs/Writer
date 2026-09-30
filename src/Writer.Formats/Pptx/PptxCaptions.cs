using System.Text;
using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P19 = DocumentFormat.OpenXml.Office2019.Presentation;

namespace Writer.Formats.Pptx;

static class PptxCaptions
{
    const string Extension="{3AFAAA56-56D3-431D-BCD4-E75A35582382}";
    const string Relationship="http://schemas.microsoft.com/office/2017/04/relationships/track";
    const string Ns="http://schemas.microsoft.com/office/powerpoint/2017/3/main";
    static P14.Media? Media(P.Picture picture)=>picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.Descendants<P14.Media>().FirstOrDefault();
    internal static string Read(OpenXmlPart owner,P.Picture picture) {
        var raw=Media(picture)?.Descendants().FirstOrDefault(e=>e.LocalName=="tracksInfo"&&e.NamespaceUri==Ns);var info=raw is null?null:new P19.TracksInfo(raw.OuterXml);var tracks=new JsonArray();
        foreach(var track in info?.TrackList?.Elements<P19.Track>()??[]) {
            var item=new JsonObject{["id"]=track.Id?.Value,["label"]=track.Label?.Value??"",["lang"]=track.Lang?.Value??""};
            if(track.Link?.Value is {} linked)item["link"]=owner.ExternalRelationships.FirstOrDefault(r=>r.Id==linked)?.Uri.OriginalString??"";
            else if(track.Embed?.Value is {} embedded&&owner.Parts.FirstOrDefault(p=>p.RelationshipId==embedded).OpenXmlPart is {} part){using var reader=new StreamReader(part.GetStream(),Encoding.UTF8);item["text"]=reader.ReadToEnd();}
            tracks.Add((JsonNode)item);
        }
        return new JsonObject{["display"]=info?.DisplayLoc?.InnerText??"media",["tracks"]=tracks}.ToJsonString();
    }
    internal static void Write(OpenXmlPart owner,P.Picture picture,string json) {
        if(Media(picture) is not {} media)throw Bad("Captions need an embedded media object");
        var data=JsonNode.Parse(json) as JsonObject??throw Bad("Expected caption settings");var tracks=data["tracks"] as JsonArray??throw Bad("Expected caption tracks");
        var display=data["display"]?.GetValue<string>()??"media";if(display is not ("media" or "slide")||tracks.Count>32)throw Bad("Invalid caption display or track count");
        var entries=tracks.Select(t=>(Id:t?["id"]?.GetValue<string>(),Label:t?["label"]?.GetValue<string>()??"",Lang:t?["lang"]?.GetValue<string>()??"",Text:t?["text"]?.GetValue<string>(),Link:t?["link"]?.GetValue<string>())).ToArray();
        if(entries.Any(e=>e.Label.Length==0||e.Text is null&&string.IsNullOrEmpty(e.Link)||e.Text is not null&&(e.Text.Length>2_000_000||!System.Text.RegularExpressions.Regex.IsMatch(e.Text,"^\\uFEFF?WEBVTT(?:[ \\t][^\\r\\n]*)?(?:\\r?\\n|$)"))||e.Link is {Length:>0}&&(!Uri.TryCreate(e.Link,UriKind.Absolute,out var uri)||uri.Scheme is not ("https" or "http"))))throw Bad("Captions require a label and UTF-8 WEBVTT text or an HTTP(S) link");
        var old=media.Descendants().Where(e=>e.LocalName=="tracksInfo"&&e.NamespaceUri==Ns).ToArray();var relations=old.SelectMany(i=>new P19.TracksInfo(i.OuterXml).Descendants<P19.Track>()).SelectMany(t=>new[]{t.Embed?.Value,t.Link?.Value}).OfType<string>().ToArray();
        var list=new P19.TrackList();
        foreach(var e in entries){var id=Guid.TryParse(e.Id,out var guid)?guid.ToString("B").ToUpperInvariant():Guid.NewGuid().ToString("B").ToUpperInvariant();var track=new P19.Track{Id=id,Label=e.Label};if(e.Lang.Length>0)track.Lang=e.Lang;
            if(e.Link is {Length:>0})track.Link=owner.AddExternalRelationship(Relationship,new Uri(e.Link)).Id;
            else {var part=owner.AddExtendedPart(Relationship,"text/vtt","vtt");using var bytes=new MemoryStream(Encoding.UTF8.GetBytes(e.Text!));part.FeedData(bytes);track.Embed=owner.GetIdOfPart(part);}
            list.Append(track);
        }
        foreach(var info in old){var parent=info.Parent;info.Remove();if(parent?.LocalName=="ext"&&!parent.HasChildren)parent.Remove();}
        if(list.HasChildren){var ext=media.GetFirstChild<P14.ExtensionList>();if(ext is null){ext=new P14.ExtensionList();media.Append(ext);}ext.Append(new P.Extension(new P19.TracksInfo(list){DisplayLoc=display=="slide"?P19.DisplayLocation.Slide:P19.DisplayLocation.Media}){Uri=Extension});}
        foreach(var id in relations.Distinct()){if(owner.RootElement?.Descendants().Any(n=>n.GetAttributes().Any(a=>a.NamespaceUri=="http://schemas.openxmlformats.org/officeDocument/2006/relationships"&&a.Value==id))==true)continue;if(owner.Parts.Any(p=>p.RelationshipId==id))owner.DeletePart(id);else if(owner.ExternalRelationships.Any(r=>r.Id==id))owner.DeleteExternalRelationship(id);}
    }
    static WriterException Bad(string message)=>new(ErrorCode.Validation,message,"Use {display:media|slide,tracks:[{id?,label,lang,text}]}.");
}
