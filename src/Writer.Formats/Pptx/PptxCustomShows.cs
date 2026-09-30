using System.Text.Json.Nodes;
using Writer.Core;
using P = DocumentFormat.OpenXml.Presentation;
namespace Writer.Formats.Pptx;
static class PptxCustomShows
{
 internal static string Read(PptxDocument doc) {
  var byRel=doc.Slides.ToDictionary(s=>doc.Presentation.GetIdOfPart(s),s=>doc.SlideIdOf(s).Id!.Value.ToString());var shows=new JsonArray();
  foreach(var show in doc.Presentation.Presentation!.CustomShowList?.Elements<P.CustomShow>()??[]){var ids=new JsonArray();foreach(var s in show.SlideList?.Elements<P.SlideListEntry>()??[])if(s.Id?.Value is {} id&&byRel.TryGetValue(id,out var slide))ids.Add((JsonNode?)JsonValue.Create(slide));shows.Add((JsonNode)new JsonObject{["name"]=show.Name?.Value??"",["slides"]=ids});}return shows.ToJsonString();
 }
 internal static void Write(PptxDocument doc,string json) {
  var data=JsonNode.Parse(json) as JsonArray??throw Bad("Custom shows must be an array");var list=new P.CustomShowList();var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);var slides=doc.Slides.ToDictionary(s=>doc.SlideIdOf(s).Id!.Value.ToString(),s=>doc.Presentation.GetIdOfPart(s));
  var old=doc.Presentation.Presentation!.CustomShowList?.Elements<P.CustomShow>().ToDictionary(s=>s.Name!.Value!,s=>s.Id!.Value,StringComparer.OrdinalIgnoreCase)??[];uint id=old.Values.DefaultIfEmpty(0u).Max()+1;
  foreach(var item in data){var name=item?["name"]?.GetValue<string>()??"";if(string.IsNullOrWhiteSpace(name)||name.Length>255||!names.Add(name))throw Bad("Custom show names must be unique and nonempty");var refs=new P.SlideList();foreach(var value in item?["slides"] as JsonArray??throw Bad("Missing slide IDs")){if(!slides.TryGetValue(value!.GetValue<string>(),out var rel))throw Bad("Custom show refers to a missing slide");refs.Append(new P.SlideListEntry{Id=rel});}if(!refs.HasChildren)throw Bad("Custom shows need at least one slide");list.Append(new P.CustomShow(refs){Name=name,Id=old.GetValueOrDefault(name,id++)});}
  doc.Presentation.Presentation.CustomShowList=list.HasChildren?list:null;
 }
 internal static void Prune(PptxDocument doc){var list=doc.Presentation.Presentation!.CustomShowList;if(list is null)return;var valid=doc.Slides.Select(s=>doc.Presentation.GetIdOfPart(s)).ToHashSet();foreach(var item in list.Descendants<P.SlideListEntry>().Where(s=>!valid.Contains(s.Id?.Value??"")).ToArray())item.Remove();foreach(var show in list.Elements<P.CustomShow>().Where(s=>s.SlideList is not {HasChildren:true}).ToArray())show.Remove();if(!list.HasChildren)list.Remove();}
 static WriterException Bad(string message)=>new(ErrorCode.Validation,message,"Use [{name,slides:[slideId,...]}]. Repeated slide IDs preserve their order.");
}
