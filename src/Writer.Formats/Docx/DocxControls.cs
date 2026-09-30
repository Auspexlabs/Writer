using System.Globalization;
using System.Security;
using System.Text.Json.Nodes;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml;
using Writer.Core;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Writer.Formats.Docx;

sealed class DocxControl(DocxDocument doc,W.Paragraph paragraph, W.SdtRun element) : Node
{
    public override string Kind => "control";
    public override object Anchor => element;
    public override IReadOnlyDictionary<string, string> GetProps() => new Dictionary<string, string> {
        ["id"] = element.SdtProperties?.GetFirstChild<W.SdtId>()?.Val?.Value.ToString(CultureInfo.InvariantCulture) ?? "0",
        ["data"] = DocxControls.Read(element).ToJsonString(),
        ["text"] = element.SdtContentRun?.InnerText ?? "",
        ["at"] = DocxFootnotes.OffsetOf(paragraph,element).ToString(CultureInfo.InvariantCulture)
    };
    public override string GetRaw() => element.OuterXml;
    public override void SetProp(string name,string value) {
        if(name=="data")DocxControls.Write(doc,element,value);
        else if(name=="at"){element.Remove();DocxFootnotes.Place(paragraph,element,int.Parse(value,CultureInfo.InvariantCulture));}
    }
    public override void Remove(){var before=element.PreviousSibling();var after=element.NextSibling();element.Remove();DocxRuns.Rejoin(before,after);}
}

static class DocxControls
{
    const string Wn="http://schemas.openxmlformats.org/wordprocessingml/2006/main", W14="http://schemas.microsoft.com/office/word/2010/wordml";
    static string Esc(string? value)=>SecurityElement.Escape(value??"")??"";
    internal static bool Supported(W.SdtRun element) => !DocxCitations.IsCitation(element)
        && element.SdtContentRun is {} content && content.ChildElements.All(e=>e is W.Run r&&r.ChildElements.All(c=>c is W.RunProperties||DocxRuns.IsTextElement(c)))
        && element.SdtProperties is {} props && props.ChildElements.Any(e=>e.LocalName is "text" or "richText" or "checkbox" or "date" or "dropDownList" or "comboBox") && !props.ChildElements.Any(e=>e.LocalName is "picture" or "docPartObj" or "docPartList" or "group" or "equation");
    internal static void EnsureIds(DocxDocument doc) {
        var all=doc.Main.Document!.Descendants<W.SdtRun>().Where(Supported).ToArray();var used=doc.Main.Document.Descendants<W.SdtId>().Select(x=>x.Val?.Value??0).ToHashSet();var seen=new HashSet<int>();
        for(var i=0;i<all.Length;i++) {
            var props=all[i].SdtProperties!;var id=props.GetFirstChild<W.SdtId>()?.Val?.Value;
            if(id is not null&&seen.Add(id.Value))continue;
            uint hash=2166136261;foreach(var ch in all[i].OuterXml+"/"+i.ToString(CultureInfo.InvariantCulture))hash=unchecked((hash^ch)*16777619);var next=(int)(hash&0x7fffffff);while(next==0||used.Contains(next))next=next==int.MaxValue?1:next+1;
            props.RemoveAllChildren<W.SdtId>();props.Append(new W.SdtId{Val=next});used.Add(next);seen.Add(next);
        }
    }
    internal static IEnumerable<Node> In(DocxDocument doc,W.Paragraph paragraph)=>paragraph.Descendants<W.SdtRun>().Where(s=>Supported(s)&&s.Ancestors<W.Paragraph>().First()==paragraph).Select(s=>(Node)new DocxControl(doc,paragraph,s));
    static string? Attr(OpenXmlElement? e,string name)=>e?.GetAttributes().FirstOrDefault(a=>a.LocalName==name).Value;
    internal static JsonObject Read(W.SdtRun element) {
        var props=element.SdtProperties!;var kind=props.ChildElements.FirstOrDefault(e=>e.LocalName is "text" or "richText" or "checkbox" or "date" or "dropDownList" or "comboBox");
        var type=kind?.LocalName switch{"checkbox"=>"checkbox","date"=>"date","dropDownList"=>"dropdown","comboBox"=>"combo","richText"=>"rich",_=>"text"};
        var data=new JsonObject{["type"]=type,["title"]=props.GetFirstChild<W.SdtAlias>()?.Val?.Value??"",["tag"]=props.GetFirstChild<W.Tag>()?.Val?.Value??"",["value"]=props.GetFirstChild<W.ShowingPlaceholder>() is null?element.SdtContentRun?.InnerText??"":"",["properties"]=props.OuterXml};
        if(props.GetFirstChild<W.Lock>()?.Val?.InnerText is {} locked)data["lock"]=locked;
        if(props.GetFirstChild<W.DataBinding>() is not null)data["bound"]=true;
        if(type=="checkbox")data["checked"]=Attr(kind?.ChildElements.FirstOrDefault(e=>e.LocalName=="checked"),"val") is "1" or "true" or "on";
        if(type=="date"){data["date"]=Attr(kind,"fullDate")?.Split('T')[0]??"";data["dateFormat"]=Attr(kind?.ChildElements.FirstOrDefault(e=>e.LocalName=="dateFormat"),"val")??"yyyy-MM-dd";}
        if(type is "dropdown" or "combo")data["items"]=new JsonArray(kind!.ChildElements.Where(e=>e.LocalName=="listItem").Select(e=>(JsonNode)new JsonObject{["value"]=Attr(e,"value")??"",["label"]=Attr(e,"displayText")??Attr(e,"value")??""}).ToArray());
        return data;
    }
    internal static Node Add(DocxDocument doc,W.Paragraph paragraph,IReadOnlyDictionary<string,string> props) {
        var element=new W.SdtRun(new W.SdtProperties(),new W.SdtContentRun());Write(doc,element,props.GetValueOrDefault("data")??"{}");
        var used=doc.Main.Document!.Descendants<W.SdtId>().Select(x=>x.Val?.Value??0).ToHashSet();int id;do{id=Random.Shared.Next(1,int.MaxValue);}while(used.Contains(id));element.SdtProperties!.RemoveAllChildren<W.SdtId>();element.SdtProperties.Append(new W.SdtId{Val=id});
        DocxFootnotes.Place(paragraph,element,props.TryGetValue("at",out var at)?int.Parse(at,CultureInfo.InvariantCulture):int.MaxValue);
        return new DocxControl(doc,paragraph,element);
    }
    internal static void Write(DocxDocument doc,W.SdtRun element,string json,bool updateBinding=true) {
        var data=JsonNode.Parse(json) as JsonObject??throw Bad("Expected a content control object");string Get(string key,string fallback="")=>data[key]?.GetValue<string>()??fallback;
        var type=Get("type","text");if(type is not ("text" or "rich" or "checkbox" or "date" or "dropdown" or "combo"))throw Bad("Unsupported content control type");
        var props=data["properties"] is {} raw?new W.SdtProperties(raw.GetValue<string>()):(W.SdtProperties?)element.SdtProperties?.CloneNode(true)??new W.SdtProperties();
        var id=element.SdtProperties?.GetFirstChild<W.SdtId>()?.Val?.Value;if(id is not null){props.RemoveAllChildren<W.SdtId>();props.Append(new W.SdtId{Val=id});}
        foreach(var child in props.ChildElements.Where(e=>e.LocalName is "text" or "richText" or "checkbox" or "date" or "dropDownList" or "comboBox" or "showingPlcHdr").ToArray())child.Remove();
        props.RemoveAllChildren<W.SdtAlias>();props.RemoveAllChildren<W.Tag>();props.Append(new W.SdtAlias{Val=Get("title")});props.Append(new W.Tag{Val=Get("tag")});var value=Get("value");string? xml=null;
        if(type=="checkbox") {var check=data["checked"]?.GetValue<bool>()==true;value=check?"☒":"☐";xml=$"<w14:checkbox xmlns:w14=\"{W14}\"><w14:checked w14:val=\"{(check?1:0)}\"/><w14:checkedState w14:val=\"2612\" w14:font=\"MS Gothic\"/><w14:uncheckedState w14:val=\"2610\" w14:font=\"MS Gothic\"/></w14:checkbox>";}
        else if(type=="date") {var date=Get("date");var format=Get("dateFormat","yyyy-MM-dd");if(date.Length>0){if(!DateTime.TryParseExact(date,"yyyy-MM-dd",CultureInfo.InvariantCulture,DateTimeStyles.None,out var parsed))throw Bad("Invalid date");try{value=parsed.ToString(format,CultureInfo.InvariantCulture);}catch(FormatException){throw Bad("Invalid date format");}}else value="";xml=$"<w:date xmlns:w=\"{Wn}\"{(date.Length>0?$" w:fullDate=\"{date}T00:00:00Z\"":"")}><w:dateFormat w:val=\"{Esc(format)}\"/><w:lid w:val=\"en-US\"/><w:storeMappedDataAs w:val=\"dateTime\"/><w:calendar w:val=\"gregorian\"/></w:date>";}
        else if(type is "dropdown" or "combo") {var items=data["items"] as JsonArray??[];if(items.Count>1000)throw Bad("Too many list choices");var tag=type=="dropdown"?"dropDownList":"comboBox";if(type=="dropdown"&&value.Length>0&&!items.Any(i=>(i?["label"]?.GetValue<string>()??i?["value"]?.GetValue<string>()??"")==value))throw Bad("Choose an item from the drop-down list");xml=$"<w:{tag} xmlns:w=\"{Wn}\">{string.Concat(items.Select(i=>$"<w:listItem w:displayText=\"{Esc(i?["label"]?.GetValue<string>()??i?["value"]?.GetValue<string>())}\" w:value=\"{Esc(i?["value"]?.GetValue<string>())}\"/>"))}</w:{tag}>";}
        else if(type=="text")props.Append(new W.SdtContentText{MultiLine=true});else props.Append(new W.SdtContentRichText());
        if(xml is not null){var wrapper=new W.SdtProperties($"<w:sdtPr xmlns:w=\"{Wn}\">{xml}</w:sdtPr>");props.Append(wrapper.FirstChild!.CloneNode(true));}
        var binding=updateBinding?PrepareBinding(doc,element,props,data,value):null;
        if(value.Length==0){props.Append(new W.ShowingPlaceholder());value=Get("title","点击填写");if(value.Length==0)value="点击填写";}
        var style=element.SdtContentRun?.Elements<W.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);var run=new W.Run();if(style is not null)run.Append(style);run.Append(DocxRuns.TextElements(value));
        element.SdtProperties=props;element.SdtContentRun=new W.SdtContentRun(run);
        binding?.Invoke();
    }
    static Action? PrepareBinding(DocxDocument doc,W.SdtRun element,W.SdtProperties props,JsonObject data,string value) {
        if(props.GetFirstChild<W.DataBinding>() is not {} binding)return null;
        var store=binding.StoreItemId?.Value;var xpath=binding.XPath?.Value;
        if(string.IsNullOrWhiteSpace(store)||string.IsNullOrWhiteSpace(xpath))throw Bad("The content control has an incomplete XML binding");
        var part=doc.Main.CustomXmlParts.FirstOrDefault(p=>string.Equals(p.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value,store,StringComparison.OrdinalIgnoreCase));
        if(part is null)throw Bad("The content control's XML data source is missing");
        var xml=new XmlDocument{PreserveWhitespace=true,XmlResolver=null};using(var stream=part.GetStream(FileMode.Open,FileAccess.Read))using(var reader=XmlReader.Create(stream,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))xml.Load(reader);
        var namespaces=new XmlNamespaceManager(xml.NameTable);
        try {
            var declarations=XElement.Parse("<bindings "+(binding.PrefixMappings?.Value??"")+"/>");
            foreach(var attribute in declarations.Attributes().Where(a=>a.IsNamespaceDeclaration))if(attribute.Name.LocalName!="xmlns")namespaces.AddNamespace(attribute.Name.LocalName,attribute.Value);
        }catch(XmlException){throw Bad("The content control has invalid XML namespace mappings");}
        XmlNode? target;try{target=xml.SelectSingleNode(xpath,namespaces);}catch(XPathException){throw Bad("The content control has an invalid XPath binding");}
        if(target is not (XmlElement or XmlAttribute or XmlText)||target is XmlElement node&&node.ChildNodes.OfType<XmlElement>().Any())throw Bad("The content control binding must select a text value or attribute");
        var type=data["type"]?.GetValue<string>()??"text";
        var bound=type=="checkbox"?(data["checked"]?.GetValue<bool>()==true?"true":"false"):type=="date"?(data["date"]?.GetValue<string>() is {Length:>0} date?date+"T00:00:00Z":""):value;
        if(type=="dropdown")bound=(data["items"] as JsonArray)?.FirstOrDefault(i=>(i?["label"]?.GetValue<string>()??i?["value"]?.GetValue<string>())==value)?["value"]?.GetValue<string>()??value;
        target.InnerText=bound;
        var mirrors=doc.Main.Document!.Descendants<W.SdtRun>().Where(s=>s!=element&&Supported(s)&&s.SdtProperties?.GetFirstChild<W.DataBinding>() is {} other&&string.Equals(other.StoreItemId?.Value,store,StringComparison.OrdinalIgnoreCase)&&other.XPath?.Value==xpath&&other.PrefixMappings?.Value==binding.PrefixMappings?.Value).ToArray();
        var updates=new List<(W.SdtRun Element,string Data)>();
        foreach(var mirror in mirrors){var next=Read(mirror);var kind=next["type"]!.GetValue<string>();next["value"]=bound;if(kind=="checkbox")next["checked"]=bound is "1" or "true" or "on";if(kind=="date")next["date"]=bound.Split('T')[0];if(kind=="dropdown")next["value"]=(next["items"] as JsonArray)?.FirstOrDefault(i=>i?["value"]?.GetValue<string>()==bound)?["label"]?.GetValue<string>()??bound;var clone=(W.SdtRun)mirror.CloneNode(true);Write(doc,clone,next.ToJsonString(),false);updates.Add((mirror,next.ToJsonString()));}
        return ()=>{using var bytes=new MemoryStream();using(var writer=XmlWriter.Create(bytes,new XmlWriterSettings{Encoding=new System.Text.UTF8Encoding(false),CloseOutput=false}))xml.Save(writer);bytes.Position=0;part.FeedData(bytes);foreach(var update in updates)Write(doc,update.Element,update.Data,false);};
    }
    static WriterException Bad(string message)=>new(ErrorCode.Validation,message,"Use {type,title,tag,value,checked,date,dateFormat,items:[{label,value}]}.");
}
