using System.Globalization;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Writer.Core;
using P = DocumentFormat.OpenXml.Presentation;

namespace Writer.Formats.Pptx;

/// <summary>A slide's animations: the main sequence of its p:timing as a list, one effect per entry, in the order they play. Each
/// names its shape by drawing id, a preset PowerPoint knows (the same presetID / presetClass / presetSubtype and behaviours it
/// writes, so PowerPoint plays and lists them as its own), how it starts (on a click, with or after the previous effect), its length
/// and delay. An effect the engine does not model reads as effect other with its markup, which a write puts back unchanged.</summary>
static class PptxAnim
{
    sealed record Preset(string Class, int Id, int Subtype, int Duration, Func<string, int, string> Body);

    const string Ns = PptxTemplate.Ns;

    static string Tgt(string s) => $"<p:tgtEl><p:spTgt spid=\"{s}\"/></p:tgtEl>";
    static string Ctn(int d, bool hold = false) => $"<p:cTn id=\"0\" dur=\"{d}\"{(hold ? " fill=\"hold\"" : "")}/>";
    static string Attr(string name) => $"<p:attrNameLst><p:attrName>{name}</p:attrName></p:attrNameLst>";
    static string Str(string v) => $"<p:strVal val=\"{v}\"/>";
    static string Flt(string v) => $"<p:fltVal val=\"{v}\"/>";
    static string Vis(string s, string val, int delay) =>
        $"<p:set><p:cBhvr><p:cTn id=\"0\" dur=\"1\" fill=\"hold\"><p:stCondLst><p:cond delay=\"{delay}\"/></p:stCondLst></p:cTn>{Tgt(s)}{Attr("style.visibility")}</p:cBhvr><p:to>{Str(val)}</p:to></p:set>";
    static string Effect(string s, int d, string transition, string filter) =>
        $"<p:animEffect transition=\"{transition}\" filter=\"{filter}\"><p:cBhvr>{Ctn(d)}{Tgt(s)}</p:cBhvr></p:animEffect>";
    static string Anim(string s, int d, string attr, string from, string to) =>
        $"<p:anim calcmode=\"lin\" valueType=\"num\"><p:cBhvr additive=\"base\">{Ctn(d, true)}{Tgt(s)}{Attr(attr)}</p:cBhvr><p:tavLst><p:tav tm=\"0\"><p:val>{from}</p:val></p:tav><p:tav tm=\"100000\"><p:val>{to}</p:val></p:tav></p:tavLst></p:anim>";

    /// <summary>The modeled effects by name. Fly and wipe come from the bottom, as PowerPoint's defaults do.</summary>
    static readonly Dictionary<string, Preset> Presets = new()
    {
        ["appear"] = new("entr", 1, 0, 1, (s, d) => Vis(s, "visible", 0)),
        ["fade"] = new("entr", 10, 0, 500, (s, d) => Vis(s, "visible", 0) + Effect(s, d, "in", "fade")),
        ["fly"] = new("entr", 2, 4, 500, (s, d) => Vis(s, "visible", 0) + Anim(s, d, "ppt_x", Str("#ppt_x"), Str("#ppt_x")) + Anim(s, d, "ppt_y", Str("1+#ppt_h/2"), Str("#ppt_y"))),
        ["float"] = new("entr", 42, 0, 1000, (s, d) => Vis(s, "visible", 0) + Effect(s, d, "in", "fade") + Anim(s, d, "ppt_x", Str("#ppt_x"), Str("#ppt_x")) + Anim(s, d, "ppt_y", Str("#ppt_y+.1"), Str("#ppt_y"))),
        ["zoom"] = new("entr", 53, 16, 500, (s, d) => Vis(s, "visible", 0) + Anim(s, d, "ppt_w", Flt("0"), Str("#ppt_w")) + Anim(s, d, "ppt_h", Flt("0"), Str("#ppt_h")) + Effect(s, d, "in", "fade")),
        ["wipe"] = new("entr", 22, 4, 500, (s, d) => Vis(s, "visible", 0) + Effect(s, d, "in", "wipe(up)")),
        ["grow"] = new("emph", 6, 0, 2000, (s, d) => $"<p:animScale><p:cBhvr>{Ctn(d, true)}{Tgt(s)}</p:cBhvr><p:by x=\"150000\" y=\"150000\"/></p:animScale>"),
        ["spin"] = new("emph", 8, 0, 2000, (s, d) => $"<p:animRot by=\"21600000\"><p:cBhvr>{Ctn(d, true)}{Tgt(s)}{Attr("r")}</p:cBhvr></p:animRot>"),
        ["transparency"] = new("emph", 9, 0, 2000, (s, d) => $"<p:set><p:cBhvr>{Ctn(d, true)}{Tgt(s)}{Attr("style.opacity")}</p:cBhvr><p:to>{Str("0.5")}</p:to></p:set><p:animEffect filter=\"alpha\"><p:cBhvr>{Ctn(d, true)}{Tgt(s)}</p:cBhvr><p:progress>{Flt("0.5")}</p:progress></p:animEffect>"),
        ["disappear"] = new("exit", 1, 0, 1, (s, d) => Vis(s, "hidden", 0)),
        ["fadeOut"] = new("exit", 10, 0, 500, (s, d) => Effect(s, d, "out", "fade") + Vis(s, "hidden", d - 1)),
        ["flyOut"] = new("exit", 2, 4, 500, (s, d) => Anim(s, d, "ppt_x", Str("#ppt_x"), Str("#ppt_x")) + Anim(s, d, "ppt_y", Str("#ppt_y"), Str("1+#ppt_h/2")) + Vis(s, "hidden", d - 1)),
        ["zoomOut"] = new("exit", 53, 16, 500, (s, d) => Anim(s, d, "ppt_w", Str("#ppt_w"), Flt("0")) + Anim(s, d, "ppt_h", Str("#ppt_h"), Flt("0")) + Effect(s, d, "out", "fade") + Vis(s, "hidden", d - 1)),
        ["wipeOut"] = new("exit", 22, 4, 500, (s, d) => Effect(s, d, "out", "wipe(down)") + Vis(s, "hidden", d - 1)),
    };

    internal static IEnumerable<string> Names => Presets.Keys;

    static readonly Dictionary<string, string> Classes = new() { ["entr"] = "entrance", ["emph"] = "emphasis", ["exit"] = "exit", ["path"] = "path" };
    static readonly Dictionary<string, string> Starts = new() { ["clickEffect"] = "click", ["withEffect"] = "with", ["afterEffect"] = "after" };

    sealed record Fx(string? Shape, string Effect, string Start, int Duration, int Delay, P.ParallelTimeNode? Other);

    static P.SequenceTimeNode? MainSeq(P.Timing? timing) =>
        timing?.Descendants<P.SequenceTimeNode>().FirstOrDefault(s => s.CommonTimeNode?.NodeType?.InnerText == "mainSeq");

    static bool IsEffect(P.CommonTimeNode c) => c.NodeType?.InnerText is { } t && Starts.ContainsKey(t);

    static int Ms(string? v) => int.TryParse(v, NumberStyles.None, CultureInfo.InvariantCulture, out var ms) ? ms : 0;

    /// <summary>How long an effect runs after its delay: its longest behaviour, with the behaviour's own delay.</summary>
    static int Span(P.CommonTimeNode effect) => effect.Descendants<P.CommonTimeNode>()
        .Select(c => Ms(c.Duration?.Value) + Ms(c.StartConditionList?.GetFirstChild<P.Condition>()?.Delay?.Value)).DefaultIfEmpty(0).Max();

    /// <summary>The slide's effects as the animations prop gives them, or null when it has none.</summary>
    internal static string? Read(SlidePart slide)
    {
        var effects = MainSeq(slide.Slide?.Timing)?.Descendants<P.CommonTimeNode>().Where(IsEffect).ToList();
        if (effects is not { Count: > 0 }) return null;
        return NodeJson.Compact(w =>
        {
            w.WriteStartArray();
            foreach (var c in effects)
            {
                var targets = c.Descendants<P.ShapeTarget>().ToList();
                var spids = targets.Select(t => t.ShapeId?.Value).Distinct().ToList();
                var shape = spids.Count == 1 ? spids[0] : null;
                var whole = shape is not null && targets.All(t => !t.HasChildren);
                var cls = c.PresetClass?.InnerText;
                var name = whole ? Presets.FirstOrDefault(p => p.Value.Class == cls && p.Value.Id == (c.PresetId?.Value ?? -1) && p.Value.Subtype == (c.PresetSubtype?.Value ?? 0)).Key : null;
                w.WriteStartObject();
                if (shape is not null) w.WriteString("shape", shape);
                w.WriteString("effect", name ?? "other");
                if (name is null) w.WriteString("class", cls is not null && Classes.TryGetValue(cls, out var cl) ? cl : "other");
                w.WriteString("start", Starts[c.NodeType!.InnerText!]);
                w.WriteNumber("duration", Math.Max(1, Span(c)));
                w.WriteNumber("delay", Ms(c.StartConditionList?.GetFirstChild<P.Condition>()?.Delay?.Value));
                if (name is null) w.WriteString("xml", c.Parent!.OuterXml);
                w.WriteEndObject();
            }
            w.WriteEndArray();
        });
    }

    static List<Fx> Parse(SlidePart slide, string json)
    {
        JsonDocument parsed;
        try { parsed = JsonDocument.Parse(json); }
        catch (JsonException) { throw Invalid("animations must be a JSON array of effects"); }
        using (parsed)
        {
            if (parsed.RootElement.ValueKind != JsonValueKind.Array) throw Invalid("animations must be a JSON array of effects");
            var ids = PptxDocument.ShapeTree(slide).Descendants<P.NonVisualDrawingProperties>().Select(p => p.Id?.Value.ToString(CultureInfo.InvariantCulture)).ToHashSet();
            var list = new List<Fx>();
            foreach (var e in parsed.RootElement.EnumerateArray())
            {
                if (e.ValueKind != JsonValueKind.Object) throw Invalid("each effect is a JSON object");
                string? Get(string k) => e.TryGetProperty(k, out var v) ? v.ValueKind == JsonValueKind.String ? v.GetString() : v.GetRawText() : null;
                var effect = Get("effect") ?? "";
                var start = Get("start") ?? "click";
                if (start is not ("click" or "with" or "after")) throw Invalid($"start '{start}': use click, with or after");
                var delay = Ms(Get("delay"));
                if (effect == "other")
                {
                    var xml = Get("xml") ?? throw Invalid("an effect other keeps its xml as read");
                    list.Add(new(Get("shape"), effect, start, 0, delay, new P.ParallelTimeNode(xml)));
                    continue;
                }
                if (!Presets.TryGetValue(effect, out var preset)) throw Invalid($"No effect called '{effect}'");
                var shape = Get("shape");
                if (shape is null || !ids.Contains(shape)) throw Invalid($"No shape with id '{shape}' on this slide");
                list.Add(new(shape, effect, start, Get("duration") is { } d ? Math.Max(1, Ms(d)) : preset.Duration, delay, null));
            }
            return list;
        }
    }

    static WriterException Invalid(string message) => new(ErrorCode.Validation, message,
        $"Example: [{{\"shape\":\"4\",\"effect\":\"fade\",\"start\":\"click\",\"duration\":500,\"delay\":0}}]; effects: {string.Join(", ", Presets.Keys)}.");

    /// <summary>Replaces the main sequence with the effects the JSON lists, grouped the way PowerPoint groups them: a click starts a
    /// group, an effect after the previous one a step in it that starts when the step before ends. Interactive sequences and media
    /// timing stay; an empty list removes the main sequence.</summary>
    internal static void Write(SlidePart slide, string json)
    {
        var list = Parse(slide, json);
        var sld = slide.Slide!;
        var timing = sld.Timing;
        var seq = MainSeq(timing);
        if (list.Count == 0)
        {
            seq?.Remove();
            if (timing is not null) Tidy(slide, timing);
            return;
        }
        if (timing is null)
        {
            timing = new P.Timing($"<p:timing {Ns}><p:tnLst><p:par><p:cTn id=\"1\" dur=\"indefinite\" restart=\"never\" nodeType=\"tmRoot\"><p:childTnLst/></p:cTn></p:par></p:tnLst></p:timing>");
            if (sld.SlideExtensionList is { } ext) sld.InsertBefore(timing, ext); else sld.Append(timing);
        }
        if (seq is null)
        {
            var root = timing.TimeNodeList!.GetFirstChild<P.ParallelTimeNode>()!.CommonTimeNode!.ChildTimeNodeList ??= new P.ChildTimeNodeList();
            seq = new P.SequenceTimeNode($"<p:seq {Ns} concurrent=\"1\" nextAc=\"seek\"><p:cTn id=\"999999\" dur=\"indefinite\" nodeType=\"mainSeq\"><p:childTnLst/></p:cTn>" +
                "<p:prevCondLst><p:cond evt=\"onPrev\" delay=\"0\"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt=\"onNext\" delay=\"0\"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>");
            root.PrependChild(seq);
        }
        var main = seq.CommonTimeNode!;
        var groups = main.ChildTimeNodeList ??= new P.ChildTimeNodeList();
        groups.RemoveAllChildren();

        // grpId: an effect's number among its shape's effects, never one a kept effect of that shape already uses
        var used = list.Where(f => f.Other is not null).Select(f => (f.Other!.CommonTimeNode?.Descendants<P.ShapeTarget>().FirstOrDefault()?.ShapeId?.Value, f.Other.CommonTimeNode?.GroupId?.Value ?? 0u)).ToHashSet();
        P.ParallelTimeNode? group = null;
        P.ChildTimeNodeList? step = null;
        int at = 0, end = 0;
        for (var i = 0; i < list.Count; i++)
        {
            var f = list[i];
            if (step is null || f.Start == "click")
            {
                var begin = i == 0 && f.Start != "click" ? $"<p:cond evt=\"onBegin\" delay=\"0\"><p:tn val=\"{main.Id!.Value}\"/></p:cond>" : ""; // plays when the slide starts
                group = new P.ParallelTimeNode($"<p:par {Ns}><p:cTn id=\"0\" fill=\"hold\"><p:stCondLst><p:cond delay=\"indefinite\"/>{begin}</p:stCondLst><p:childTnLst/></p:cTn></p:par>");
                groups.Append(group);
                at = end = 0;
                step = Step(group!, at);
            }
            else if (f.Start == "after")
            {
                at = end;
                step = Step(group!, at);
            }
            P.ParallelTimeNode par;
            if (f.Other is { } other)
            {
                par = other;
                var ctn = par.CommonTimeNode!;
                ctn.NodeType = f.Start switch { "click" => P.TimeNodeValues.ClickEffect, "with" => P.TimeNodeValues.WithEffect, _ => P.TimeNodeValues.AfterEffect };
                var conds = ctn.StartConditionList ??= new P.StartConditionList();
                (conds.GetFirstChild<P.Condition>() ?? conds.AppendChild(new P.Condition())).Delay = f.Delay.ToString(CultureInfo.InvariantCulture);
                end = Math.Max(end, at + f.Delay + Span(ctn));
            }
            else
            {
                var p = Presets[f.Effect];
                var grp = 0u;
                while (!used.Add((f.Shape, grp))) grp++;
                par = new P.ParallelTimeNode($"<p:par {Ns}><p:cTn id=\"0\" presetID=\"{p.Id}\" presetClass=\"{p.Class}\" presetSubtype=\"{p.Subtype}\" fill=\"hold\" grpId=\"{grp}\" nodeType=\"{f.Start}Effect\">" +
                    $"<p:stCondLst><p:cond delay=\"{f.Delay}\"/></p:stCondLst><p:childTnLst>{p.Body(f.Shape!, f.Duration)}</p:childTnLst></p:cTn></p:par>");
                end = Math.Max(end, at + f.Delay + f.Duration);
            }
            step!.Append(par);
        }
        Renumber(timing);
        Tidy(slide, timing);
    }

    /// <summary>A step of a click group: a par starting at ms after the click, whose children are the effects.</summary>
    static P.ChildTimeNodeList Step(P.ParallelTimeNode group, int at)
    {
        var step = new P.ParallelTimeNode($"<p:par {Ns}><p:cTn id=\"0\" fill=\"hold\"><p:stCondLst><p:cond delay=\"{at}\"/></p:stCondLst><p:childTnLst/></p:cTn></p:par>");
        group.CommonTimeNode!.ChildTimeNodeList!.Append(step);
        return step.CommonTimeNode!.ChildTimeNodeList!;
    }

    /// <summary>Numbers every time node 1, 2, … in document order, as PowerPoint does, and points each reference at its node's new number.</summary>
    static void Renumber(P.Timing timing)
    {
        var map = new Dictionary<uint, uint>();
        uint next = 1;
        foreach (var c in timing.Descendants<P.CommonTimeNode>())
        {
            if (c.Id?.Value is { } old and > 0) map.TryAdd(old, next);
            c.Id = next++;
        }
        foreach (var tn in timing.Descendants<P.TimeNode>())
            if (tn.Val?.Value is { } v && map.TryGetValue(v, out var n)) tn.Val = n;
    }

    /// <summary>Drops the effects whose shape the slide no longer has (a removed shape takes its animations with it), then what that
    /// leaves empty; the save runs it on every slide, so a timing never names a missing shape.</summary>
    internal static void Prune(SlidePart slide)
    {
        if (slide.Slide?.Timing is not { } timing) return;
        var ids = PptxDocument.ShapeTree(slide).Descendants<P.NonVisualDrawingProperties>().Select(p => p.Id?.Value).ToHashSet();
        var gone = timing.Descendants<P.ShapeTarget>().Where(t => !ids.Contains(t.ShapeId?.Value is { } s && uint.TryParse(s, out var n) ? n : null)).ToList();
        if (gone.Count == 0 && timing.BuildList is null) return;
        foreach (var t in gone)
        {
            OpenXmlElement? owner = t.Ancestors<P.ParallelTimeNode>().FirstOrDefault(p => p.CommonTimeNode is { } c && IsEffect(c));
            owner ??= t.Ancestors<P.SequenceTimeNode>().FirstOrDefault(); // an interactive sequence its shape triggered
            if (owner?.Parent is not null) owner.Remove(); // else it went with an effect removed already
        }
        Tidy(slide, timing);
    }

    /// <summary>Removes what holds nothing any more (a click group without effects, an empty sequence, the timing itself) and keeps the
    /// build list to the effects left: an entry per shape and group id, a new one for a shape (bldP) or table (bldGraphic) PowerPoint builds.</summary>
    static void Tidy(SlidePart slide, P.Timing timing)
    {
        bool Empty(OpenXmlElement e) => e is P.ParallelTimeNode { CommonTimeNode: { } c } && !IsEffect(c) && c.NodeType?.InnerText != "tmRoot" && c.ChildTimeNodeList is { HasChildren: false }
            || e is P.SequenceTimeNode { CommonTimeNode.ChildTimeNodeList.HasChildren: false };
        for (var hollow = timing.Descendants().Where(Empty).ToList(); hollow.Count > 0; hollow = timing.Descendants().Where(Empty).ToList())
            foreach (var e in hollow) e.Remove();
        foreach (var list in timing.Descendants<P.ChildTimeNodeList>().Where(l => !l.HasChildren).ToList()) list.Remove();
        var root = timing.TimeNodeList?.GetFirstChild<P.ParallelTimeNode>()?.CommonTimeNode;
        if (root?.ChildTimeNodeList is not { HasChildren: true })
        {
            timing.Remove();
            return;
        }
        var effects = timing.Descendants<P.CommonTimeNode>().Where(IsEffect)
            .Select(c => (Shape: c.Descendants<P.ShapeTarget>().FirstOrDefault()?.ShapeId?.Value, Group: (c.GroupId?.Value ?? 0u).ToString(CultureInfo.InvariantCulture)))
            .Where(e => e.Shape is not null).Distinct().ToList();
        string? Attr(OpenXmlElement e, string name) => e.GetAttributes().FirstOrDefault(a => a.LocalName == name).Value;
        var builds = timing.BuildList ??= new P.BuildList();
        foreach (var b in builds.ChildElements.Where(b => !effects.Contains((Attr(b, "spid"), Attr(b, "grpId") ?? "0"))).ToList()) b.Remove();
        var tree = PptxDocument.ShapeTree(slide);
        foreach (var (shape, group) in effects)
        {
            if (builds.ChildElements.Any(b => Attr(b, "spid") == shape && (Attr(b, "grpId") ?? "0") == group)) continue;
            var target = tree.Descendants<P.NonVisualDrawingProperties>().FirstOrDefault(p => p.Id?.Value.ToString(CultureInfo.InvariantCulture) == shape)?.Parent?.Parent;
            if (target is P.Shape) builds.Append(new P.BuildParagraph($"<p:bldP {Ns} spid=\"{shape}\" grpId=\"{group}\" animBg=\"1\"/>"));
            else if (target is P.GraphicFrame) builds.Append(new P.BuildGraphics($"<p:bldGraphic {Ns} spid=\"{shape}\" grpId=\"{group}\"><p:bldAsOne/></p:bldGraphic>"));
        }
        if (!builds.HasChildren) builds.Remove();
    }
}
