using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Writer.Tests;

/// <summary>Part-by-part comparison of two Office packages. XML parts are compared after canonicalisation,
/// everything else byte for byte.</summary>
public static class PackageCompare
{
    /// <summary>Null when the packages are equivalent, otherwise a message naming the first difference.</summary>
    public static string? Diff(byte[] a, byte[] b, params string[] exceptParts)
    {
        var za = Parts(a);
        var zb = Parts(b);
        var onlyA = za.Keys.Except(zb.Keys).Except(exceptParts).ToList();
        var onlyB = zb.Keys.Except(za.Keys).Except(exceptParts).ToList();
        if (onlyA.Count > 0 || onlyB.Count > 0)
            return $"part lists differ: missing [{string.Join(", ", onlyA)}] added [{string.Join(", ", onlyB)}]";
        foreach (var name in za.Keys.OrderBy(k => k, StringComparer.Ordinal))
        {
            if (exceptParts.Contains(name) || !zb.ContainsKey(name)) continue;
            if (IsXml(name))
            {
                var ca = Canonical(za[name]);
                var cb = Canonical(zb[name]);
                if (ca != cb)
                {
                    var i = 0;
                    while (i < ca.Length && i < cb.Length && ca[i] == cb[i]) i++;
                    return $"{name} differs near offset {i}: original …{Snip(ca, i)}… saved …{Snip(cb, i)}…";
                }
            }
            else if (!za[name].AsSpan().SequenceEqual(zb[name])) return $"{name}: bytes differ";
        }
        return null;
    }

    public static Dictionary<string, byte[]> Parts(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package));
        return zip.Entries.ToDictionary(e => e.FullName, e =>
        {
            using var s = e.Open();
            var m = new MemoryStream();
            s.CopyTo(m);
            return m.ToArray();
        });
    }

    public static bool IsXml(string partName) => partName.EndsWith(".xml") || partName.EndsWith(".rels");

    public static string Canonical(byte[] xml) => Element(Root(xml));

    static XElement Root(byte[] xml) => XDocument.Parse(Encoding.UTF8.GetString(xml).TrimStart('﻿')).Root!;

    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Compares word/document.xml after removing the block (paragraph or table) at blockIndex from both bodies.</summary>
    public static string? DiffBodyExcept(byte[] a, byte[] b, int blockIndex)
    {
        var ra = Root(Parts(a)["word/document.xml"]);
        var rb = Root(Parts(b)["word/document.xml"]);
        Block(ra, blockIndex).Remove();
        Block(rb, blockIndex).Remove();
        var ca = Element(ra);
        var cb = Element(rb);
        if (ca == cb) return null;
        var i = 0;
        while (i < ca.Length && i < cb.Length && ca[i] == cb[i]) i++;
        return $"word/document.xml differs outside block {blockIndex} near offset {i}: original …{Snip(ca, i)}… saved …{Snip(cb, i)}…";
    }

    static XElement Block(XElement root, int index) =>
        root.Element(W + "body")!.Elements().Where(e => e.Name == W + "p" || e.Name == W + "tbl").ElementAt(index);

    static void Write(XElement e, StringBuilder sb)
    {
        sb.Append('<').Append(e.Name);
        foreach (var a in e.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString(), StringComparer.Ordinal))
            sb.Append(' ').Append(a.Name).Append("=\"").Append(a.Value).Append('"');
        sb.Append('>');
        var children = e.Nodes().Select(n => n switch
        {
            XElement c => Element(c),
            XText t => System.Security.SecurityElement.Escape(t.Value),
            _ => "",
        });
        if (e.Name.LocalName is "Relationships" or "Types") children = children.OrderBy(c => c, StringComparer.Ordinal);
        foreach (var c in children) sb.Append(c);
        sb.Append("</>");
    }

    static string Element(XElement e)
    {
        var sb = new StringBuilder();
        Write(e, sb);
        return sb.ToString();
    }

    static string Snip(string s, int at)
    {
        var start = Math.Max(0, at - 40);
        return s.Substring(start, Math.Min(120, s.Length - start));
    }
}
