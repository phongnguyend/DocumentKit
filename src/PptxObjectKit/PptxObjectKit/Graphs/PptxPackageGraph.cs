using System.Xml.Linq;
using PptxObjectKit.Models;

namespace PptxObjectKit.Graphs;

internal sealed class PptxPackageGraph(PptxObjectDocument document)
{
    internal static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    internal static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    internal static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    internal static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    internal static string RelationshipPart(string part)
    {
        var slash = part.LastIndexOf('/');
        return (slash < 0 ? "" : part[..(slash + 1)]) + "_rels/" + part[(slash + 1)..] + ".rels";
    }

    internal IEnumerable<XElement> Relationships(string part)
    {
        var path = RelationshipPart(part);
        return document.Parts.ContainsKey(path) ? document.Xml(path).Root!.Elements(Rel + "Relationship") : [];
    }

    internal static string Target(string part, XElement relationship)
    {
        if ((string?)relationship.Attribute("TargetMode") == "External")
        {
            throw new NotSupportedException("External slide parts are not supported.");
        }
        var target = (string?)relationship.Attribute("Target") ?? throw new InvalidDataException("Missing relationship target.");
        var uri = new Uri(new Uri("https://package.invalid/" + part), target);
        if (uri.Host != "package.invalid" || uri.Scheme != "https" || uri.Query.Length != 0 || uri.Fragment.Length != 0)
        {
            throw new InvalidDataException("Invalid package relationship target.");
        }
        return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
    }

    internal string PresentationPart()
    {
        var root = Relationships("").SingleOrDefault(rel => ((string?)rel.Attribute("Type"))?.EndsWith("/officeDocument", StringComparison.Ordinal) == true)
            ?? throw new InvalidDataException("Missing presentation relationship.");
        return Target("", root);
    }

    internal List<(XElement Id, string Part)> Slides()
    {
        var part = PresentationPart();
        var presentation = document.Xml(part);
        if (presentation.Root?.Name != P + "presentation")
        {
            throw new NotSupportedException("Only Transitional OOXML PowerPoint presentations are supported.");
        }
        var relationships = Relationships(part).ToDictionary(rel => (string)rel.Attribute("Id")!);
        var result = new List<(XElement, string)>();
        foreach (var id in presentation.Root.Element(P + "sldIdLst")?.Elements(P + "sldId") ?? [])
        {
            if (!relationships.TryGetValue((string?)id.Attribute(R + "id") ?? "", out var relationship)
                || !((string?)relationship.Attribute("Type") ?? "").EndsWith("/slide", StringComparison.Ordinal))
            {
                throw new InvalidDataException("Slide relationship is missing or invalid.");
            }
            var path = Target(part, relationship);
            if (document.Xml(path).Root?.Name != P + "sld")
            {
                throw new InvalidDataException("Invalid slide XML.");
            }
            result.Add((id, path));
        }
        if (result.Count > 2000)
        {
            throw new NotSupportedException("Presentation exceeds 2,000 slides.");
        }
        return result;
    }
}
