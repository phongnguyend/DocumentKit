using System.Xml.Linq;
using PptxObjectKit.Models;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Graphs;

internal sealed class PptxTextGeometry
{
    private readonly XDocument? layout;
    private readonly XDocument? master;

    internal PptxTextGeometry(PptxObjectDocument document, string slidePart)
    {
        var graph = new PptxPackageGraph(document);
        string? Related(string part, string kind)
        {
            var relation = graph.Relationships(part).FirstOrDefault(item => ((string?)item.Attribute("Type"))?.EndsWith("/" + kind, StringComparison.Ordinal) == true);
            return relation is null ? null : Target(part, relation);
        }
        var layoutPart = Related(slidePart, "slideLayout");
        layout = layoutPart is null ? null : document.Xml(layoutPart);
        var masterPart = layoutPart is null ? null : Related(layoutPart, "slideMaster");
        master = masterPart is null ? null : document.Xml(masterPart);
    }

    private static XElement? Placeholder(XElement shape, XDocument? document)
    {
        var placeholder = shape.Descendants(P + "ph").FirstOrDefault();
        if (placeholder is null || document is null)
        {
            return null;
        }
        var candidates = document.Descendants(P + "sp").Where(item => item.Descendants(P + "ph").Any()).ToArray();
        return candidates.FirstOrDefault(item => ((string?)item.Descendants(P + "ph").First().Attribute("idx") ?? "0") == ((string?)placeholder.Attribute("idx") ?? "0"))
            ?? candidates.FirstOrDefault(item => (string?)item.Descendants(P + "ph").First().Attribute("type") == (string?)placeholder.Attribute("type"));
    }

    internal XElement? Frame(XElement shape)
    {
        var inherited = Placeholder(shape, layout);
        var baseShape = Placeholder(inherited ?? shape, master);
        return shape.Element(P + "spPr")?.Element(A + "xfrm")
            ?? inherited?.Element(P + "spPr")?.Element(A + "xfrm")
            ?? baseShape?.Element(P + "spPr")?.Element(A + "xfrm");
    }

    internal PptxTextBounds? Bounds(XElement shape)
    {
        var frame = Frame(shape);
        if (frame?.Element(A + "off") is not { } offset || frame.Element(A + "ext") is not { } size)
        {
            return null;
        }
        return new(((double?)offset.Attribute("x") ?? 0) / 12700, ((double?)offset.Attribute("y") ?? 0) / 12700,
            ((double?)size.Attribute("cx") ?? 0) / 12700, ((double?)size.Attribute("cy") ?? 0) / 12700,
            ((double?)frame.Attribute("rot") ?? 0) / 60000, shape.Ancestors(P + "grpSp").Any());
    }
}
