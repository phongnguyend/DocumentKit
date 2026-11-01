using System.Xml.Linq;
using PptxObjectKit.Graphs;
using PptxObjectKit.Models;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Editors;

/// <summary>Atomic edits: each operation returns a new document.</summary>
public static class PptxSlideEditor
{
    private static PptxObjectDocument Editable(PptxObjectDocument source)
    {
        if (source.PartNames.Any(path => path.StartsWith("_xmlsignatures/", StringComparison.Ordinal)))
        {
            throw new NotSupportedException("Editing signed presentations is not supported.");
        }
        return source.Clone();
    }

    public static PptxObjectDocument Reorder(PptxObjectDocument source, IReadOnlyList<int> order)
    {
        var document = Editable(source);
        var graph = new PptxPackageGraph(document);
        var count = graph.Slides().Count;
        if (order.Count != count || order.Distinct().Count() != count || order.Any(number => number < 1 || number > count))
        {
            throw new ArgumentException("Provide each slide number exactly once.", nameof(order));
        }
        var part = graph.PresentationPart();
        var xml = document.Xml(part);
        var list = xml.Root!.Element(P + "sldIdLst")!;
        var ids = list.Elements().ToArray();
        list.ReplaceNodes(order.Select(number => new XElement(ids[number - 1])));
        document.SetXml(part, xml);
        return document;
    }

    public static PptxObjectDocument Delete(PptxObjectDocument source, int number)
    {
        var document = Editable(source);
        var graph = new PptxPackageGraph(document);
        var slides = graph.Slides();
        if (number < 1 || number > slides.Count || slides.Count <= 1)
        {
            throw new ArgumentException("Choose a valid slide; the last slide cannot be deleted.");
        }
        var part = graph.PresentationPart();
        var xml = document.Xml(part);
        var id = xml.Root!.Element(P + "sldIdLst")!.Elements().ElementAt(number - 1);
        // Remove presentation-level references to the deleted slide (custom shows/sections).
        var relationshipId = (string)id.Attribute(R + "id")!;
        var slideId = (string)id.Attribute("id")!;
        id.Remove();
        foreach (var reference in xml.Descendants().Where(element =>
            (element.Name == P + "sld" && (string?)element.Attribute(R + "id") == relationshipId)
            || (element.Name.LocalName == "sldId" && (string?)element.Attribute("id") == slideId)).ToList())
        {
            reference.Remove();
        }
        var relPart = RelationshipPart(part);
        var rels = document.Xml(relPart);
        rels.Root!.Elements().Single(rel => (string?)rel.Attribute("Id") == relationshipId).Remove();
        document.SetXml(relPart, rels);
        document.SetXml(part, xml);
        // Retain unreferenced parts: other slides may link to them. Deletion is not sanitization.
        return document;
    }

    public static PptxObjectDocument UpdateText(PptxObjectDocument source, int number, string objectId, string text)
    {
        if (text.Length > 100000)
        {
            throw new ArgumentException("Text is too long.");
        }
        var document = Editable(source);
        var slides = new PptxPackageGraph(document).Slides();
        if (number < 1 || number > slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(number));
        }
        var part = slides[number - 1].Part;
        var xml = document.Xml(part);
        var shape = xml.Descendants(P + "sp").SingleOrDefault(element =>
            (string?)element.Element(P + "nvSpPr")?.Element(P + "cNvPr")?.Attribute("id") == objectId)
            ?? throw new ArgumentException("Text object not found.");
        var body = shape.Element(P + "txBody") ?? throw new ArgumentException("Object has no editable text.");
        var oldParagraphs = body.Elements(A + "p").ToList();
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var replacement = new List<XElement>();
        for (var i = 0; i < lines.Length; i++)
        {
            var original = oldParagraphs.ElementAtOrDefault(i) ?? oldParagraphs.FirstOrDefault();
            var paragraph = new XElement(A + "p");
            if (original?.Element(A + "pPr") is { } paragraphProperties)
            {
                paragraph.Add(new XElement(paragraphProperties));
            }
            var properties = original?.Element(A + "r")?.Element(A + "rPr");
            var run = new XElement(A + "r");
            if (properties is not null)
            {
                var style = new XElement(properties);
                style.Elements().Where(element => element.Name.LocalName.StartsWith("hlink", StringComparison.Ordinal)).Remove();
                run.Add(style);
            }
            run.Add(new XElement(A + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), lines[i]));
            paragraph.Add(run);
            replacement.Add(paragraph);
        }
        body.Elements(A + "p").Remove();
        body.Add(replacement);
        document.SetXml(part, xml);
        return document;
    }
}
