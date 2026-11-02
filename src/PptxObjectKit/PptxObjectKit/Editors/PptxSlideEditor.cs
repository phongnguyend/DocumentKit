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

    /// <summary>Copies a slide and its owned parts immediately after the source.</summary>
    public static PptxObjectDocument Duplicate(PptxObjectDocument source, int number) => Insert(source, number, false);

    /// <summary>Inserts an empty white slide after the specified slide.</summary>
    public static PptxObjectDocument InsertBlank(PptxObjectDocument source, int afterNumber) => Insert(source, afterNumber, true);

    private static PptxObjectDocument Insert(PptxObjectDocument source, int afterNumber, bool blank)
    {
        var document = Editable(source);
        var graph = new PptxPackageGraph(document);
        var slides = graph.Slides();
        if (afterNumber < 1 || afterNumber > slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(afterNumber));
        }
        if (slides.Count >= 2000)
        {
            throw new NotSupportedException("Presentation exceeds 2,000 slides.");
        }
        var sourcePart = slides[afterNumber - 1].Part;
        var types = document.Xml("[Content_Types].xml");
        var typeNamespace = types.Root!.Name.Namespace;
        var copies = new Dictionary<string, string>(StringComparer.Ordinal);
        string CopyPart(string path, int depth = 0)
        {
            if (copies.TryGetValue(path, out var existing))
            {
                return existing;
            }
            if (depth > 64 || copies.Count >= 2000)
            {
                throw new NotSupportedException("Slide dependency graph is too complex to duplicate.");
            }
            var slash = path.LastIndexOf('/');
            var extension = Path.GetExtension(path);
            var target = path[..(slash + 1)] + "copy-" + Guid.NewGuid().ToString("N") + extension;
            copies.Add(path, target);
            document.Parts[target] = document.GetPart(path);
            var contentType = types.Root.Elements(typeNamespace + "Override").FirstOrDefault(item => (string?)item.Attribute("PartName") == "/" + path);
            if (contentType is not null)
            {
                var copy = new XElement(contentType);
                copy.SetAttributeValue("PartName", "/" + target);
                types.Root.Add(copy);
            }
            var relationships = new XElement(Rel + "Relationships");
            foreach (var relation in graph.Relationships(path))
            {
                var copy = new XElement(relation);
                if ((string?)relation.Attribute("TargetMode") != "External")
                {
                    var originalTarget = Target(path, relation);
                    var kind = ((string?)relation.Attribute("Type") ?? "").Split('/').Last();
                    var shared = kind is "slideLayout" or "slideMaster" or "notesMaster" or "theme" or "image" or "audio" or "video" or "slide";
                    var copiedTarget = copies.TryGetValue(originalTarget, out var mapped) ? mapped : shared ? originalTarget : CopyPart(originalTarget, depth + 1);
                    copy.SetAttributeValue("Target", "/" + copiedTarget);
                }
                relationships.Add(copy);
            }
            if (relationships.HasElements)
            {
                document.SetXml(RelationshipPart(target), new XDocument(relationships));
            }
            return target;
        }
        string newPart;
        if (blank)
        {
            newPart = sourcePart[..(sourcePart.LastIndexOf('/') + 1)] + "blank-" + Guid.NewGuid().ToString("N") + ".xml";
            var shapeTree = new XElement(P + "spTree",
                new XElement(P + "nvGrpSpPr", new XElement(P + "cNvPr", new XAttribute("id", 1), new XAttribute("name", "")), new XElement(P + "cNvGrpSpPr"), new XElement(P + "nvPr")),
                new XElement(P + "grpSpPr"));
            document.SetXml(newPart, new XDocument(new XElement(P + "sld", new XAttribute("showMasterSp", 0),
                new XElement(P + "cSld", new XElement(P + "bg", new XElement(P + "bgPr", new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", "FFFFFF"))), new XElement(A + "effectLst"))), shapeTree),
                new XElement(P + "clrMapOvr", new XElement(A + "masterClrMapping")))));
            var layout = graph.Relationships(sourcePart).FirstOrDefault(item => ((string?)item.Attribute("Type"))?.EndsWith("/slideLayout", StringComparison.Ordinal) == true);
            if (layout is not null)
            {
                var copy = new XElement(layout);
                copy.SetAttributeValue("Target", "/" + Target(sourcePart, layout));
                document.SetXml(RelationshipPart(newPart), new XDocument(new XElement(Rel + "Relationships", copy)));
            }
            types.Root.Add(new XElement(typeNamespace + "Override", new XAttribute("PartName", "/" + newPart), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")));
        }
        else
        {
            newPart = CopyPart(sourcePart);
        }
        var presentationPart = graph.PresentationPart();
        var presentation = document.Xml(presentationPart);
        var list = presentation.Root!.Element(P + "sldIdLst")!;
        var ids = list.Elements(P + "sldId").ToArray();
        var nextId = Math.Max(256, ids.Max(item => (long?)item.Attribute("id") ?? 255) + 1);
        if (nextId > int.MaxValue)
        {
            nextId = Enumerable.Range(256, ids.Length + 1).First(value => !ids.Any(item => (long?)item.Attribute("id") == value));
        }
        var relationshipId = "slide" + Guid.NewGuid().ToString("N");
        ids[afterNumber - 1].AddAfterSelf(new XElement(P + "sldId", new XAttribute("id", nextId), new XAttribute(R + "id", relationshipId)));
        var presentationRels = document.Xml(RelationshipPart(presentationPart));
        presentationRels.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", relationshipId), new XAttribute("Type", R.NamespaceName + "/slide"), new XAttribute("Target", "/" + newPart)));
        document.SetXml(RelationshipPart(presentationPart), presentationRels);
        document.SetXml(presentationPart, presentation);
        document.SetXml("[Content_Types].xml", types);
        return document;
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

    /// <summary>Updates a text box frame without changing its text, formatting, or wrapping settings.</summary>
    public static PptxObjectDocument TransformText(PptxObjectDocument source, int number, string objectId, double x, double y, double width, double height)
    {
        if (!new[] { x, y, width, height }.All(value => double.IsFinite(value) && Math.Abs(value) <= 100000) || width <= 0 || height <= 0)
        {
            throw new ArgumentException("Provide finite position and positive text box dimensions in points.");
        }
        var document = Editable(source);
        var slides = new PptxPackageGraph(document).Slides();
        if (number < 1 || number > slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(number));
        }
        var part = slides[number - 1].Part;
        var xml = document.Xml(part);
        var shape = xml.Descendants(P + "sp").SingleOrDefault(item => (string?)item.Element(P + "nvSpPr")?.Element(P + "cNvPr")?.Attribute("id") == objectId && item.Element(P + "txBody") is not null)
            ?? throw new ArgumentException("Text object not found.");
        if (shape.Ancestors(P + "grpSp").Any())
        {
            throw new NotSupportedException("Moving or resizing grouped text boxes is not supported yet.");
        }
        var originalFrame = new PptxTextGeometry(document, part).Frame(shape) ?? throw new NotSupportedException("Text box placement could not be resolved.");
        var frame = new XElement(originalFrame);
        frame.Element(A + "off")?.Remove();
        frame.Element(A + "ext")?.Remove();
        frame.Add(new XElement(A + "off", new XAttribute("x", (long)Math.Round(x * 12700)), new XAttribute("y", (long)Math.Round(y * 12700))),
            new XElement(A + "ext", new XAttribute("cx", Math.Max(1, (long)Math.Round(width * 12700))), new XAttribute("cy", Math.Max(1, (long)Math.Round(height * 12700)))));
        var properties = shape.Element(P + "spPr");
        if (properties is null)
        {
            properties = new XElement(P + "spPr");
            shape.Element(P + "nvSpPr")!.AddAfterSelf(properties);
        }
        properties.Element(A + "xfrm")?.Remove();
        properties.AddFirst(frame);
        document.SetXml(part, xml);
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
