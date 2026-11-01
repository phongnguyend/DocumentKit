using System.Xml.Linq;
using PptxObjectKit.Graphs;
using PptxObjectKit.Models;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Editors;

public static class PptxImageEditor
{
    public static PptxObjectDocument Edit(PptxObjectDocument source, int slideNumber, string objectId, string action, byte[]? image = null, double x = 36, double y = 36, double width = 180, double height = 120)
    {
        if (source.PartNames.Any(path => path.StartsWith("_xmlsignatures/", StringComparison.Ordinal)))
        {
            throw new NotSupportedException("Editing signed presentations is not supported.");
        }
        var document = source.Clone();
        var slides = new PptxPackageGraph(document).Slides();
        if (slideNumber < 1 || slideNumber > slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(slideNumber));
        }
        var part = slides[slideNumber - 1].Part;
        var xml = document.Xml(part);
        var picture = xml.Descendants(P + "pic").SingleOrDefault(item => (string?)item.Element(P + "nvPicPr")?.Element(P + "cNvPr")?.Attribute("id") == objectId);
        if (action != "add-image" && picture is null)
        {
            throw new ArgumentException("Image object not found.");
        }
        if (action == "delete-image")
        {
            picture!.Remove();
        }
        else if (action is "add-image" or "replace-image" or "resize-image")
        {
            if (action == "add-image")
            {
                var id = xml.Descendants(P + "cNvPr").Select(item => (long?)item.Attribute("id") ?? 0).DefaultIfEmpty().Max() + 1;
                picture = new XElement(P + "pic", new XElement(P + "nvPicPr", new XElement(P + "cNvPr", new XAttribute("id", id), new XAttribute("name", "Image " + id)), new XElement(P + "cNvPicPr"), new XElement(P + "nvPr")), new XElement(P + "blipFill", new XElement(A + "blip"), new XElement(A + "stretch", new XElement(A + "fillRect"))), new XElement(P + "spPr", new XElement(A + "xfrm"), new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))));
                xml.Descendants(P + "spTree").First().Add(picture);
            }
            if (action is "add-image" or "resize-image")
            {
                if (!new[] { x, y, width, height }.All(value => double.IsFinite(value) && Math.Abs(value) <= 100000) || width <= 0 || height <= 0)
                {
                    throw new ArgumentException("Provide finite position and positive image dimensions in points.");
                }
                if (picture!.Ancestors(P + "grpSp").Any())
                {
                    throw new NotSupportedException("Position and size changes for grouped images are not supported yet.");
                }
                var properties = picture.Element(P + "spPr")!;
                var frame = properties.Element(A + "xfrm");
                if (frame is null)
                {
                    frame = new XElement(A + "xfrm");
                    properties.AddFirst(frame);
                }
                frame.Element(A + "off")?.Remove();
                frame.Element(A + "ext")?.Remove();
                frame.Add(new XElement(A + "off", new XAttribute("x", (long)Math.Round(x * 12700)), new XAttribute("y", (long)Math.Round(y * 12700))), new XElement(A + "ext", new XAttribute("cx", (long)Math.Round(width * 12700)), new XAttribute("cy", (long)Math.Round(height * 12700))));
            }
            if (action is "add-image" or "replace-image")
            {
                if (image is null || image.Length > 5 * 1024 * 1024)
                {
                    throw new ArgumentException("Choose a PNG or JPEG image up to 5 MB.");
                }
                var extension = image.AsSpan().StartsWith(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }) ? "png" : image.AsSpan().StartsWith(new byte[] { 255, 216, 255 }) ? "jpg" : throw new ArgumentException("Only PNG or JPEG image bytes are supported.");
                var path = "ppt/media/image-" + Guid.NewGuid().ToString("N") + "." + extension;
                document.Parts[path] = image.ToArray();
                var relationshipPart = RelationshipPart(part);
                var relationships = document.Parts.ContainsKey(relationshipPart) ? document.Xml(relationshipPart) : new XDocument(new XElement(Rel + "Relationships"));
                var relationshipId = "image" + Guid.NewGuid().ToString("N");
                relationships.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", relationshipId), new XAttribute("Type", R.NamespaceName + "/image"), new XAttribute("Target", "/" + path)));
                document.SetXml(relationshipPart, relationships);
                var fill = picture!.Element(P + "blipFill")!;
                var blip = fill.Element(A + "blip")!;
                blip.SetAttributeValue(R + "embed", relationshipId);
                blip.Attribute(R + "link")?.Remove();
                // Alternate SVG/effect payloads can otherwise keep displaying the old image.
                blip.RemoveNodes();
                var types = document.Xml("[Content_Types].xml");
                types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/" + path), new XAttribute("ContentType", extension == "png" ? "image/png" : "image/jpeg")));
                document.SetXml("[Content_Types].xml", types);
            }
        }
        else
        {
            throw new ArgumentException("Unknown image action.");
        }
        document.SetXml(part, xml);
        return document;
    }
}
