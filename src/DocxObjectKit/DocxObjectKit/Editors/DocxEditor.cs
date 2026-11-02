using System.Xml.Linq;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Editors;

public static class DocxEditor
{
    public static XElement Paragraph(XDocument xml, string id) => Find(xml.Descendants(W + "p"), id, "p");

    public static XElement Picture(XDocument xml, string id) => Find(Pictures(xml), id, "i");

    public static IEnumerable<XElement> Pictures(XDocument xml) => xml.Descendants().Where(item => item.Name == WP + "inline" || item.Name == WP + "anchor");

    private static XElement Find(IEnumerable<XElement> elements, string id, string prefix)
    {
        if (!id.StartsWith(prefix, StringComparison.Ordinal) || !int.TryParse(id.AsSpan(prefix.Length), out var index) || index < 0)
        {
            throw new ArgumentException("Invalid object identifier.");
        }
        return elements.ElementAtOrDefault(index) ?? throw new ArgumentException("Object no longer exists.");
    }

    public static DocxObjectDocument SetText(DocxObjectDocument document, string id, string text)
    {
        if (text.Length > 100000)
        {
            throw new ArgumentException("Text is too long.");
        }
        var xml = document.ReadXml("word/document.xml");
        var paragraph = Paragraph(xml, id);
        if (paragraph.Descendants(W + "drawing").Any())
        {
            throw new NotSupportedException("Edit an image-containing paragraph through its image controls; replacing its text would remove drawings.");
        }
        var style = paragraph.Descendants(W + "rPr").FirstOrDefault();
        paragraph.Elements().Where(item => item.Name != W + "pPr").Remove();
        var run = new XElement(W + "r", style is null ? null : new XElement(style));
        var lines = text.Replace("\r", "").Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
            {
                run.Add(new XElement(W + "br"));
            }
            run.Add(new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), lines[i]));
        }
        paragraph.Add(run);
        return document.WithXml("word/document.xml", xml);
    }

    public static DocxObjectDocument AddText(DocxObjectDocument document, string text, string? afterParagraph = null)
    {
        if (text.Length > 100000)
        {
            throw new ArgumentException("Text is too long.");
        }
        var xml = document.ReadXml("word/document.xml");
        var paragraph = new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));
        if (afterParagraph is not null)
        {
            var anchor = Paragraph(xml, afterParagraph);
            (anchor.Ancestors(W + "tbl").LastOrDefault() ?? anchor).AddAfterSelf(paragraph);
        }
        else
        {
            xml.Root!.Element(W + "body")!.AddFirst(paragraph);
        }
        return document.WithXml("word/document.xml", xml);
    }

    public static DocxObjectDocument DeleteText(DocxObjectDocument document, string id)
    {
        var xml = document.ReadXml("word/document.xml");
        var paragraph = Paragraph(xml, id);
        if (paragraph.Parent?.Name == W + "tc" && paragraph.Parent.Elements(W + "p").Count() == 1)
        {
            return SetText(document, id, "");
        }
        paragraph.Remove();
        return document.WithXml("word/document.xml", xml);
    }

    public static void ValidateBounds(double x, double y, double width, double height)
    {
        if (!new[] { x, y, width, height }.All(double.IsFinite) || x < 0 || y < 0 || width <= 0 || height <= 0 || x + width > 20000 || y + height > 20000)
        {
            throw new ArgumentException("Bounds must be finite, positive and within 20000 points.");
        }
    }

    public static DocxObjectDocument TransformText(DocxObjectDocument document, string id, double x, double y, double width, double height)
    {
        ValidateBounds(x, y, width, height);
        var xml = document.ReadXml("word/document.xml");
        var paragraph = Paragraph(xml, id);
        if (paragraph.Ancestors(W + "tc").Any())
        {
            throw new NotSupportedException("Text inside tables flows with its cell and cannot be positioned independently.");
        }
        var properties = paragraph.Element(W + "pPr");
        if (properties is null)
        {
            properties = new XElement(W + "pPr");
            paragraph.AddFirst(properties);
        }
        properties.Elements(W + "framePr").Remove();
        properties.Add(new XElement(W + "framePr", new XAttribute(W + "hAnchor", "page"), new XAttribute(W + "vAnchor", "page"), new XAttribute(W + "wrap", "none"), new XAttribute(W + "hRule", "atLeast"), new XAttribute(W + "x", (long)(x * 20)), new XAttribute(W + "y", (long)(y * 20)), new XAttribute(W + "w", (long)(width * 20)), new XAttribute(W + "h", (long)(height * 20))));
        return document.WithXml("word/document.xml", xml);
    }

    public static string ImagePart(DocxObjectDocument document, XElement picture)
    {
        var id = (string?)picture.Descendants(A + "blip").FirstOrDefault()?.Attribute(R + "embed");
        var relation = document.ReadXml("word/_rels/document.xml.rels").Root!.Elements().FirstOrDefault(item => (string?)item.Attribute("Id") == id);
        if (relation is null || (string?)relation.Attribute("TargetMode") == "External")
        {
            throw new NotSupportedException("External or missing pictures are not loaded.");
        }
        return Resolve("word/document.xml", (string?)relation.Attribute("Target") ?? "");
    }

    public static DocxObjectDocument Image(DocxObjectDocument document, string kind, string id, byte[]? bytes, double x, double y, double width, double height, string? afterParagraph = null, bool updateBounds = false)
    {
        var xml = document.ReadXml("word/document.xml");
        XElement picture;
        if (kind == "add-image")
        {
            picture = new XElement(WP + "inline", new XElement(WP + "extent", new XAttribute("cx", 1), new XAttribute("cy", 1)), new XElement(WP + "docPr", new XAttribute("id", xml.Descendants(WP + "docPr").Select(item => Number(item, "id")).DefaultIfEmpty().Max() + 1), new XAttribute("name", "Added image")), new XElement(A + "graphic", new XElement(A + "graphicData", new XAttribute("uri", PIC.NamespaceName), new XElement(PIC + "pic", new XElement(PIC + "nvPicPr", new XElement(PIC + "cNvPr", new XAttribute("id", 0), new XAttribute("name", "Image")), new XElement(PIC + "cNvPicPr")), new XElement(PIC + "blipFill", new XElement(A + "blip"), new XElement(A + "stretch", new XElement(A + "fillRect"))), new XElement(PIC + "spPr", new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(A + "ext", new XAttribute("cx", 1), new XAttribute("cy", 1))), new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst")))))));
            var paragraph = new XElement(W + "p", new XElement(W + "r", new XElement(W + "drawing", picture)));
            if (afterParagraph is not null)
            {
                var anchor = Paragraph(xml, afterParagraph);
                var block = anchor.Ancestors(W + "tbl").LastOrDefault() ?? anchor;
                block.AddAfterSelf(paragraph);
            }
            else
            {
                xml.Root!.Element(W + "body")!.AddFirst(paragraph);
            }
        }
        else
        {
            picture = Picture(xml, id);
        }
        if (kind == "delete-image")
        {
            picture.Ancestors(W + "drawing").First().Remove();
            return document.WithXml("word/document.xml", xml);
        }
        if (kind is "add-image" or "replace-image")
        {
            if (bytes is null || bytes.Length > 5 * 1024 * 1024)
            {
                throw new ArgumentException("Choose an image no larger than 5 MB.");
            }
            var extension = bytes.AsSpan().StartsWith(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }) ? "png" : bytes.AsSpan().StartsWith(new byte[] { 255, 216, 255 }) ? "jpg" : throw new ArgumentException("Images must be PNG or JPEG.");
            var name = "image-" + Guid.NewGuid().ToString("N") + "." + extension;
            var relationId = "rId" + Guid.NewGuid().ToString("N");
            var rels = document.PartNames.Contains("word/_rels/document.xml.rels") ? document.ReadXml("word/_rels/document.xml.rels") : new XDocument(new XElement(REL + "Relationships"));
            rels.Root!.Add(new XElement(REL + "Relationship", new XAttribute("Id", relationId), new XAttribute("Type", R.NamespaceName + "/image"), new XAttribute("Target", "media/" + name)));
            picture.Descendants(A + "blip").First().SetAttributeValue(R + "embed", relationId);
            picture.Descendants(A + "blip").First().Attribute(R + "link")?.Remove();
            var types = document.ReadXml("[Content_Types].xml");
            if (!types.Root!.Elements().Any(item => (string?)item.Attribute("Extension") == extension))
            {
                types.Root.Add(new XElement(types.Root.Name.Namespace + "Default", new XAttribute("Extension", extension), new XAttribute("ContentType", extension == "png" ? "image/png" : "image/jpeg")));
            }
            document = document.WithXml("word/_rels/document.xml.rels", rels).WithXml("[Content_Types].xml", types).WithPart("word/media/" + name, bytes);
        }
        if (kind is "add-image" or "resize-image" || updateBounds)
        {
            ValidateBounds(x, y, width, height);
            picture.Name = WP + "anchor";
            foreach (var attribute in new[] { "simplePos", "relativeHeight", "behindDoc", "locked", "layoutInCell", "allowOverlap", "distT", "distB", "distL", "distR" })
            {
                picture.SetAttributeValue(attribute, attribute is "layoutInCell" or "allowOverlap" ? "1" : "0");
            }
            picture.Elements().Where(item => item.Name.Namespace == WP && (item.Name.LocalName.StartsWith("wrap", StringComparison.Ordinal) || item.Name.LocalName is "simplePos" or "positionH" or "positionV")).Remove();
            picture.AddFirst(new XElement(WP + "simplePos", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(WP + "positionH", new XAttribute("relativeFrom", "page"), new XElement(WP + "posOffset", (long)(x * 12700))), new XElement(WP + "positionV", new XAttribute("relativeFrom", "page"), new XElement(WP + "posOffset", (long)(y * 12700))));
            var extent = picture.Element(WP + "extent")!;
            extent.SetAttributeValue("cx", (long)(width * 12700));
            extent.SetAttributeValue("cy", (long)(height * 12700));
            (picture.Element(WP + "effectExtent") ?? extent).AddAfterSelf(new XElement(WP + "wrapNone"));
            foreach (var size in picture.Descendants(A + "ext"))
            {
                size.SetAttributeValue("cx", (long)(width * 12700));
                size.SetAttributeValue("cy", (long)(height * 12700));
            }
        }
        return document.WithXml("word/document.xml", xml);
    }
}
