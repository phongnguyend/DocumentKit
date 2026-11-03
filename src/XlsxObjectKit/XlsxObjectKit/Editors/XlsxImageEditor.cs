using System.Xml.Linq;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;
using static XlsxObjectKit.Extractors.XlsxImageExtractor;

namespace XlsxObjectKit.Editors;

public static class XlsxImageEditor
{
    public static XlsxObjectDocument Edit(XlsxObjectDocument document, string sheetId, string action, string? imageId = null,
        byte[]? bytes = null, double x = 0, double y = 0, double width = 240, double height = 160)
    {
        if (action is not ("add-image" or "replace-image" or "transform-image" or "delete-image"))
        {
            throw new ArgumentException("Unknown image action.");
        }
        var sheet = XlsxWorkbookExtractor.Extract(document).Sheets.Single(item => item.Id == sheetId);
        XlsxEditor.Guard(document, sheet.Part);
        if (action == "transform-image" && sheet.Images.Any(image => image.Id == imageId && !image.Editable))
        {
            throw new NotSupportedException("Moving cropped or rotated pictures is not supported yet.");
        }
        if (action is "add-image" or "transform-image" && (!new[] { x, y, width, height }.All(value => double.IsFinite(value) && value >= 0 && value <= 10000000) || width < 1 || height < 1))
        {
            throw new ArgumentException("Provide a nonnegative position and positive image size in pixels.");
        }
        string? mime = null;
        if (action is "add-image" or "replace-image")
        {
            mime = bytes is { Length: > 0 and <= 5242880 } ? Mime(bytes) : null;
            if (mime is null)
            {
                throw new ArgumentException("Choose PNG or JPEG image bytes, up to 5 MB.");
            }
        }
        var xml = document.ReadXml(sheet.Part);
        var sheetRelations = Relations(document, sheet.Part);
        var reference = xml.Root!.Element(S + "drawing");
        var oldPart = Target(document, sheet.Part, (string?)reference?.Attribute(R + "id"));
        var drawing = oldPart is null ? new XDocument(new XElement(D + "wsDr")) : document.ReadXml(oldPart);
        var relations = oldPart is null ? new XDocument(new XElement(Rel + "Relationships")) : Relations(document, oldPart);
        // Isolate edits even when two worksheets share a drawing part.
        if (oldPart is not null)
        {
            foreach (var item in relations.Root!.Elements().Where(item => (string?)item.Attribute("TargetMode") != "External"))
            {
                item.SetAttributeValue("Target", "/" + Resolve(oldPart, (string)item.Attribute("Target")!));
            }
        }
        var picture = drawing.Root!.Elements().Elements(D + "pic").SingleOrDefault(item => (string?)item.Element(D + "nvPicPr")?.Element(D + "cNvPr")?.Attribute("id") == imageId);
        if (action != "add-image" && picture is null)
        {
            throw new ArgumentException("Image not found on the selected worksheet.");
        }
        if (action == "delete-image")
        {
            picture!.Parent!.Remove();
        }
        else
        {
            if (action == "add-image")
            {
                var id = drawing.Descendants(D + "cNvPr").Select(item => (long?)item.Attribute("id") ?? 0).DefaultIfEmpty().Max() + 1;
                picture = new XElement(D + "pic", new XElement(D + "nvPicPr", new XElement(D + "cNvPr", new XAttribute("id", id), new XAttribute("name", "Image " + id)), new XElement(D + "cNvPicPr")),
                    new XElement(D + "blipFill", new XElement(A + "blip"), new XElement(A + "stretch", new XElement(A + "fillRect"))),
                    new XElement(D + "spPr", new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))));
            }
            if (action is "add-image" or "replace-image")
            {
                var media = "xl/media/image-" + Guid.NewGuid().ToString("N") + (mime == "image/png" ? ".png" : ".jpg");
                document = document.WithPart(media, bytes!);
                var rid = "image" + Guid.NewGuid().ToString("N");
                relations.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", rid), new XAttribute("Type", R.NamespaceName + "/image"), new XAttribute("Target", "/" + media)));
                picture!.Descendants(A + "blip").Single().SetAttributeValue(R + "embed", rid);
                var types = document.ReadXml("[Content_Types].xml");
                var ct = types.Root!.Name.Namespace;
                types.Root.Add(new XElement(ct + "Override", new XAttribute("PartName", "/" + media), new XAttribute("ContentType", mime!)));
                document = document.WithXml("[Content_Types].xml", types);
            }
            if (action is "add-image" or "transform-image")
            {
                var oldAnchor = picture!.Parent;
                if (oldAnchor is not null)
                {
                    picture.Remove();
                }
                var properties = picture.Element(D + "spPr")!;
                var transform = new XElement(A + "xfrm",
                    new XElement(A + "off", new XAttribute("x", (long)Math.Round(x * 9525)), new XAttribute("y", (long)Math.Round(y * 9525))),
                    new XElement(A + "ext", new XAttribute("cx", (long)Math.Round(width * 9525)), new XAttribute("cy", (long)Math.Round(height * 9525))));
                if (properties.Element(A + "xfrm") is { } previousTransform)
                {
                    previousTransform.ReplaceWith(transform);
                }
                else
                {
                    properties.AddFirst(transform);
                }
                var anchor = new XElement(D + "absoluteAnchor", new XElement(D + "pos", new XAttribute("x", (long)Math.Round(x * 9525)), new XAttribute("y", (long)Math.Round(y * 9525))),
                    new XElement(D + "ext", new XAttribute("cx", (long)Math.Round(width * 9525)), new XAttribute("cy", (long)Math.Round(height * 9525))), picture, new XElement(D + "clientData"));
                if (oldAnchor is null)
                {
                    drawing.Root.Add(anchor);
                }
                else
                {
                    oldAnchor.ReplaceWith(anchor);
                }
            }
        }
        var part = "xl/drawings/drawing-" + Guid.NewGuid().ToString("N") + ".xml";
        var drawingId = "drawing" + Guid.NewGuid().ToString("N");
        sheetRelations.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", drawingId), new XAttribute("Type", R.NamespaceName + "/drawing"), new XAttribute("Target", "/" + part)));
        if (reference is null)
        {
            reference = new XElement(S + "drawing");
            var following = xml.Root.Elements().FirstOrDefault(item => new[] { "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst" }.Contains(item.Name.LocalName));
            if (following is null)
            {
                xml.Root.Add(reference);
            }
            else
            {
                following.AddBeforeSelf(reference);
            }
        }
        reference.SetAttributeValue(R + "id", drawingId);
        var contentTypes = document.ReadXml("[Content_Types].xml");
        contentTypes.Root!.Add(new XElement(contentTypes.Root.Name.Namespace + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")));
        return document.WithXml(sheet.Part, xml).WithXml(RelationsPart(sheet.Part), sheetRelations).WithXml(part, drawing)
            .WithXml(RelationsPart(part), relations).WithXml("[Content_Types].xml", contentTypes);
    }
}

