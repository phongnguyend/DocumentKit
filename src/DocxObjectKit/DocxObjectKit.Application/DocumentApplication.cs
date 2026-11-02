using System.Xml.Linq;
using DocxObjectKit.Editors;
using DocxObjectKit.Extensions.Pdf;
using DocxObjectKit.Readers;
using DocxObjectKit.Writers;

namespace DocxObjectKit.Application;

public sealed record DocxEditRequest(string Kind, int PageNumber = 1, string ObjectId = "", string Text = "", string? ImageBase64 = null, double X = 36, double Y = 36, double Width = 180, double Height = 120, bool UpdateBounds = false, DocxOcrLine[]? OcrLines = null, bool ReplaceOcrImage = false);

public sealed record DocxOcrLine(string Text, double X, double Y, double Width, double Height, double Baseline);

public static class DocumentApplication
{
    public static DocxLayoutResult Inspect(byte[] bytes, CancellationToken cancellationToken = default)
    {
        using var engine = new DocxLayout(DocxReader.Read(bytes));
        return engine.Build(cancellationToken);
    }

    private static DocxObjectKit.Models.DocxObjectDocument ApplyOcr(DocxObjectKit.Models.DocxObjectDocument document, DocxEditRequest request)
    {
        var lines = request.OcrLines ?? [];
        if (lines.Length == 0 || lines.Length > 2000 || lines.Any(line => string.IsNullOrWhiteSpace(line.Text) || line.Text.Length > 10000 || !new[] { line.X, line.Y, line.Width, line.Height, line.Baseline }.All(double.IsFinite) || line.X < 0 || line.Y < 0 || line.Width <= 0 || line.Height <= 0 || line.X + line.Width > 1.00001 || line.Y + line.Height > 1.00001 || line.Baseline < line.Y || line.Baseline > line.Y + line.Height + .00001))
        {
            throw new ArgumentException("OCR requires text with finite normalized bounds inside the image.");
        }
        using var engine = new DocxLayout(document);
        var image = engine.Build().Pages.SelectMany(page => page.Images).FirstOrDefault(item => item.Id == request.ObjectId) ?? throw new ArgumentException("Image not found.");
        var xml = document.ReadXml("word/document.xml");
        var picture = DocxEditor.Picture(xml, request.ObjectId);
        var anchor = picture.Ancestors(DocxObjectKit.Models.DocxNamespaces.W + "p").First();
        var w = DocxObjectKit.Models.DocxNamespaces.W;
        foreach (var line in lines)
        {
            var size = Math.Max(2, image.Height * line.Height * .8);
            var paragraph = new XElement(w + "p",
                new XElement(w + "pPr", new XElement(w + "framePr",
                    new XAttribute(w + "hAnchor", "page"), new XAttribute(w + "vAnchor", "page"), new XAttribute(w + "wrap", "none"),
                    new XAttribute(w + "x", (long)((image.X + line.X * image.Width) * 20)), new XAttribute(w + "y", (long)((image.Y + line.Y * image.Height) * 20)),
                    new XAttribute(w + "w", (long)(line.Width * image.Width * 20)), new XAttribute(w + "h", (long)(line.Height * image.Height * 20)))),
                new XElement(w + "r", new XElement(w + "rPr", new XElement(w + "sz", new XAttribute(w + "val", (int)(size * 2)))), new XElement(w + "t", line.Text)));
            anchor.AddAfterSelf(paragraph);
            anchor = paragraph;
        }
        if (request.ReplaceOcrImage)
        {
            var originalParagraph = picture.Ancestors(w + "p").First();
            if (picture.Name.LocalName == "inline")
            {
                var properties = originalParagraph.Element(w + "pPr");
                if (properties is null)
                {
                    properties = new XElement(w + "pPr");
                    originalParagraph.AddFirst(properties);
                }
                properties.Elements(w + "spacing").Remove();
                properties.Add(new XElement(w + "spacing", new XAttribute(w + "after", (long)(image.Height * 20))));
            }
            picture.Ancestors(w + "drawing").First().Remove();
        }
        return document.WithXml("word/document.xml", xml);
    }

    public static byte[] Edit(byte[] bytes, DocxEditRequest request)
    {
        var document = DocxReader.Read(bytes);
        string? anchor = null;
        if (request.Kind is "add-image" or "add-text")
        {
            using var engine = new DocxLayout(document);
            var pages = engine.Build().Pages;
            if (request.PageNumber < 1 || request.PageNumber > pages.Count)
            {
                throw new ArgumentException("Invalid page number.");
            }
            var page = pages[request.PageNumber - 1];
            anchor = page.TextObjects.FirstOrDefault(item => !item.Bounds.Grouped)?.Id;
            if (anchor is null && page.Images.FirstOrDefault() is { } firstImage)
            {
                var xml = document.ReadXml("word/document.xml");
                var paragraph = DocxEditor.Picture(xml, firstImage.Id).Ancestors(DocxObjectKit.Models.DocxNamespaces.W + "p").First();
                anchor = "p" + xml.Descendants(DocxObjectKit.Models.DocxNamespaces.W + "p").TakeWhile(item => item != paragraph).Count();
            }
            if (anchor is null && request.PageNumber > 1)
            {
                throw new NotSupportedException("This page has no independent paragraph anchor. Add content on a page with a complete paragraph or image.");
            }
        }
        byte[]? image = null;
        if (request.ImageBase64 is not null)
        {
            if (request.ImageBase64.Length > 7 * 1024 * 1024)
            {
                throw new ArgumentException("Image exceeds the upload limit.");
            }
            image = Convert.FromBase64String(request.ImageBase64);
        }
        var edited = request.Kind switch
        {
            "apply-image-ocr" => ApplyOcr(document, request),
            "add-text" => DocxEditor.AddText(document, request.Text, anchor),
            "text" => DocxEditor.SetText(document, request.ObjectId, request.Text),
            "delete-text" => DocxEditor.DeleteText(document, request.ObjectId),
            "transform-text" => DocxEditor.TransformText(document, request.ObjectId, request.X, request.Y, request.Width, request.Height),
            "add-image" or "replace-image" or "delete-image" or "resize-image" => DocxEditor.Image(document, request.Kind, request.ObjectId, image, request.X, request.Y, request.Width, request.Height, anchor, request.UpdateBounds),
            _ => throw new ArgumentException("Unknown edit operation.")
        };
        return DocxWriter.Write(edited);
    }
}
