using DocxObjectKit.Editors;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extractors;

public sealed record DocxParagraph(string Id, string Text, string? Style, bool InTable);
public sealed record DocxPicture(string Id, string Name, string? PartName, double Width, double Height);
public sealed record DocxDocumentContent(IReadOnlyList<DocxParagraph> Paragraphs, IReadOnlyList<DocxPicture> Images);

/// <summary>Extracts semantic objects without requiring a native rendering dependency.</summary>
public static class DocxDocumentExtractor
{
    public static DocxDocumentContent Extract(DocxObjectDocument document)
    {
        var xml = document.ReadXml("word/document.xml");
        var paragraphs = xml.Descendants(W + "p").Select((paragraph, index) => new DocxParagraph("p" + index,
            string.Concat(paragraph.Descendants().Where(item => !item.Ancestors(W + "del").Any()).Select(item => item.Name == W + "t" ? item.Value : item.Name == W + "br" ? "\n" : item.Name == W + "tab" ? "\t" : "")),
            Value(paragraph.Element(W + "pPr")?.Element(W + "pStyle")), paragraph.Ancestors(W + "tc").Any())).ToArray();
        var images = DocxEditor.Pictures(xml).Select((picture, index) =>
        {
            string? part;
            try
            {
                part = DocxEditor.ImagePart(document, picture);
            }
            catch (NotSupportedException)
            {
                part = null;
            }
            return new DocxPicture("i" + index, (string?)picture.Element(WP + "docPr")?.Attribute("name") ?? "Image", part,
                Number(picture.Element(WP + "extent"), "cx") / 12700, Number(picture.Element(WP + "extent"), "cy") / 12700);
        }).ToArray();
        return new(paragraphs, images);
    }
}
