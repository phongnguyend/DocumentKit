using System.IO.Compression;
using System.Text;
using PptxObjectKit.Application;
using PptxObjectKit.Converters;
using PptxObjectKit.Editors;
using PptxObjectKit.Extractors;
using PptxObjectKit.Readers;
using PptxObjectKit.Writers;

var checks = 0;
void Check(bool condition, string message)
{
    if (!condition)
    {
        throw new Exception(message);
    }
    checks++;
}
void Reject(Action action)
{
    try
    {
        action();
    }
    catch (Exception exception) when (exception is InvalidDataException or ArgumentException or NotSupportedException)
    {
        checks++;
        return;
    }
    throw new Exception("Expected rejection.");
}
var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures/sample.pptx"));
var document = PptxReader.Read(bytes);
var inspection = PptxSlideExtractor.Extract(document);
Check(inspection.Slides.Count == 2, "Slide count");
Check(inspection.Slides[1].PartName.EndsWith("unusual-name.xml"), "Use relationships, not numeric filenames");
var reordered = PptxSlideEditor.Reorder(document, [2, 1]);
Check(PptxSlideExtractor.Extract(PptxReader.Read(PptxWriter.Write(reordered))).Slides[0].Title == "Second slide", "Reorder round trip");
Check(PptxSlideExtractor.Extract(document).Slides[0].Title != "Second slide", "Source immutability");
var editedBytes = PresentationApplication.Edit(bytes, new("text", 1, "2", "Changed title\nSecond paragraph"));
var edited = PptxReader.Read(editedBytes);
Check(PptxSlideExtractor.Extract(edited).Slides[0].TextObjects[0].Text == "Changed title\nSecond paragraph", "Multiline text edit");
foreach (var part in document.PartNames.Where(part => part != "ppt/slides/slide1.xml"))
{
    Check(document.GetPart(part).SequenceEqual(edited.GetPart(part)), "Preserve untouched package parts: " + part);
}
var deleted = PptxReader.Read(PptxWriter.Write(PptxSlideEditor.Delete(document, 1)));
Check(PptxSlideExtractor.Extract(deleted).Slides.Single().Title == "Second slide", "Delete slide");
Reject(() => PptxSlideEditor.Delete(deleted, 1));
Reject(() => PptxSlideEditor.Reorder(document, [1, 1]));
Reject(() => PptxSlideEditor.UpdateText(document, 1, "missing", "Text"));
var markdown = PptxMarkdownConverter.Convert(edited);
Check(markdown.Contains("Changed title") && markdown.Contains("Second paragraph") && markdown.Contains("# Slide 2:"), "Markdown paragraphs and order");
var escaped = PptxSlideEditor.UpdateText(document, 1, "2", "<script>alert(1)</script>");
Check(!PptxMarkdownConverter.Convert(escaped).Contains("<script>"), "Escape HTML in Markdown");
Reject(() => PptxReader.Read([1, 2, 3]));
byte[] AddPart(string name)
{
    using var memory = new MemoryStream();
    memory.Write(bytes);
    using (var archive = new ZipArchive(memory, ZipArchiveMode.Update, true))
    {
        using var output = archive.CreateEntry(name).Open();
        output.Write(Encoding.UTF8.GetBytes("test"));
    }
    return memory.ToArray();
}
Reject(() => PptxReader.Read(AddPart("../unsafe.xml")));
Reject(() => PptxSlideEditor.Delete(PptxReader.Read(AddPart("_xmlsignatures/sig1.xml")), 1));
// Exercise Markdown content beyond plain text without relying on private documents.
using var richPackage = new MemoryStream();
richPackage.Write(bytes);
using (var archive = new ZipArchive(richPackage, ZipArchiveMode.Update, true))
{
    void Put(string name, string xml)
    {
        archive.GetEntry(name)?.Delete();
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(xml);
    }
    var slideXml = Encoding.UTF8.GetString(document.GetPart("ppt/slides/slide1.xml"));
    slideXml = slideXml.Replace("</p:spTree>", """
        <p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tr>
        <a:tc><a:txBody><a:p><a:r><a:t>Cell | one</a:t></a:r></a:p></a:txBody></a:tc>
        <a:tc><a:txBody><a:p><a:r><a:t>Cell two</a:t></a:r></a:p></a:txBody></a:tc>
        </a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
        <p:pic><p:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="picture"/></p:blipFill></p:pic></p:spTree>
        """);
    Put("ppt/slides/slide1.xml", slideXml);
    Put("ppt/slides/_rels/slide1.xml.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="notes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
        <Relationship Id="picture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image.png"/>
        </Relationships>
        """);
    Put("ppt/notesSlides/notesSlide1.xml", """
        <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
        <p:txBody><a:p><a:r><a:t>Speaker details</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>
        """);
    using var image = archive.CreateEntry("ppt/media/image.png").Open();
    image.Write(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jQ1sAAAAASUVORK5CYII="));
}
var richDocument = PptxReader.Read(richPackage.ToArray());
var richMarkdown = PptxMarkdownConverter.Convert(richDocument);
Check(richMarkdown.Contains("Cell \\| one | Cell two"), "Markdown table cell escaping");
Check(richMarkdown.Contains("data:image/png;base64,"), "Markdown embedded picture");
Check(richMarkdown.Contains("Speaker details"), "Markdown speaker notes");
var plainMarkdown = PptxMarkdownConverter.Convert(richDocument, new(false, false));
Check(!plainMarkdown.Contains("Speaker details") && !plainMarkdown.Contains("data:image"), "Markdown options exclude notes and images");
var png = richDocument.GetPart("ppt/media/image.png");
var withImage = PptxImageEditor.Edit(document, 1, "", "add-image", png, 20, 30, 100, 80);
var imageInfo = PptxSlideExtractor.Extract(withImage).Slides[0].Images.Single();
Check(imageInfo.X == 20 && imageInfo.Width == 100, "Inserted image geometry");
Check(PptxSlideExtractor.Extract(document).Slides[0].Images.Count == 0, "Image insertion preserves source");
var imageMoved = PptxImageEditor.Edit(withImage, 1, imageInfo.Id, "resize-image", x: 50, y: 60, width: 200, height: 160);
Check(PptxSlideExtractor.Extract(imageMoved).Slides[0].Images.Single().X == 50, "Image position edit");
var imageReplaced = PptxImageEditor.Edit(imageMoved, 1, imageInfo.Id, "replace-image", png);
var replacedInfo = PptxSlideExtractor.Extract(PptxReader.Read(PptxWriter.Write(imageReplaced))).Slides[0].Images.Single();
Check(replacedInfo.PartName != imageInfo.PartName && replacedInfo.X == 50, "Image replacement uses new resource and retains position");
Check(imageReplaced.GetPart(imageInfo.PartName).SequenceEqual(png), "Old image resource unchanged for shared references");
Check(PptxSlideExtractor.Extract(PptxImageEditor.Edit(imageReplaced, 1, imageInfo.Id, "delete-image")).Slides[0].Images.Count == 0, "Delete image instance");
Reject(() => PptxImageEditor.Edit(document, 1, "", "add-image", [1, 2, 3]));
Reject(() => PptxImageEditor.Edit(withImage, 1, imageInfo.Id, "resize-image", width: -1));
Reject(() => PptxImageEditor.Edit(withImage, 1, imageInfo.Id, "resize-image", x: double.NaN));
var combinedImageEdit = PresentationApplication.Edit(PptxWriter.Write(withImage), new("replace-image", 1, imageInfo.Id, ImageBase64: Convert.ToBase64String(png), X: 5, Y: 6, Width: 80, Height: 60, UpdateBounds: true));
var combinedImage = PptxSlideExtractor.Extract(PptxReader.Read(combinedImageEdit)).Slides[0].Images.Single();
Check(combinedImage.Width == 80 && combinedImage.X == 5 && combinedImage.PartName != imageInfo.PartName, "Replace image and apply source dimensions atomically");
Console.WriteLine($"Passed {checks} package, editing, Markdown and validation checks.");
