using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using PptxObjectKit.Application;
using PptxObjectKit.Extensions.Markdown;
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
var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jQ1sAAAAASUVORK5CYII=");
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
var ocrLines = new PptxOcrLine[] { new("Hello OCR", .1, .2, .5, .15, .32), new("Vietnamese: Xin ch?o", .1, .5, .7, .2, .65) };
var ocrKeep = PptxImageOcrEditor.Apply(withImage, 1, imageInfo.Id, ocrLines);
var ocrRoundTrip = PptxReader.Read(PptxWriter.Write(ocrKeep));
Check(PptxSlideExtractor.Extract(ocrRoundTrip).Slides[0].Images.Count == 1, "OCR keep image mode");
Check(PptxSlideExtractor.Extract(ocrRoundTrip).Slides[0].TextObjects.Any(item => item.Text == "Hello OCR"), "OCR creates extractable text");
Check(Encoding.UTF8.GetString(ocrRoundTrip.GetPart("ppt/slides/slide1.xml")).Contains("alpha val=\"0\""), "OCR keep image text is transparent");
var ocrReplace = PptxImageOcrEditor.Apply(withImage, 1, imageInfo.Id, ocrLines, true);
Check(PptxSlideExtractor.Extract(ocrReplace).Slides[0].Images.Count == 0, "OCR replacement removes image instance");
Check(PptxMarkdownConverter.Convert(ocrReplace).Contains("Hello OCR"), "OCR text included in Markdown");
var ocrTextId = PptxSlideExtractor.Extract(ocrReplace).Slides[0].TextObjects.First(item => item.Text == "Hello OCR").Id;
Check(PptxSlideExtractor.Extract(PptxSlideEditor.UpdateText(ocrReplace, 1, ocrTextId, "Corrected")).Slides[0].TextObjects.Any(item => item.Text == "Corrected"), "OCR boxes remain editable");
Check(PptxSlideExtractor.Extract(withImage).Slides[0].TextObjects.All(item => item.Text != "Hello OCR"), "OCR preserves source document");
Reject(() => PptxImageOcrEditor.Apply(withImage, 1, imageInfo.Id, [new("Invalid", .8, 0, .5, .1)]));
Reject(() => PptxImageOcrEditor.Apply(withImage, 1, imageInfo.Id, [new("Invalid", double.NaN, 0, .5, .1)]));
Reject(() => PptxImageOcrEditor.Apply(withImage, 1, imageInfo.Id, []));
using var transformedPackage = new MemoryStream();
transformedPackage.Write(PptxWriter.Write(withImage));
using (var archive = new ZipArchive(transformedPackage, ZipArchiveMode.Update, true))
{
    var entry = archive.GetEntry("ppt/slides/slide1.xml")!;
    XDocument xml;
    using (var stream = entry.Open())
    {
        xml = XDocument.Load(stream);
    }
    XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
    XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    xml.Descendants(p + "pic").Single().Element(p + "spPr")!.Element(a + "xfrm")!.SetAttributeValue("rot", 5400000);
    entry.Delete();
    using var output = archive.CreateEntry("ppt/slides/slide1.xml").Open();
    xml.Save(output);
}
var rotatedOcr = PptxImageOcrEditor.Apply(PptxReader.Read(transformedPackage.ToArray()), 1, imageInfo.Id, ocrLines, true);
Check(Encoding.UTF8.GetString(rotatedOcr.GetPart("ppt/slides/slide1.xml")).Contains("rot=\"5400000\""), "OCR retains image rotation in text group");
using (var archive = new ZipArchive(transformedPackage, ZipArchiveMode.Update, true))
{
    var entry = archive.GetEntry("ppt/slides/slide1.xml")!;
    XDocument xml;
    using (var stream = entry.Open())
    {
        xml = XDocument.Load(stream);
    }
    XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
    XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    xml.Descendants(p + "pic").Single().Element(p + "blipFill")!.Add(new XElement(a + "srcRect", new XAttribute("l", 10000)));
    entry.Delete();
    using var output = archive.CreateEntry("ppt/slides/slide1.xml").Open();
    xml.Save(output);
}
Reject(() => PptxImageOcrEditor.Apply(PptxReader.Read(transformedPackage.ToArray()), 1, imageInfo.Id, ocrLines));
Reject(() => PptxImageOcrEditor.Apply(PptxReader.Read(AddPart("_xmlsignatures/sig1.xml")), 1, imageInfo.Id, ocrLines));
var duplicate = PptxSlideEditor.Duplicate(RichPresentationFixture.Create(bytes), 1);
var duplicateRoundTrip = PptxSlideExtractor.Extract(PptxReader.Read(PptxWriter.Write(duplicate)));
Check(duplicateRoundTrip.Slides.Count == 3 && duplicateRoundTrip.Slides[0].Title == duplicateRoundTrip.Slides[1].Title, "Duplicate inserted immediately after source");
Check(duplicateRoundTrip.Slides[0].PartName != duplicateRoundTrip.Slides[1].PartName && duplicateRoundTrip.Slides.Select(slide => slide.Id).Distinct().Count() == 3, "Duplicate has independent part and unique slide ID");
Check(duplicateRoundTrip.Slides[1].Notes == "Speaker details" && duplicateRoundTrip.Slides[1].PictureCount == 1, "Duplicate preserves notes and pictures");
Check(duplicate.PartNames.Count(part => part.StartsWith("ppt/notesSlides/", StringComparison.Ordinal) && part.EndsWith(".xml", StringComparison.Ordinal)) == 2, "Duplicate owns an independent notes part");
var editedDuplicate = PptxSlideEditor.UpdateText(duplicate, 2, duplicateRoundTrip.Slides[1].TextObjects[0].Id, "Only duplicate changed");
Check(PptxSlideExtractor.Extract(editedDuplicate).Slides[0].Title == duplicateRoundTrip.Slides[0].Title && PptxSlideExtractor.Extract(editedDuplicate).Slides[1].Title == "Only duplicate changed", "Duplicate text changes do not alter source");
var blankAfter = PptxSlideExtractor.Extract(PptxReader.Read(PptxWriter.Write(PptxSlideEditor.InsertBlank(document, 1))));
Check(blankAfter.Slides.Count == 3 && blankAfter.Slides[1].TextObjects.Count == 0 && blankAfter.Slides[1].PictureCount == 0 && blankAfter.Slides[1].Notes == "", "Blank inserted without objects or notes");
Check(blankAfter.Width == inspection.Width && blankAfter.Height == inspection.Height && blankAfter.Slides[2].Title == "Second slide", "Blank retains presentation size and following slide");
var blankEnd = PptxSlideExtractor.Extract(PptxSlideEditor.InsertBlank(document, 2));
Check(blankEnd.Slides[1].Title == "Second slide" && blankEnd.Slides[2].TextObjects.Count == 0, "Blank appended at end");
Check(PptxSlideExtractor.Extract(document).Slides.Count == 2, "Insertion preserves input document");
Reject(() => PptxSlideEditor.InsertBlank(document, 0));
Reject(() => PptxSlideEditor.Duplicate(document, 3));
Reject(() => PptxSlideEditor.Duplicate(PptxReader.Read(AddPart("_xmlsignatures/sig1.xml")), 1));
var textBox = inspection.Slides[0].TextObjects[0];
Check(textBox.Bounds is { Width: > 0, Height: > 0 }, "Text bounds extracted");
var textMoved = PptxSlideEditor.TransformText(document, 1, textBox.Id, 40, 50, 240, 60);
var movedBox = PptxSlideExtractor.Extract(PptxReader.Read(PptxWriter.Write(textMoved))).Slides[0].TextObjects[0];
Check(movedBox.Bounds is { X: 40, Y: 50, Width: 240, Height: 60 }, "Text box move and resize round trip");
XNamespace presentationNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
var originalBody = XDocument.Parse(Encoding.UTF8.GetString(document.GetPart(inspection.Slides[0].PartName))).Descendants(presentationNs + "txBody").First();
var movedBody = XDocument.Parse(Encoding.UTF8.GetString(textMoved.GetPart(inspection.Slides[0].PartName))).Descendants(presentationNs + "txBody").First();
Check(XNode.DeepEquals(originalBody, movedBody), "Text transform preserves content, font sizes and wrapping");
Check(PptxSlideExtractor.Extract(document).Slides[0].TextObjects[0].Bounds == textBox.Bounds, "Text transform leaves source unchanged");
Reject(() => PptxSlideEditor.TransformText(document, 1, textBox.Id, double.NaN, 0, 100, 50));
Reject(() => PptxSlideEditor.TransformText(document, 1, textBox.Id, 0, 0, 0, 50));
Reject(() => PptxSlideEditor.TransformText(ocrReplace, 1, ocrTextId, 0, 0, 100, 50));
using var inheritedPackage = new MemoryStream();
inheritedPackage.Write(PptxWriter.Write(document));
using (var archive = new ZipArchive(inheritedPackage, ZipArchiveMode.Update, true))
{
    var entry = archive.GetEntry("ppt/slides/slide1.xml")!;
    XDocument xml;
    using (var stream = entry.Open())
    {
        xml = XDocument.Load(stream);
    }
    XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
    var shape = xml.Descendants(presentationNs + "sp").First();
    var frame = shape.Element(presentationNs + "spPr")!.Element(a + "xfrm")!;
    var inheritedFrame = new XElement(frame);
    inheritedFrame.SetAttributeValue("rot", 1800000);
    frame.Remove();
    shape.Element(presentationNs + "nvSpPr")!.Element(presentationNs + "nvPr")!.Add(new XElement(presentationNs + "ph", new XAttribute("idx", 42)));
    entry.Delete();
    using (var output = archive.CreateEntry("ppt/slides/slide1.xml").Open())
    {
        xml.Save(output);
    }
    archive.GetEntry("ppt/slides/_rels/slide1.xml.rels")?.Delete();
    using (var writer = new StreamWriter(archive.CreateEntry("ppt/slides/_rels/slide1.xml.rels").Open()))
    {
        writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"layout\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/inherited.xml\"/></Relationships>");
    }
    var layout = new XDocument(new XElement(presentationNs + "sldLayout", new XElement(presentationNs + "cSld", new XElement(presentationNs + "spTree", new XElement(presentationNs + "sp", new XElement(presentationNs + "nvSpPr", new XElement(presentationNs + "nvPr", new XElement(presentationNs + "ph", new XAttribute("idx", 42)))), new XElement(presentationNs + "spPr", inheritedFrame))))));
    using var layoutOutput = archive.CreateEntry("ppt/slideLayouts/inherited.xml").Open();
    layout.Save(layoutOutput);
}
var inheritedDocument = PptxReader.Read(inheritedPackage.ToArray());
Check(PptxSlideExtractor.Extract(inheritedDocument).Slides[0].TextObjects[0].Bounds?.Rotation == 30, "Placeholder geometry resolves from layout");
var movedPlaceholder = PptxSlideEditor.TransformText(inheritedDocument, 1, textBox.Id, 20, 30, 200, 70);
Check(PptxSlideExtractor.Extract(movedPlaceholder).Slides[0].TextObjects[0].Bounds is { X: 20, Rotation: 30 }, "Placeholder edit creates explicit frame and preserves rotation");
Check(movedPlaceholder.GetPart("ppt/slideLayouts/inherited.xml").SequenceEqual(inheritedDocument.GetPart("ppt/slideLayouts/inherited.xml")), "Placeholder transform does not modify shared layout");
Console.WriteLine($"Passed {checks} package, editing, Markdown and validation checks.");
