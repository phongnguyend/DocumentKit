using DocxObjectKit.Readers;
using DocxObjectKit.Writers;
using DocxObjectKit.Editors;
using static DocxObjectKit.Models.DocxNamespaces;


var count = 0;
void Check(bool value, string message)
{
    if (!value)
    {
        throw new Exception(message);
    }
    count++;
}
void Reject(Action action)
{
    try
    {
        action();
    }
    catch (Exception exception) when (exception is ArgumentException or InvalidDataException or NotSupportedException)
    {
        count++;
        return;
    }
    throw new Exception("Expected rejected input.");
}
var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures/sample.docx"));
var source = DocxReader.Read(bytes);
var content = DocxObjectKit.Extractors.DocxDocumentExtractor.Extract(source);
Check(content.Paragraphs[0].Text == "Document toolkit sample" && content.Images.Count == 2, "Semantic extraction failed.");
var edited = DocxEditor.SetText(source, "p0", "Updated title");
Check(source.ReadXml("word/document.xml").Descendants(W + "t").First().Value == "Document toolkit sample", "Edits mutate source.");
Check(edited.ReadXml("word/document.xml").Descendants(W + "t").First().Value == "Updated title", "Text replacement failed.");
var roundTrip = DocxReader.Read(DocxWriter.Write(edited));
Check(source.GetPart("customXml/preserved.xml").SequenceEqual(roundTrip.GetPart("customXml/preserved.xml")), "Unknown parts were not preserved.");
Check(source.GetPart("word/styles.xml").SequenceEqual(roundTrip.GetPart("word/styles.xml")), "Styles changed.");
var image = source.GetPart("word/media/sample.png");
var replaced = DocxEditor.Image(source, "replace-image", "i0", image, 0, 0, 100, 50);
var pictures = DocxEditor.Pictures(replaced.ReadXml("word/document.xml")).ToArray();
Check(DocxEditor.ImagePart(replaced, pictures[0]) != DocxEditor.ImagePart(replaced, pictures[1]), "Image replacement modifies shared resource.");
Check(DocxEditor.ImagePart(replaced, pictures[1]) == "word/media/sample.png", "Other placement changed.");
var moved = DocxEditor.Image(replaced, "resize-image", "i0", null, 120, 150, 200, 80);
var anchor = DocxEditor.Picture(moved.ReadXml("word/document.xml"), "i0");
Check(anchor.Name == WP + "anchor" && anchor.Element(WP + "positionH")!.Element(WP + "posOffset")!.Value == "1524000", "Image positioning not persisted.");
Check(DocxEditor.Pictures(DocxEditor.Image(moved, "delete-image", "i0", null, 0, 0, 1, 1).ReadXml("word/document.xml")).Count() == 1, "Image deletion failed.");
Check(DocxEditor.AddText(source, "New paragraph", "p0").ReadXml("word/document.xml").Descendants(W + "t").Any(item => item.Value == "New paragraph"), "Add text failed.");
Check(!DocxEditor.DeleteText(source, "p0").ReadXml("word/document.xml").Descendants(W + "t").Any(item => item.Value == "Document toolkit sample"), "Delete paragraph failed.");
Reject(() => DocxReader.Read([1, 2, 3]));
Reject(() => DocxEditor.SetText(source, "p99999", "bad"));
Reject(() => DocxEditor.Image(source, "replace-image", "i0", [1, 2], 0, 0, 1, 1));
Reject(() => DocxEditor.TransformText(source, "p0", double.NaN, 0, 10, 10));
Reject(() => DocxEditor.TransformText(source, "p3", 0, 0, 10, 10));
Console.WriteLine($"Passed {count} DOCX package/edit checks.");
