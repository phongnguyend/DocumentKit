using System.IO.Compression;
using System.Text;
using PptxObjectKit.Readers;
using PptxObjectKit.Editors;
using PptxObjectKit.Extensions.Markdown;

if (typeof(PptxObjectKit.Extensions.Markdown.PptxMarkdownConverter).Assembly.GetName().Name != "PptxObjectKit.Extensions.Markdown"
    || typeof(PptxObjectKit.Models.PptxObjectDocument).Assembly.GetReferencedAssemblies().Any(reference => reference.Name == "PptxObjectKit.Extensions.Markdown"))
{
    throw new Exception("Markdown must live in an optional extension without a core dependency.");
}

var count = 0;
void Check(bool condition, string message)
{
    if (!condition)
    {
        throw new Exception(message);
    }
    count++;
}
var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures/sample.pptx"));
var document = PptxReader.Read(bytes);
var edited = PptxSlideEditor.UpdateText(document, 1, "2", "Changed title\nSecond paragraph");
var markdown = PptxMarkdownConverter.Convert(edited);
Check(markdown.Contains("Changed title") && markdown.Contains("Second paragraph") && markdown.Contains("# Slide 2:"), "Markdown paragraphs and order");
var escaped = PptxSlideEditor.UpdateText(document, 1, "2", "<script>alert(1)</script>");
Check(!PptxMarkdownConverter.Convert(escaped).Contains("<script>"), "Escape HTML in Markdown");
var richDocument = RichPresentationFixture.Create(bytes);
var richMarkdown = PptxMarkdownConverter.Convert(richDocument);
Check(richMarkdown.Contains("Cell \\| one | Cell two"), "Markdown table cell escaping");
Check(richMarkdown.Contains("data:image/png;base64,"), "Markdown embedded picture");
Check(richMarkdown.Contains("Speaker details"), "Markdown speaker notes");
var plainMarkdown = PptxMarkdownConverter.Convert(richDocument, new(false, false));
Check(!plainMarkdown.Contains("Speaker details") && !plainMarkdown.Contains("data:image"), "Markdown options exclude notes and images");
Console.WriteLine($"Passed {count} PPTX Markdown checks.");
