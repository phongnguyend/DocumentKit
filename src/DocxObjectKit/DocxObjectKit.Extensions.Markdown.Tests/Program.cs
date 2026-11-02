using DocxObjectKit.Readers;
using DocxObjectKit.Extensions.Markdown;

if (typeof(DocxObjectKit.Extensions.Markdown.DocxMarkdownConverter).Assembly.GetName().Name != "DocxObjectKit.Extensions.Markdown"
    || typeof(DocxObjectKit.Models.DocxObjectDocument).Assembly.GetReferencedAssemblies().Any(reference => reference.Name == "DocxObjectKit.Extensions.Markdown"))
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
var source = DocxReader.Read(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures/sample.docx")));
var markdown = DocxMarkdownConverter.Convert(source);
Check(markdown.Contains("# Document toolkit sample") && markdown.Contains("**Red bold text**"), "Markdown headings or formatting missing.");
Check(markdown.Contains("| Feature | Status |") && markdown.Contains("data:image/png;base64,"), "Markdown tables/images missing.");
Check(!DocxMarkdownConverter.Convert(source, false).Contains("data:image"), "Image exclusion ignored.");
Console.WriteLine($"Passed {count} DOCX Markdown checks.");
