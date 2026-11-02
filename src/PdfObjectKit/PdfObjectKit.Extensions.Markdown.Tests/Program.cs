if (typeof(PdfObjectKit.Extensions.Markdown.PdfMarkdownConverter).Assembly.GetName().Name != "PdfObjectKit.Extensions.Markdown"
    || typeof(PdfObjectKit.Models.PdfObjectDocument).Assembly.GetReferencedAssemblies().Any(reference => reference.Name == "PdfObjectKit.Extensions.Markdown"))
{
    throw new Exception("Markdown must live in an optional extension without a core dependency.");
}

MarkdownTests.Run();
MarkdownImageTests.Run();
