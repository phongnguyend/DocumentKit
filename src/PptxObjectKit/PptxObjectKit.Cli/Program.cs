using System.Text.Json;
using PptxObjectKit.Extensions.Pdf;
using PptxObjectKit.Application;
using PptxObjectKit.Extensions.Markdown;
using PptxObjectKit.Extractors;
using PptxObjectKit.Readers;
using PptxObjectKit.Writers;

if (args.Length < 3 || args[0] is not ("inspect" or "to-markdown" or "rewrite" or "edit" or "to-pdf"))
{
    Console.WriteLine("Usage: pptxobjectkit <inspect|to-markdown|to-pdf|rewrite> input.pptx output\n       pptxobjectkit edit input.pptx output.pptx edit.json\n       pptxobjectkit to-pdf input.pptx output.pdf [options.json]\nOutput must not already exist. Editing supports text, reorder, and delete-slide.");
    return args.Length == 1 && args[0] == "--help" ? 0 : 1;
}
try
{
    var input = await File.ReadAllBytesAsync(args[1]);
    var document = PptxReader.Read(input);
    var options = new JsonSerializerOptions(JsonSerializerDefaults.Web) { WriteIndented = true };
    byte[] output;
    switch (args[0])
    {
        case "inspect":
            output = JsonSerializer.SerializeToUtf8Bytes(PptxSlideExtractor.Extract(document), options);
            break;
        case "to-markdown":
            output = System.Text.Encoding.UTF8.GetBytes(PptxMarkdownConverter.Convert(document));
            break;
        case "to-pdf":
            var pdfOptions = args.Length > 3 ? JsonSerializer.Deserialize<PptxPdfOptions>(await File.ReadAllTextAsync(args[3]), options) : null;
            var pdf = PptxPdfConverter.Convert(document, pdfOptions);
            output = pdf.Bytes;
            foreach (var warning in pdf.Warnings)
            {
                Console.Error.WriteLine(warning);
            }
            break;
        case "edit":
            if (args.Length != 4)
            {
                throw new ArgumentException("Provide an edit JSON file.");
            }
            output = PresentationApplication.Edit(input, JsonSerializer.Deserialize<PptxEditRequest>(await File.ReadAllTextAsync(args[3]), options)!);
            break;
        default:
            output = PptxWriter.Write(document);
            break;
    }
    await using var stream = new FileStream(args[2], FileMode.CreateNew, FileAccess.Write);
    await stream.WriteAsync(output);
    Console.WriteLine($"Saved {Path.GetFullPath(args[2])}");
    return 0;
}
catch (Exception exception)
{
    Console.Error.WriteLine(exception.Message);
    return 1;
}
