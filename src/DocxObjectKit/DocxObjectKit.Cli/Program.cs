using DocxObjectKit.Application;
using DocxObjectKit.Readers;
using DocxObjectKit.Converters;
using DocxObjectKit.Extensions.Pdf;
using System.Text.Json;

try
{
    if (args.Length < 2)
    {
        Console.Error.WriteLine("Usage: DocxObjectKit.Cli inspect input.docx | markdown input.docx output.md | to-pdf input.docx output.pdf | edit input.docx output.docx request.json");
        return 1;
    }
    var bytes = File.ReadAllBytes(args[1]);
    if (args[0] == "inspect")
    {
        Console.WriteLine(JsonSerializer.Serialize(DocumentApplication.Inspect(bytes), new JsonSerializerOptions(JsonSerializerDefaults.Web) { WriteIndented = true }));
        return 0;
    }
    if (args.Length < 3)
    {
        throw new ArgumentException("An output path is required.");
    }
    byte[] output;
    switch (args[0])
    {
        case "markdown":
            output = System.Text.Encoding.UTF8.GetBytes(DocxMarkdownConverter.Convert(DocxReader.Read(bytes)));
            break;
        case "to-pdf":
            var result = DocxPdfConverter.Convert(DocxReader.Read(bytes));
            output = result.Bytes;
            foreach (var warning in result.Warnings)
            {
                Console.Error.WriteLine(warning);
            }
            break;
        case "edit" when args.Length == 4:
            output = DocumentApplication.Edit(bytes, JsonSerializer.Deserialize<DocxEditRequest>(File.ReadAllText(args[3]), new JsonSerializerOptions(JsonSerializerDefaults.Web)) ?? throw new ArgumentException("Invalid edit request."));
            break;
        default:
            throw new ArgumentException("Unknown command.");
    }
    using var file = new FileStream(args[2], FileMode.CreateNew);
    file.Write(output);
    return 0;
}
catch (Exception exception)
{
    Console.Error.WriteLine(exception.Message);
    return 1;
}
