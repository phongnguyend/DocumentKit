using System.Text.Json;
using DocxObjectKit.Extensions.Pdf;
using System.Xml;
using DocxObjectKit.Application;
using DocxObjectKit.Converters;
using DocxObjectKit.Readers;

const int limit = 50 * 1024 * 1024;
var builder = WebApplication.CreateBuilder(args);
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = limit + 1024 * 1024;
});
builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(options =>
{
    options.MemoryBufferThreshold = limit;
    options.MultipartBodyLengthLimit = limit;
});
var app = builder.Build();
app.Use(async (context, next) =>
{
    try
    {
        await next();
    }
    catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or ArgumentException or XmlException or JsonException or InvalidOperationException or OverflowException)
    {
        await Results.Problem("Cannot process this document: " + exception.Message, statusCode: 422).ExecuteAsync(context);
    }
});
app.MapGet("/api/health", () => Results.Ok(new { status = "healthy" }));
app.MapPost("/api/documents/inspect", async (HttpRequest request) =>
{
    return Results.Ok(DocumentApplication.Inspect(await Read(request), request.HttpContext.RequestAborted));
});
app.MapPost("/api/documents/markdown", async (HttpRequest request) =>
{
    var document = DocxReader.Read(await Read(request));
    var images = request.Query["images"] != "false";
    return Results.Ok(new { markdown = DocxMarkdownConverter.Convert(document, images) });
});
app.MapPost("/api/documents/pdf", async (HttpRequest request) =>
{
    var document = DocxReader.Read(await Read(request));
    int[]? pages = null;
    if (request.Query.TryGetValue("pages", out var selected))
    {
        pages = selected.ToString().Split(',').Select(value => int.TryParse(value, out var number) ? number : throw new ArgumentException("Page numbers must be comma-separated integers.")).ToArray();
    }
    var result = DocxPdfConverter.Convert(document, new DocxPdfOptions(pages, request.Query["strict"] == "true"), request.HttpContext.RequestAborted);
    return Results.Ok(new { pdfBase64 = Convert.ToBase64String(result.Bytes), result.PageCount, result.Warnings });
});
app.MapPost("/api/documents/edit", async (HttpRequest request) =>
{
    var form = await request.ReadFormAsync(request.HttpContext.RequestAborted);
    var file = form.Files.GetFile("file") ?? throw new ArgumentException("Missing DOCX file.");
    var edit = JsonSerializer.Deserialize<DocxEditRequest>(form["edit"].ToString(), new JsonSerializerOptions(JsonSerializerDefaults.Web))
        ?? throw new ArgumentException("Missing edit request.");
    using var memory = new MemoryStream();
    await file.CopyToAsync(memory, request.HttpContext.RequestAborted);
    var bytes = DocumentApplication.Edit(memory.ToArray(), edit);
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "edited.docx");
});
app.Run();

static async Task<byte[]> Read(HttpRequest request)
{
    using var memory = new MemoryStream();
    var buffer = new byte[81920];
    int length;
    while ((length = await request.Body.ReadAsync(buffer, request.HttpContext.RequestAborted)) > 0)
    {
        if (memory.Length + length > limit)
        {
            throw new InvalidDataException("Documents must be no larger than 50 MB.");
        }
        memory.Write(buffer, 0, length);
    }
    return memory.ToArray();
}
