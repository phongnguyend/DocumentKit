using System.Text.Json;
using System.Xml;
using PptxObjectKit.Application;
using PptxObjectKit.Converters;
using PptxObjectKit.Extractors;
using PptxObjectKit.Readers;

const int limit = 50 * 1024 * 1024;
var builder = WebApplication.CreateBuilder(args);
builder.WebHost.ConfigureKestrel(options => options.Limits.MaxRequestBodySize = limit + 1024 * 1024);
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
        await Results.Problem("Cannot process this presentation: " + exception.Message, statusCode: 422).ExecuteAsync(context);
    }
});
app.MapGet("/api/health", () => Results.Ok(new { status = "healthy" }));
app.MapPost("/api/documents/inspect", async (HttpRequest request) =>
{
    var document = PptxReader.Read(await Read(request));
    return Results.Ok(PptxSlideExtractor.Extract(document));
});
app.MapPost("/api/documents/markdown", async (HttpRequest request) =>
{
    var document = PptxReader.Read(await Read(request));
    var options = new PptxMarkdownOptions(request.Query["notes"] != "false", request.Query["images"] != "false");
    return Results.Ok(new { markdown = PptxMarkdownConverter.Convert(document, options) });
});
app.MapPost("/api/documents/edit", async (HttpRequest request) =>
{
    var form = await request.ReadFormAsync(request.HttpContext.RequestAborted);
    var file = form.Files.GetFile("file") ?? throw new ArgumentException("Missing PPTX file.");
    var edit = JsonSerializer.Deserialize<PptxEditRequest>(form["edit"].ToString(), new JsonSerializerOptions(JsonSerializerDefaults.Web))
        ?? throw new ArgumentException("Missing edit request.");
    using var memory = new MemoryStream();
    await file.CopyToAsync(memory, request.HttpContext.RequestAborted);
    var bytes = PresentationApplication.Edit(memory.ToArray(), edit);
    return Results.File(bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "edited.pptx");
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
            throw new InvalidDataException("Presentations must be no larger than 50 MB.");
        }
        memory.Write(buffer, 0, length);
    }
    return memory.ToArray();
}
