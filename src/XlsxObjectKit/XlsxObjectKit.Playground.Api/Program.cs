using System.Text.Json;
using System.Xml;
using XlsxObjectKit.Application;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Writers;

const int limit = 50 * 1024 * 1024;
const string mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
var builder = WebApplication.CreateBuilder(args);
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = limit + 1024 * 1024;
});
builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = limit + 1024 * 1024;
});
var app = builder.Build();
app.Use(async (context, next) =>
{
    try
    {
        await next();
    }
    catch (Exception exception) when (exception is InvalidDataException or ArgumentException or NotSupportedException or XmlException or JsonException or InvalidOperationException or OverflowException)
    {
        await Results.Problem("Cannot process this workbook: " + exception.Message, statusCode: 422).ExecuteAsync(context);
    }
});
app.MapGet("/api/health", () => Results.Ok(new { status = "healthy" }));
app.MapPost("/api/documents/new", () => Results.File(XlsxWriter.Write(XlsxEditor.Create()), mime));
app.MapPost("/api/documents/inspect", async (HttpRequest request) => Results.Ok(WorkbookApplication.Inspect(await Read(request.Body, request.HttpContext.RequestAborted))));
app.MapPost("/api/documents/edit", async (HttpRequest request) =>
{
    var form = await request.ReadFormAsync(request.HttpContext.RequestAborted);
    var file = form.Files.GetFile("file") ?? throw new ArgumentException("Missing XLSX file.");
    var edit = JsonSerializer.Deserialize<XlsxEditRequest>(form["edit"].ToString(), new JsonSerializerOptions(JsonSerializerDefaults.Web)) ?? throw new ArgumentException("Missing edit request.");
    using var input = file.OpenReadStream();
    var result = WorkbookApplication.EditWithDiagnostics(await Read(input, request.HttpContext.RequestAborted), edit);
    request.HttpContext.Response.Headers["X-Calculation-Warnings"] = Uri.EscapeDataString(JsonSerializer.Serialize(result.Warnings.Take(10).Select(warning => warning.Length > 300 ? warning[..300] : warning)));
    return Results.File(result.Bytes, mime, "edited.xlsx");
});
app.Run();

static async Task<byte[]> Read(Stream stream, CancellationToken cancellation)
{
    using var memory = new MemoryStream();
    var buffer = new byte[81920];
    int count;
    while ((count = await stream.ReadAsync(buffer, cancellation)) > 0)
    {
        if (memory.Length + count > limit)
        {
            throw new InvalidDataException("Workbooks must be no larger than 50 MB.");
        }
        memory.Write(buffer, 0, count);
    }
    return memory.ToArray();
}
