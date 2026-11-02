using PptxObjectKit.Readers;
using PptxObjectKit.Extensions.Pdf;
using Microsoft.AspNetCore.Http.Features;

const long MaxFileSize = 10 * 1024 * 1024;
var builder = WebApplication.CreateBuilder(args);
builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = MaxFileSize + (1024 * 1024);
});
builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = MaxFileSize + (1024 * 1024);
});
var app = builder.Build();
app.MapGet("/health", () => Results.Ok(new { status = "healthy" }));
app.MapPost("/convert", async (IFormFile file, HttpContext context, CancellationToken cancellationToken) =>
{
    if (file.Length == 0 || file.Length > MaxFileSize)
    {
        return Results.BadRequest("Upload a non-empty file no larger than 10 MB.");
    }

    var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
    if (extension != ".pptx")
    {
        return Results.BadRequest("This engine accepts .pptx files only. Use LibreOffice for legacy .ppt files.");
    }

    try
    {
        await using var input = file.OpenReadStream();
        using var bytes = new MemoryStream();
        await input.CopyToAsync(bytes, cancellationToken);
        var document = PptxReader.Read(bytes.ToArray());
        var result = PptxPdfConverter.Convert(document, cancellationToken: cancellationToken);
        context.Response.Headers["X-Conversion-Warning-Count"] = result.Warnings.Count.ToString(System.Globalization.CultureInfo.InvariantCulture);
        var pdf = result.Bytes;
        var downloadName = Path.GetFileNameWithoutExtension(file.FileName.Replace('\\', '/')) + ".pdf";
        return Results.File(pdf, "application/pdf", downloadName);
    }
    catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or ArgumentException or System.Xml.XmlException or FormatException or OverflowException or EndOfStreamException)
    {
        return Results.Problem(exception.Message, statusCode: StatusCodes.Status422UnprocessableEntity);
    }
})
.DisableAntiforgery()
.WithName("ConvertPowerPointToPdf");
app.Run();
