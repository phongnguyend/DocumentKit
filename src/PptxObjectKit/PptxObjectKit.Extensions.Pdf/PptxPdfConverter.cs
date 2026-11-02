using PptxObjectKit.Models;
using PptxObjectKit.Extractors;
using SkiaSharp;

namespace PptxObjectKit.Extensions.Pdf;

public sealed class PptxPdfOptions
{
    public IReadOnlyList<int>? Slides { get; init; }

    public bool IncludeHiddenSlides { get; init; } = true;

    public bool FailOnWarnings { get; init; }
}

public sealed record PptxPdfResult(byte[] Bytes, int PageCount, IReadOnlyList<string> Warnings);

public static class PptxPdfConverter
{
    public static PptxPdfResult Convert(PptxObjectDocument document, PptxPdfOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(document);
        options ??= new();
        var info = PptxSlideExtractor.Extract(document);
        var pages = options.Slides?.ToArray() ?? Enumerable.Range(1, info.Slides.Count).ToArray();
        if (pages.Length == 0 || pages.Length > 500 || pages.Distinct().Count() != pages.Length || pages.Any(page => page < 1 || page > info.Slides.Count))
        {
            throw new ArgumentException("Choose 1 to 500 distinct, valid slide numbers.");
        }
        var warnings = new HashSet<string>(StringComparer.Ordinal);
        using var renderer = new SlideRenderer(document, warnings, cancellationToken);
        using var output = new MemoryStream();
        // A zero-initialized metadata struct otherwise requests quality-zero JPEG encoding.
        using var pdf = SKDocument.CreatePdf(output, new SKDocumentPdfMetadata
        {
            Creator = "PptxObjectKit native PDF converter",
            Producer = "PptxObjectKit / Skia",
            RasterDpi = 144,
            EncodingQuality = SKDocumentPdfMetadata.DefaultEncodingQuality
        });
        if (pdf is null)
        {
            throw new InvalidOperationException("The native PDF writer could not be initialized.");
        }
        var count = 0;
        foreach (var page in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var scene = renderer.Load(info.Slides[page - 1].PartName, page);
            if (!options.IncludeHiddenSlides && (string?)scene.Xml.Root?.Attribute("show") == "0")
            {
                continue;
            }
            var canvas = pdf.BeginPage((float)info.Width, (float)info.Height);
            renderer.Render(scene, canvas, (float)info.Width, (float)info.Height);
            pdf.EndPage();
            count++;
        }
        if (count == 0)
        {
            throw new ArgumentException("No visible slides were selected.");
        }
        if (options.FailOnWarnings && warnings.Count > 0)
        {
            throw new NotSupportedException("Strict PDF conversion rejected: " + string.Join(" ", warnings));
        }
        pdf.Close();
        return new(output.ToArray(), count, warnings.Order(StringComparer.Ordinal).ToArray());
    }
}
