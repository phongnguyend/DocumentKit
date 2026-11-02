using PdfObjectKit.Options;
using PdfObjectKit.Models;

namespace PdfObjectKit.Extensions.Markdown;

public sealed class PdfMarkdownOptions : PdfAnalysisOptions
{
    /// <summary>
    /// Renders an image using its enriched metadata, one-based page number, and image occurrence number.
    /// Return Markdown without surrounding block separators, or an empty string to omit the image.
    /// Custom output is emitted verbatim; the caller is responsible for escaping it.
    /// </summary>
    public Func<PdfStructuredImage, int, int, string> ImageRender { get; init; } = DefaultMarkDownImageRender.Render;
}
