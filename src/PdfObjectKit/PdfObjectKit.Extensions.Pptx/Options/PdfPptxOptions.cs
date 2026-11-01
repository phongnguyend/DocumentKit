namespace PdfObjectKit.Options;

/// <summary>Editable slide export. Sizes are in points; all slides share one size.</summary>
public sealed class PdfPptxOptions
{
    /// <summary>One-based page numbers in output order. Null exports all pages.</summary>
    public IReadOnlyList<int>? PageNumbers { get; init; }

    public double? SlideWidth { get; init; }

    public double? SlideHeight { get; init; }

    public bool IncludeImages { get; init; } = true;

    public bool IncludePaths { get; init; } = true;
}
