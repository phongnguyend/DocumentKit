using PdfObjectKit.Models;

namespace PdfObjectKit.Options;

/// <summary>Shared options for layout analysis, independent of output format.</summary>
public class PdfAnalysisOptions
{
    /// <summary>Explicit table regions in unrotated PDF coordinates. Regions must not overlap.</summary>
    public IReadOnlyList<PdfTableHint> TableHints { get; init; } = [];

    public PdfHeuristicOptions Heuristics { get; init; } = new();

    /// <summary>Optional body text height in PDF user space; otherwise estimated by character-weighted frequency.</summary>
    public double? BodyTextHeight { get; init; }

    /// <summary>Opt-in removal of hyphens between alphabetic wrapped words. Can alter legitimate compound words.</summary>
    public bool JoinHyphenatedWords { get; init; }
}
