namespace PdfObjectKit.Options;

public sealed class PdfDocxOptions
{
    /// <summary>One-based pages in output order. Null selects all pages; limited to 100 pages.</summary>
    public IReadOnlyList<int>? PageNumbers { get; init; }

    public bool IncludeImages { get; init; } = true;

    /// <summary>Start each source page on a new Word page. Reflow can still increase page counts.</summary>
    public bool PageBreaks { get; init; } = true;
}
