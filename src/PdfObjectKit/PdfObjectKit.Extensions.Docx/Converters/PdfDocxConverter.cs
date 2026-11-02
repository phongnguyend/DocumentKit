using PdfObjectKit.Analysis;
using PdfObjectKit.Diagnostics;
using PdfObjectKit.Graphs;
using PdfObjectKit.Models;
using PdfObjectKit.Options;

namespace PdfObjectKit.Converters;

/// <summary>Exports editable, reflowable Word content; does not reproduce fixed PDF page layout.</summary>
public static class PdfDocxConverter
{
    public static PdfDocxResult Convert(PdfObjectDocument document, PdfDocxOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(document);
        options ??= new();
        cancellationToken.ThrowIfCancellationRequested();
        if (document.Security is not null)
        {
            throw new NotSupportedException("Unlock the PDF before exporting Word content.");
        }
        var pages = new PdfGraph(document).Pages();
        var selected = options.PageNumbers?.ToArray() ?? Enumerable.Range(1, pages.Count).ToArray();
        if (selected.Length is < 1 or > 100 || selected.Any(page => page < 1 || page > pages.Count) || selected.Distinct().Count() != selected.Length)
        {
            throw new ArgumentException("Choose 1-100 distinct, valid page numbers.");
        }
        var structure = PdfDocumentAnalyzer.Analyze(document);
        var warnings = new HashSet<string>
        {
            "Word export uses reflowable paragraphs and heuristic reading order, headings and tables. Page counts and layout may differ from the PDF.",
            "Fonts are referenced by name, not embedded. Text colors, vector graphics, clipping, image rotation/cropping, vertical table merges, interactive fields, annotations and document metadata are not preserved. Images contain their full source pixels. Review before sharing."
        };
        foreach (var diagnostic in structure.Diagnostics.Where(item => item.Severity >= PdfDiagnosticSeverity.Warning
            && item.Code != PdfDiagnosticCode.ImagesOmitted && (item.PageNumber is null || selected.Contains(item.PageNumber.Value))))
        {
            warnings.Add(diagnostic.Message);
        }
        using var writer = new DocxPackage();
        for (var index = 0; index < selected.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var page = structure.Pages.Single(item => item.PageNumber == selected[index]);
            if (index > 0 && options.PageBreaks)
            {
                writer.PageBreak();
            }
            foreach (var block in page.Blocks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (block.Kind == PdfContentBlockKind.Image)
                {
                    if (options.IncludeImages && block.Image?.Source is { } image)
                    {
                        try
                        {
                            writer.Image(document, page.PageNumber, image);
                        }
                        catch (Exception ex) when (ex is NotSupportedException or InvalidDataException)
                        {
                            warnings.Add($"Page {page.PageNumber}, image {image.ImageNumber}: image omitted. {ex.Message}");
                        }
                    }
                }
                else if (block.Kind == PdfContentBlockKind.Table && block.Table is { } table)
                {
                    writer.Table(table);
                }
                else
                {
                    writer.Paragraph(block);
                }
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new(writer.Finish(), selected.Length, writer.ParagraphCount, writer.TableCount, writer.ImageCount, warnings.ToArray());
    }
}
