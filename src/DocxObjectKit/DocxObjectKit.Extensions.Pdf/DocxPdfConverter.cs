using DocxObjectKit.Models;
using SkiaSharp;

namespace DocxObjectKit.Extensions.Pdf;

public sealed record DocxPdfOptions(IReadOnlyList<int>? Pages = null, bool FailOnWarnings = false);
public sealed record DocxPdfResult(byte[] Bytes, int PageCount, IReadOnlyList<string> Warnings);

public static class DocxPdfConverter
{
    public static DocxPdfResult Convert(DocxObjectDocument document, DocxPdfOptions? options = null, CancellationToken cancellationToken = default)
    {
        options ??= new();
        using var engine = new DocxLayout(document);
        var layout = engine.Build(cancellationToken);
        var pages = options.Pages?.ToArray() ?? Enumerable.Range(1, layout.Pages.Count).ToArray();
        if (pages.Length == 0 || pages.Distinct().Count() != pages.Length || pages.Any(number => number < 1 || number > layout.Pages.Count))
        {
            throw new ArgumentException("Choose distinct, valid page numbers.");
        }
        if (options.FailOnWarnings && layout.Warnings.Length > 0)
        {
            throw new NotSupportedException(string.Join(" ", layout.Warnings));
        }
        using var output = new MemoryStream();
        using var pdf = SKDocument.CreatePdf(output, new SKDocumentPdfMetadata { Creator = "DocxObjectKit", Producer = "DocxObjectKit / Skia", RasterDpi = 144, EncodingQuality = SKDocumentPdfMetadata.DefaultEncodingQuality }) ?? throw new InvalidOperationException("PDF engine unavailable.");
        foreach (var number in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var page = layout.Pages[number - 1];
            var canvas = pdf.BeginPage((float)layout.Width, (float)layout.Height);
            canvas.Clear(SKColors.White);
            using var border = new SKPaint { Color = SKColors.Gray, Style = SKPaintStyle.Stroke, StrokeWidth = .5f, IsAntialias = true };
            foreach (var rule in page.Rules)
            {
                var rect = SKRect.Create((float)rule.X, (float)rule.Y, (float)rule.Width, (float)rule.Height);
                if (rule.Fill is { Length: 6 } fill && SKColor.TryParse("#" + fill, out var color))
                {
                    using var background = new SKPaint { Color = color };
                    canvas.DrawRect(rect, background);
                }
                if (rule.Edges is { } edges)
                {
                    SKPoint[] corners = [new(rect.Left, rect.Top), new(rect.Right, rect.Top), new(rect.Right, rect.Bottom), new(rect.Left, rect.Bottom)];
                    for (var edgeIndex = 0; edgeIndex < 4; edgeIndex++)
                    {
                        if (edges[edgeIndex] is not { Width: > 0 } edge)
                        {
                            continue;
                        }
                        border.Color = SKColor.TryParse("#" + edge.Color, out var edgeColor) ? edgeColor : SKColors.Black;
                        border.StrokeWidth = (float)edge.Width;
                        border.StrokeCap = SKStrokeCap.Square;
                        canvas.DrawLine(corners[edgeIndex], corners[(edgeIndex + 1) % 4], border);
                    }
                }
                if (rule.StrokeWidth > 0)
                {
                    border.Color = SKColor.TryParse("#" + rule.Color, out var ruleColor) ? ruleColor : SKColors.Gray;
                    border.StrokeWidth = (float)rule.StrokeWidth;
                    canvas.DrawRect(rect, border);
                }
            }
            foreach (var image in page.Images)
            {
                var media = layout.Media[image.PartName];
                using var data = SKData.CreateCopy(System.Convert.FromBase64String(media[(media.IndexOf(',') + 1)..]));
                using var bitmap = SKImage.FromEncodedData(data);
                if (bitmap is not null)
                {
                    canvas.Save();
                    canvas.Translate((float)(image.X + image.Width / 2), (float)(image.Y + image.Height / 2));
                    canvas.RotateDegrees((float)image.Rotation);
                    canvas.Scale(image.FlipH ? -1 : 1, image.FlipV ? -1 : 1);
                    canvas.Translate((float)(-image.Width / 2), (float)(-image.Height / 2));
                    canvas.ClipRect(SKRect.Create((float)image.Width, (float)image.Height));
                    var expandedWidth = image.Width / (1 - image.CropLeft - image.CropRight);
                    var expandedHeight = image.Height / (1 - image.CropTop - image.CropBottom);
                    canvas.DrawImage(bitmap, SKRect.Create((float)(-image.CropLeft * expandedWidth), (float)(-image.CropTop * expandedHeight), (float)expandedWidth, (float)expandedHeight), new SKSamplingOptions(SKFilterMode.Linear));
                    canvas.Restore();
                }
            }
            foreach (var run in page.Runs)
            {
                if (run.Background is { Length: 6 } background && SKColor.TryParse("#" + background, out var backgroundColor))
                {
                    using var shading = new SKPaint { Color = backgroundColor };
                    canvas.DrawRect(SKRect.Create((float)run.X, (float)run.Y, (float)run.Width, (float)run.Height), shading);
                }
                using var font = engine.Font(run.FontFamily, run.Bold, run.Italic, run.FontSize);
                using var paint = new SKPaint { Color = SKColor.Parse("#" + run.Color), IsAntialias = true };
                canvas.DrawText(run.Text, (float)run.X, (float)run.Baseline, SKTextAlign.Left, font, paint);
                if (run.Underline)
                {
                    paint.StrokeWidth = Math.Max(.5f, (float)run.FontSize / 16);
                    canvas.DrawLine((float)run.X, (float)run.Baseline + 1, (float)(run.X + run.Width), (float)run.Baseline + 1, paint);
                }
            }
            pdf.EndPage();
        }
        pdf.Close();
        return new(output.ToArray(), pages.Length, layout.Warnings);
    }
}
