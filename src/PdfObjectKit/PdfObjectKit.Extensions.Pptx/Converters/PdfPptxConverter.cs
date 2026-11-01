using PdfObjectKit.Graphs;
using PdfObjectKit.Interpreters;
using PdfObjectKit.Models;
using PdfObjectKit.Options;
using PdfObjectKit.Rendering;
using static PdfObjectKit.Primitives.PdfPrimitives;

namespace PdfObjectKit.Converters;

/// <summary>Exports editable text, pictures and supported paths; does not rasterize slides.</summary>
public static class PdfPptxConverter
{
    public static PdfPptxResult Convert(PdfObjectDocument document, PdfPptxOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(document);
        options ??= new();
        if (document.Security is not null)
        {
            throw new NotSupportedException("Unlock the PDF before exporting PowerPoint slides.");
        }
        if (options.SlideWidth.HasValue != options.SlideHeight.HasValue
            || options.SlideWidth is { } sw && (!double.IsFinite(sw) || sw < 72 || sw > 4032)
            || options.SlideHeight is { } sh && (!double.IsFinite(sh) || sh < 72 || sh > 4032))
        {
            throw new ArgumentException("Specify both slide dimensions, each between 72 and 4032 points.");
        }
        var graph = new PdfGraph(document);
        var pages = graph.Pages();
        var selected = options.PageNumbers?.ToArray() ?? Enumerable.Range(1, pages.Count).ToArray();
        if (selected.Length is < 1 or > 100 || selected.Any(number => number < 1 || number > pages.Count) || selected.Distinct().Count() != selected.Length)
        {
            throw new ArgumentException("Choose 1-100 distinct, valid page numbers.");
        }
        var warnings = new HashSet<string>
        {
            "Editable conversion is approximate. Fonts are referenced by name, not embedded; install matching fonts to reduce substitution. Text wrapping and spacing may differ."
        };
        var geometry = selected.Select(number => Geometry(graph, pages[number - 1].Dictionary)).ToArray();
        var defaultScale = Math.Min(1, 4032 / Math.Max(geometry[0].Width, geometry[0].Height));
        var width = options.SlideWidth ?? Math.Max(72, geometry[0].Width * defaultScale);
        var height = options.SlideHeight ?? Math.Max(72, geometry[0].Height * defaultScale);
        if (geometry.Any(page => Math.Abs(page.Width - width) > .01 || Math.Abs(page.Height - height) > .01))
        {
            warnings.Add("Pages are fitted proportionally and centered within one common slide size.");
        }
        using var package = new PptxPackage();
        var textCount = 0;
        var imageCount = 0;
        var shapeCount = 0;
        for (var i = 0; i < selected.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var page = geometry[i];
            var scale = Math.Min(width / page.Width, height / page.Height);
            var transform = new PdfMatrix(scale, 0, 0, scale, (width - page.Width * scale) / 2, (height - page.Height * scale) / 2).Multiply(page.Transform);
            var display = PdfCanvasExporter.Export(document, selected[i]);
            foreach (var warning in display.Warnings.Where(w => !w.StartsWith("Experimental:", StringComparison.Ordinal) && !w.Contains("uses a browser substitute", StringComparison.Ordinal)))
            {
                warnings.Add($"Page {selected[i]}: {warning}");
            }
            var writer = new PptxSlideWriter(package, transform, width, height, options, message => warnings.Add($"Page {selected[i]}: {message}"));
            writer.Write(display.Commands, cancellationToken);
            textCount += writer.TextCount;
            imageCount += writer.ImageCount;
            shapeCount += writer.ShapeCount;
            if (writer.TextCount == 0)
            {
                warnings.Add($"Page {selected[i]}: no editable visible text was found. Scanned text needs OCR before export.");
            }
        }
        warnings.Add("PDF links, comments, interactive fields, tags and original document metadata are not preserved as PowerPoint features. Supported field appearances become ordinary slide objects.");
        return new(package.Finish(width, height), selected.Length, textCount, imageCount, shapeCount, warnings.ToArray());
    }

    private static (double Width, double Height, PdfMatrix Transform) Geometry(PdfGraph graph, PdfValue page)
    {
        PdfValue? Inherit(string key)
        {
            var current = page;
            var visited = new HashSet<PdfValue>();
            while (visited.Add(current))
            {
                if (graph.Get(current, key) is { } value)
                {
                    return value;
                }
                if (graph.Get(current, "Parent") is not { } parent)
                {
                    return null;
                }
                current = parent;
            }
            throw new InvalidDataException("Cyclic page inheritance.");
        }
        var media = Inherit("MediaBox")?.Items?.Select(Number).ToArray() ?? [0, 0, 612, 792];
        var crop = Inherit("CropBox")?.Items?.Select(Number).ToArray() ?? media;
        if (media.Length != 4 || crop.Length != 4 || media.Concat(crop).Any(value => !double.IsFinite(value)))
        {
            throw new InvalidDataException("Invalid PDF page box.");
        }
        var left = Math.Max(media[0], crop[0]);
        var bottom = Math.Max(media[1], crop[1]);
        var right = Math.Min(media[2], crop[2]);
        var top = Math.Min(media[3], crop[3]);
        var rotate = Number(Inherit("Rotate") ?? Num(0));
        var unit = graph.Number(page, "UserUnit", 1);
        if (!double.IsFinite(rotate) || rotate % 90 != 0 || !double.IsFinite(unit) || unit <= 0 || unit > 75000 || right <= left || top <= bottom)
        {
            throw new InvalidDataException("Invalid page rotation or dimensions.");
        }
        var angle = (rotate % 360 + 360) % 360;
        var transform = angle switch
        {
            90 => new PdfMatrix(0, unit, unit, 0, -bottom * unit, -left * unit),
            180 => new PdfMatrix(-unit, 0, 0, unit, right * unit, -bottom * unit),
            270 => new PdfMatrix(0, -unit, -unit, 0, top * unit, right * unit),
            _ => new PdfMatrix(unit, 0, 0, -unit, -left * unit, top * unit)
        };
        return angle is 90 or 270 ? ((top - bottom) * unit, (right - left) * unit, transform) : ((right - left) * unit, (top - bottom) * unit, transform);
    }
}
