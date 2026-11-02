using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private sealed record TableParagraphLayout(XElement Properties, List<(double Top, double Bottom)> Lines);

    private static double TableCellBoundary(List<TableParagraphLayout> paragraphs, double start, double end, double capacity)
    {
        var boundary = end;
        foreach (var paragraph in paragraphs)
        {
            var lines = paragraph.Lines.Where(line => line.Bottom > start + .01).ToArray();
            if (lines.Length == 0 || lines[0].Top >= boundary || lines[^1].Bottom <= boundary)
            {
                continue;
            }
            var fitting = lines.TakeWhile(line => line.Bottom <= boundary).Count();
            boundary = Math.Min(boundary, lines[fitting].Top);
            if (Enabled(paragraph.Properties, "keepLines") && lines[0].Top >= start
                && lines[^1].Bottom - lines[0].Top <= capacity)
            {
                boundary = Math.Min(boundary, lines[0].Top);
            }
            else if (paragraph.Properties.Element(W + "widowControl") is null || Enabled(paragraph.Properties, "widowControl"))
            {
                if (fitting == 1 && lines.Length > 1 && lines[1].Bottom - lines[0].Top <= capacity)
                {
                    boundary = Math.Min(boundary, lines[0].Top);
                }
                else if (fitting >= 2 && lines.Length - fitting == 1)
                {
                    var retained = fitting == 2 ? 0 : fitting - 1;
                    if (lines[^1].Bottom - lines[retained].Top <= capacity)
                    {
                        boundary = Math.Min(boundary, lines[retained].Top);
                    }
                }
            }
        }
        return Math.Max(start, boundary);
    }

    private static void ShiftTableCell(DocxPage content, List<TableParagraphLayout> paragraphs, double boundary, double shift)
    {
        for (var index = 0; index < content.Runs.Count; index++)
        {
            var run = content.Runs[index];
            if (run.Y >= boundary - .01)
            {
                content.Runs[index] = run with { Y = run.Y + shift, Baseline = run.Baseline + shift };
            }
        }
        foreach (var paragraph in paragraphs)
        {
            for (var index = 0; index < paragraph.Lines.Count; index++)
            {
                var line = paragraph.Lines[index];
                if (line.Top >= boundary - .01)
                {
                    paragraph.Lines[index] = (line.Top + shift, line.Bottom + shift);
                }
            }
        }
    }
}
