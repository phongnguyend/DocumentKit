using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private int? widowBreakAfter;
    private List<(double Top, double Bottom)>? measuredLines;

    private static double LineAdvance(SkiaSharp.SKFont font, XElement? spacing, double size)
    {
        return Value(spacing, "lineRule") switch
        {
            "exact" => Math.Max(.1, Number(spacing, "line", size * 24) / 20),
            "atLeast" => Math.Max(font.Spacing, Number(spacing, "line") / 20),
            _ => font.Spacing * Math.Max(.01, Number(spacing, "line", 240) / 240)
        };
    }

    private bool CanMeasureParagraph(XElement paragraph)
    {
        return paragraph.Name == W + "p"
            && paragraph.Element(W + "pPr")?.Element(W + "sectPr") is null
            && Properties(paragraph, "pPr").Element(W + "framePr") is null
            && !paragraph.Descendants().Any(item => item.Name == W + "drawing" || item.Name == W + "pict" || item.Name == W + "object"
                || item.Name == W + "br" && Value(item, "type") is "page" or "column");
    }

    private void KeepParagraphsTogether(XElement[] blocks, int index)
    {
        if (measuringTable || !CanMeasureParagraph(blocks[index]))
        {
            return;
        }
        var first = Properties(blocks[index], "pPr");
        if (Enabled(first, "pageBreakBefore"))
        {
            return;
        }
        var end = index;
        while (end + 1 < blocks.Length && end - index < 255 && Enabled(Properties(blocks[end], "pPr"), "keepNext")
            && CanMeasureParagraph(blocks[end + 1]) && !Enabled(Properties(blocks[end + 1], "pPr"), "pageBreakBefore"))
        {
            end++;
        }
        var followingTable = end + 1 < blocks.Length && blocks[end + 1].Name == W + "tbl" && Enabled(Properties(blocks[end], "pPr"), "keepNext") ? blocks[end + 1] : null;
        var savedFlowBottom = flowBottom;
        var savedPages = result.Pages.ToArray();
        var savedPage = page;
        var savedY = y;
        var savedAfter = pendingAfter;
        var savedStyle = previousParagraphStyle;
        var savedContextual = previousContextual;
        var savedMeasuring = measuringTable;
        var savedCounters = counters.ToDictionary(item => item.Key, item => (int[])item.Value.Clone());
        var savedInitialized = initializedCounters.ToDictionary(item => item.Key, item => (bool[])item.Value.Clone());
        var savedLines = measuredLines;
        var required = 0d;
        var firstLines = Array.Empty<(double Top, double Bottom)>();
        try
        {
            page = new(0, "pagination-measurement", "", [], [], [], []);
            result.Pages.Clear();
            result.Pages.Add(page);
            measuringTable = true;
            for (var paragraphIndex = index; paragraphIndex <= end; paragraphIndex++)
            {
                var beforeRun = page.Runs.Count;
                measuredLines = paragraphIndex == index ? new() : null;
                Paragraph(blocks[paragraphIndex], FlowLeft, FlowWidth);
                required = y - pendingAfter - savedY;
                if (paragraphIndex == index)
                {
                    firstLines = measuredLines!.ToArray();
                    measuredLines = null;
                }
                // Keep-next binds to the first line, unless the following paragraph itself must stay whole.
                if (paragraphIndex == end && end > index && !Enabled(Properties(blocks[end], "pPr"), "keepLines"))
                {
                    var firstLine = page.Runs.Skip(beforeRun).FirstOrDefault();
                    if (firstLine is not null)
                    {
                        var nextLine = page.Runs.Skip(beforeRun).FirstOrDefault(run => run.Baseline > firstLine.Baseline + .01);
                        if (nextLine is not null)
                        {
                            required = firstLine.Y + nextLine.Baseline - firstLine.Baseline - savedY;
                        }
                    }
                }
            }
            if (followingTable is not null)
            {
                Table(followingTable, rowLimit: 1);
                required = y - savedY;
            }
        }
        finally
        {
            measuredLines = savedLines;
            result.Pages.Clear();
            result.Pages.AddRange(savedPages);
            flowBottom = savedFlowBottom;
            page = savedPage;
            y = savedY;
            pendingAfter = savedAfter;
            previousParagraphStyle = savedStyle;
            previousContextual = savedContextual;
            measuringTable = savedMeasuring;
            counters.Clear();
            initializedCounters.Clear();
            foreach (var item in savedCounters)
            {
                counters.Add(item.Key, item.Value);
            }
            foreach (var item in savedInitialized)
            {
                initializedCounters.Add(item.Key, item.Value);
            }
        }
        var capacity = FlowLimit - columnTop;
        var keep = Enabled(first, "keepLines") || end > index || followingTable is not null;
        var widow = first.Element(W + "widowControl") is null || Enabled(first, "widowControl");
        var fitting = firstLines.TakeWhile(line => line.Bottom <= FlowLimit).Count();
        if (widow && !keep && fitting >= 2 && firstLines.Length - fitting == 1 && firstLines.Length >= 4)
        {
            widowBreakAfter = fitting - 1;
        }
        var shortWidow = widow && firstLines.Length == 3 && fitting == 2 && firstLines[^1].Bottom - firstLines[0].Top <= capacity;
        var orphan = widow && fitting == 1 && firstLines.Length > 1 && firstLines[1].Bottom - firstLines[0].Top <= capacity;
        if (y > columnTop && ((keep && required <= capacity && y + required > FlowLimit) || orphan || shortWidow))
        {
            AdvanceFlow();
            pendingAfter = 0;
            previousParagraphStyle = null;
            previousContextual = false;
        }
    }
}
