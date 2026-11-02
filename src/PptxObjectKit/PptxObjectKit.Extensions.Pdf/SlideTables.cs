using System.Xml.Linq;
using SkiaSharp;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Extensions.Pdf;

internal sealed partial class SlideRenderer
{
    private void Table(SKCanvas canvas, XElement table, float width, float height)
    {
        var columns = D(C(table, "tblGrid"), "gridCol").Select(column => N(column, "w") / Emu).ToArray();
        var rows = table.Elements(A + "tr").ToArray();
        if (columns.Length == 0 || columns.Length > 1000 || rows.Length > 1000)
        {
            throw new NotSupportedException("Invalid or excessively large table.");
        }
        var heights = rows.Select(row => N(row, "h") / Emu).ToArray();
        var props = C(table, "tblPr");
        var styleId = C(props, "tableStyleId")?.Value;
        var stylesPart = Related(graph.PresentationPart(), "tableStyles");
        var styles = stylesPart is null ? null : document.Xml(stylesPart);
        var style = D(styles, "tblStyle").FirstOrDefault(item => (string?)item.Attribute("styleId") == styleId);
        if (styleId is not null && style is null)
        {
            Warn("A built-in table style is unavailable; explicit cell formatting is preserved.");
        }
        float y = 0;
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            float x = 0;
            var cells = rows[rowIndex].Elements(A + "tc").ToArray();
            for (var col = 0; col < Math.Min(cells.Length, columns.Length); col++)
            {
                var cell = cells[col];
                if (N(cell, "hMerge") != 0 || N(cell, "vMerge") != 0)
                {
                    x += columns[col];
                    continue;
                }
                var cellWidth = columns.Skip(col).Take(Math.Clamp((int)N(cell, "gridSpan", 1), 1, columns.Length - col)).Sum();
                var cellHeight = heights.Skip(rowIndex).Take(Math.Clamp((int)N(cell, "rowSpan", 1), 1, rows.Length - rowIndex)).Sum();
                var cellProps = new XElement(C(cell, "tcPr") ?? new XElement(A + "tcPr"));
                var sections = new List<XElement?>();
                if (rowIndex == 0 && N(props, "firstRow") != 0)
                {
                    sections.Add(C(style, "firstRow"));
                }
                if (rowIndex == rows.Length - 1 && N(props, "lastRow") != 0)
                {
                    sections.Add(C(style, "lastRow"));
                }
                if (col == 0 && N(props, "firstCol") != 0)
                {
                    sections.Add(C(style, "firstCol"));
                }
                if (N(props, "bandRow") != 0)
                {
                    sections.Add(C(style, rowIndex % 2 == 0 ? "band1H" : "band2H"));
                }
                sections.Add(C(style, "wholeTbl"));
                foreach (var section in sections)
                {
                    Merge(cellProps, C(section, "tcStyle"));
                    if (C(C(section, "tcStyle"), "fill") is { } fill)
                    {
                        Merge(cellProps, fill);
                    }
                }
                canvas.Save();
                canvas.Translate(x, y);
                using (var fill = Fill(cellProps, cellWidth, cellHeight))
                {
                    if (fill is not null)
                    {
                        canvas.DrawRect(0, 0, cellWidth, cellHeight, fill);
                    }
                }
                foreach (var (side, start, end) in new[] { ("lnL", new SKPoint(0, 0), new SKPoint(0, cellHeight)), ("lnR", new SKPoint(cellWidth, 0), new SKPoint(cellWidth, cellHeight)), ("lnT", new SKPoint(0, 0), new SKPoint(cellWidth, 0)), ("lnB", new SKPoint(0, cellHeight), new SKPoint(cellWidth, cellHeight)) })
                {
                    var border = C(cellProps, side);
                    using var paint = Fill(border, cellWidth, cellHeight);
                    if (paint is not null)
                    {
                        paint.StrokeWidth = N(border, "w", Emu) / Emu;
                        canvas.DrawLine(start, end, paint);
                    }
                }
                if (C(cell, "txBody") is { } source)
                {
                    var body = new XElement(source);
                    var bodyPr = C(body, "bodyPr")!;
                    foreach (var (from, to) in new[] { ("marL", "lIns"), ("marR", "rIns"), ("marT", "tIns"), ("marB", "bIns"), ("anchor", "anchor") })
                    {
                        if (cellProps.Attribute(from) is { } value)
                        {
                            bodyPr.SetAttributeValue(to, value.Value);
                        }
                    }
                    foreach (var run in D(body, "rPr"))
                    {
                        foreach (var section in sections)
                        {
                            Merge(run, C(section, "tcTxStyle"));
                        }
                    }
                    Text(canvas, body, cellWidth, cellHeight);
                }
                canvas.Restore();
                x += columns[col];
            }
            y += heights[rowIndex];
        }
    }
}
