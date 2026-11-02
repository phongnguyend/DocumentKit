using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private static IEnumerable<XElement> Blocks(XElement container)
    {
        foreach (var child in container.Elements())
        {
            if (child.Name == W + "sdt")
            {
                if (child.Element(W + "sdtContent") is { } content)
                {
                    foreach (var block in Blocks(content))
                    {
                        yield return block;
                    }
                }
            }
            else
            {
                yield return child;
            }
        }
    }

    private void Table(XElement table, double? containerX = null, double? containerWidth = null, int depth = 0, int? rowLimit = null)
    {
        if (depth > 16)
        {
            throw new NotSupportedException("Nested tables exceed 16 levels.");
        }
        var origin = containerX ?? FlowLeft;
        var available = containerWidth ?? (columnWidths.Length > 1 ? columnWidths.Min() : FlowWidth);
        double ShiftX() => containerX.HasValue ? 0 : FlowLeft - origin;
        if (columnWidths.Distinct().Count() > 1)
        {
            warnings.Add("Tables in unequal columns use the narrowest column width.");
        }
        var grid = table.Element(W + "tblGrid")?.Elements(W + "gridCol").Select(item => Math.Max(1, Number(item, "w", 1000) / 20)).ToArray() ?? [];
        if (grid.Length == 0)
        {
            var columns = table.Elements(W + "tr").Select(row => row.Elements(W + "tc").Sum(cell => Math.Max(1, (int)Number(cell.Element(W + "tcPr")?.Element(W + "gridSpan"), "val", 1)))).DefaultIfEmpty(1).Max();
            grid = Enumerable.Repeat(available / Math.Max(1, columns), Math.Max(1, columns)).ToArray();
        }
        var properties = table.Element(W + "tblPr");
        var gridWidth = grid.Sum();
        var widthElement = properties?.Element(W + "tblW");
        var preferredWidth = Value(widthElement, "type") switch
        {
            "pct" => available * Number(widthElement, "w", 5000) / 5000,
            "dxa" => Number(widthElement, "w", gridWidth * 20) / 20,
            _ => gridWidth
        };
        var width = Math.Min(available, preferredWidth > 0 ? preferredWidth : available);
        grid = grid.Select(value => value * width / gridWidth).ToArray();
        var tableX = origin + Number(properties?.Element(W + "tblInd"), "w") / 20;
        var tableStyleName = Value(properties?.Element(W + "tblStyle"));
        var tableStyle = tableStyleName is not null && styles.TryGetValue(tableStyleName, out var foundStyle) ? foundStyle : null;
        var merges = new Dictionary<int, (DocxPage Page, int Rule)>();
        var rows = table.Elements(W + "tr").Take(rowLimit ?? int.MaxValue).ToArray();
        var flowingMerges = new Dictionary<int, (double RemainingHeight, int RemainingRows)>();
        var header = new DocxPage(0, "table-header", "", [], [], [], []);
        var headerHeight = 0d;
        var collectingHeader = true;
        void NextTableRegion()
        {
            AdvanceFlow();
            if (headerHeight <= 0 || headerHeight >= (result.Height - top - bottom) / 2)
            {
                return;
            }
            var shiftX = ShiftX();
            page.Runs.AddRange(header.Runs.Select(run => run with { X = run.X + shiftX, Y = run.Y + y, Baseline = run.Baseline + y }));
            page.Rules.AddRange(header.Rules.Select(rule => rule with { X = rule.X + shiftX, Y = rule.Y + y }));
            page.Images.AddRange(header.Images.Select(image => image with { X = image.X + shiftX, Y = image.Y + y }));
            page.TextObjects.AddRange(header.TextObjects.Select(item => item with { Bounds = item.Bounds with { X = item.Bounds.X + shiftX, Y = item.Bounds.Y + y, Grouped = true } }));
            y += headerHeight;
        }
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var row = rows[rowIndex];
            var cells = new List<(DocxPage Content, double Height, double X, double Width, XElement Properties, int Column, XElement Borders, List<TableParagraphLayout> Paragraphs)>();
            var column = (int)Number(row.Element(W + "trPr")?.Element(W + "gridBefore"), "val");
            foreach (var cell in row.Elements(W + "tc"))
            {
                var cellFormat = TableFormatting(tableStyle, properties, rowIndex, rows.Length, column, grid.Length);
                var effectiveBorders = new XElement(cellFormat.Element(W + "tblPr")?.Element(W + "tblBorders") ?? new XElement(W + "tblBorders"));
                foreach (var edge in properties?.Element(W + "tblBorders")?.Elements() ?? [])
                {
                    MergeProperty(effectiveBorders, edge);
                }
                var cellProperties = new XElement(cellFormat.Element(W + "tcPr") ?? new XElement(W + "tcPr"));
                foreach (var property in cell.Element(W + "tcPr")?.Elements() ?? [])
                {
                    MergeProperty(cellProperties, property);
                }
                var span = Math.Clamp((int)Number(cellProperties.Element(W + "gridSpan"), "val", 1), 1, grid.Length);
                var cellX = tableX + grid.Take(column).Sum();
                var cellWidth = Math.Max(12, grid.Skip(column).Take(span).Sum());
                var previousPage = page;
                var previousY = y;
                var previousMeasuring = measuringTable;
                var previousStyle = previousParagraphStyle;
                var previousContext = previousContextual;
                previousParagraphStyle = null;
                var previousSpacing = pendingAfter;
                pendingAfter = 0;
                var margins = cellProperties.Element(W + "tcMar") ?? properties?.Element(W + "tblCellMar") ?? cellFormat.Element(W + "tblPr")?.Element(W + "tblCellMar");
                var padTop = Number(margins?.Element(W + "top"), "w", 60) / 20;
                var padBottom = Number(margins?.Element(W + "bottom"), "w", 60) / 20;
                var padLeft = Number(margins?.Element(W + "left"), "w", 80) / 20;
                var padRight = Number(margins?.Element(W + "right"), "w", 80) / 20;
                var previousFormatting = tableFormatting;
                tableFormatting = cellFormat;
                var scratch = new DocxPage(0, "measurement", "", [], [], [], []);
                result.Pages.Add(scratch);
                page = scratch;
                y = padTop;
                measuringTable = true;
                var paragraphs = new List<TableParagraphLayout>();
                try
                {
                    foreach (var block in Blocks(cell))
                    {
                        if (block.Name == W + "p")
                        {
                            var savedLines = measuredLines;
                            measuredLines = new();
                            try
                            {
                                Paragraph(block, cellX + padLeft, Math.Max(12, cellWidth - padLeft - padRight), true);
                                paragraphs.Add(new(Properties(block, "pPr"), measuredLines));
                            }
                            finally
                            {
                                measuredLines = savedLines;
                            }
                        }
                        else if (block.Name == W + "tbl")
                        {
                            Table(block, cellX + padLeft, Math.Max(12, cellWidth - padLeft - padRight), depth + 1);
                        }
                    }
                    cells.Add((scratch, y + padBottom, cellX, cellWidth, cellProperties, column, effectiveBorders, paragraphs));
                }
                finally
                {
                    result.Pages.Remove(scratch);
                    page = previousPage;
                    y = previousY;
                    measuringTable = previousMeasuring;
                    tableFormatting = previousFormatting;
                    pendingAfter = previousSpacing;
                    previousParagraphStyle = previousStyle;
                    previousContextual = previousContext;
                }
                column += span;
            }
            var mergedStarts = new HashSet<int>();
            foreach (var cell in cells)
            {
                var merge = cell.Properties.Element(W + "vMerge");
                if (merge is null || Value(merge) == "restart")
                {
                    flowingMerges.Remove(cell.Column);
                }
                var span = Math.Max(1, (int)Number(cell.Properties.Element(W + "gridSpan"), "val", 1));
                var mergeRows = Value(merge) == "restart" ? CountMergedRows(rows, rowIndex, cell.Column, span) : 1;
                // A top-aligned text cell can paint across its row boundaries when
                // all its text fits on this page. More complex merges retain the
                // conservative per-row layout rather than losing clipped content.
                if (Value(merge) == "restart" && mergeRows > 1
                    && Value(cell.Properties.Element(W + "vAlign")) is not ("center" or "bottom")
                    && cell.Content.Images.Count == 0 && cell.Content.Rules.Count == 0
                    && (measuringTable || y + cell.Height <= FlowLimit && cell.Height <= FlowLimit - columnTop - headerHeight)
                    && !Enabled(row.Element(W + "trPr") ?? new XElement(W + "trPr"), "tblHeader"))
                {
                    flowingMerges[cell.Column] = (cell.Height, mergeRows);
                    mergedStarts.Add(cell.Column);
                }
            }
            var rowHeight = Math.Max(cells.Where(cell => !flowingMerges.ContainsKey(cell.Column)).Select(cell => cell.Height).DefaultIfEmpty(0).Max(), Number(row.Element(W + "trPr")?.Element(W + "trHeight"), "val") / 20);
            foreach (var merge in flowingMerges.Values)
            {
                if (merge.RemainingRows == 1)
                {
                    rowHeight = Math.Max(rowHeight, merge.RemainingHeight);
                }
            }
            rowHeight = Math.Max(.1, rowHeight);
            var pageCapacity = result.Height - top - bottom;
            var rowProperties = row.Element(W + "trPr") ?? new XElement(W + "trPr");
            var cannotSplit = Enabled(rowProperties, "cantSplit") || Enabled(rowProperties, "tblHeader");
            var minimumHeight = Number(row.Element(W + "trPr")?.Element(W + "trHeight"), "val") / 20;
            if (!measuringTable && y > columnTop
                && ((cannotSplit && rowHeight <= pageCapacity && y + rowHeight > FlowLimit) || y + minimumHeight > FlowLimit))
            {
                NextTableRegion();
                merges.Clear();
            }
            var headerRow = collectingHeader && Enabled(row.Element(W + "trPr") ?? new XElement(W + "trPr"), "tblHeader");
            collectingHeader = headerRow;
            var rowStart = y;
            var rowPage = page;
            var startRun = page.Runs.Count;
            var startRule = page.Rules.Count;
            var startImage = page.Images.Count;
            var startObject = page.TextObjects.Count;
            double consumed = 0;
            while (consumed < rowHeight)
            {
                var height = measuringTable ? rowHeight - consumed : Math.Min(rowHeight - consumed, FlowLimit - y);
                if (!measuringTable && consumed + height < rowHeight)
                {
                    var independentCells = flowingMerges.Count == 0 && cells.All(cell => cell.Content.Images.Count == 0 && cell.Content.Rules.Count == 0
                        && Value(cell.Properties.Element(W + "vAlign")) is not ("center" or "bottom"));
                    if (independentCells)
                    {
                        var boundaries = cells.Select(cell => TableCellBoundary(cell.Paragraphs, consumed, consumed + height, FlowLimit - columnTop - headerHeight)).ToArray();
                        if (boundaries.All(boundary => boundary <= consumed) && y > columnTop + headerHeight + .01)
                        {
                            height = 0;
                        }
                        else
                        {
                            for (var cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                            {
                                var cell = cells[cellIndex];
                                var shift = consumed + height - boundaries[cellIndex];
                                if (shift > .01)
                                {
                                    if (boundaries[cellIndex] <= consumed && y <= columnTop + headerHeight + .01)
                                    {
                                        throw new NotSupportedException("A table paragraph cannot fit below the repeated header in one page region.");
                                    }
                                    ShiftTableCell(cell.Content, cell.Paragraphs, boundaries[cellIndex], shift);
                                    cell.Height += shift;
                                    cells[cellIndex] = cell;
                                    rowHeight = Math.Max(rowHeight, cell.Height);
                                }
                            }
                        }
                    }
                    else
                    {
                        var boundary = consumed + height;
                        foreach (var cell in cells)
                        {
                            var padding = Value(cell.Properties.Element(W + "vAlign")) switch
                            {
                                "center" => Math.Max(0, (rowHeight - cell.Height) / 2),
                                "bottom" => Math.Max(0, rowHeight - cell.Height),
                                _ => 0
                            };
                            foreach (var run in cell.Content.Runs)
                            {
                                if (run.Y + padding < boundary && run.Y + padding + run.Height > boundary)
                                {
                                    boundary = Math.Max(consumed, run.Y + padding);
                                }
                            }
                            foreach (var image in cell.Content.Images)
                            {
                                if (image.Y < boundary && image.Y + image.Height > boundary && image.Height <= pageCapacity)
                                {
                                    boundary = Math.Max(consumed, image.Y);
                                }
                            }
                        }
                        if (boundary <= consumed && y <= columnTop + headerHeight + .01)
                        {
                            throw new NotSupportedException("A table line or image cannot fit below the repeated header in one page region.");
                        }
                        if (boundary > consumed || y > columnTop)
                        {
                            height = boundary - consumed;
                        }
                    }
                }
                if (height <= 0)
                {
                    NextTableRegion();
                    merges.Clear();
                    continue;
                }
                foreach (var cell in cells)
                {
                    var padding = Value(cell.Properties.Element(W + "vAlign")) switch
                    {
                        "center" => Math.Max(0, (rowHeight - cell.Height) / 2),
                        "bottom" => Math.Max(0, rowHeight - cell.Height),
                        _ => 0
                    };
                    var runs = cell.Content.Runs.Where(run => mergedStarts.Contains(cell.Column)
                        ? consumed == 0
                        : !flowingMerges.ContainsKey(cell.Column) && run.Y + padding >= consumed && run.Y + padding < consumed + height)
                        .Select(run => run with { X = run.X + ShiftX(), Y = run.Y + padding - consumed + y, Baseline = run.Baseline + padding - consumed + y }).ToArray();
                    page.Runs.AddRange(runs);
                    foreach (var group in runs.GroupBy(run => run.Id))
                    {
                        var original = cell.Content.TextObjects.First(item => item.Id == group.Key);
                        var minY = group.Min(item => item.Y);
                        page.TextObjects.Add(original with { Bounds = original.Bounds with { X = original.Bounds.X + ShiftX(), Y = minY, Height = group.Max(item => item.Y + item.Height) - minY, Grouped = true } });
                    }
                    page.Images.AddRange(cell.Content.Images.Where(image => image.Y >= consumed && image.Y < consumed + height).Select(image => image with { X = image.X + ShiftX(), Y = image.Y - consumed + y, Grouped = true }));
                    var fill = Value(cell.Properties.Element(W + "shd"), "fill");
                    var cellBorders = cell.Properties.Element(W + "tcBorders");
                    DocxBorder? Edge(string side, string fallback)
                    {
                        var edge = cellBorders?.Element(W + side) ?? cell.Borders.Element(W + fallback);
                        if (Value(edge) is null or "nil" or "none")
                        {
                            return null;
                        }
                        var color = Value(edge, "color");
                        return new(Math.Max(0, Number(edge, "sz", 4) / 8), color is null or "auto" ? "000000" : color);
                    }
                    var span = Math.Max(1, (int)Number(cell.Properties.Element(W + "gridSpan"), "val", 1));
                    DocxBorder?[] edges = [
                        Edge("top", rowIndex == 0 ? "top" : "insideH"),
                        Edge("right", cell.Column + span >= grid.Length ? "right" : "insideV"),
                        Edge("bottom", rowIndex == rows.Length - 1 ? "bottom" : "insideH"),
                        Edge("left", cell.Column == 0 ? "left" : "insideV")
                    ];
                    var merge = cell.Properties.Element(W + "vMerge");
                    if (merge is not null && Value(merge) != "restart" && merges.TryGetValue(cell.Column, out var active) && active.Page == page)
                    {
                        var existing = page.Rules[active.Rule];
                        page.Rules[active.Rule] = existing with { Height = y + height - existing.Y, Edges = [existing.Edges?[0], existing.Edges?[1], edges[2], existing.Edges?[3]] };
                    }
                    else
                    {
                        page.Rules.Add(new(cell.X + ShiftX(), y, cell.Width, height, fill, 0, Edges: edges));
                        if (merge is not null)
                        {
                            merges[cell.Column] = (page, page.Rules.Count - 1);
                        }
                        else
                        {
                            merges.Remove(cell.Column);
                        }
                    }
                    // Parent backgrounds precede nested borders/fills in paint order.
                    page.Rules.AddRange(cell.Content.Rules.Where(rule => rule.Y < consumed + height && rule.Y + rule.Height > consumed)
                        .Select(rule => rule with { X = rule.X + ShiftX(), Y = Math.Max(rule.Y, consumed) - consumed + y, Height = Math.Min(rule.Y + rule.Height, consumed + height) - Math.Max(rule.Y, consumed) }));
                }
                consumed += height;
                y += height;
                flowBottom = Math.Max(flowBottom, y);
                if (consumed < rowHeight)
                {
                    NextTableRegion();
                    merges.Clear();
                }
            }
            if (headerRow && page == rowPage)
            {
                var shift = headerHeight - rowStart;
                header.Runs.AddRange(page.Runs.Skip(startRun).Select(run => run with { Y = run.Y + shift, Baseline = run.Baseline + shift }));
                header.Rules.AddRange(page.Rules.Skip(startRule).Select(rule => rule with { Y = rule.Y + shift }));
                header.Images.AddRange(page.Images.Skip(startImage).Select(image => image with { Y = image.Y + shift }));
                header.TextObjects.AddRange(page.TextObjects.Skip(startObject).Select(item => item with { Bounds = item.Bounds with { Y = item.Bounds.Y + shift } }));
                headerHeight += rowHeight;
            }
            foreach (var columnIndex in flowingMerges.Keys.ToArray())
            {
                var merge = flowingMerges[columnIndex];
                if (merge.RemainingRows <= 1)
                {
                    flowingMerges.Remove(columnIndex);
                }
                else
                {
                    flowingMerges[columnIndex] = (Math.Max(0, merge.RemainingHeight - rowHeight), merge.RemainingRows - 1);
                }
            }
        }
    }

    private static int CountMergedRows(XElement[] rows, int start, int column, int span)
    {
        var count = 1;
        for (var index = start + 1; index < rows.Length; index++)
        {
            var currentColumn = (int)Number(rows[index].Element(W + "trPr")?.Element(W + "gridBefore"), "val");
            XElement? continuation = null;
            foreach (var cell in rows[index].Elements(W + "tc"))
            {
                var properties = cell.Element(W + "tcPr");
                var currentSpan = Math.Max(1, (int)Number(properties?.Element(W + "gridSpan"), "val", 1));
                if (currentColumn == column && currentSpan == span)
                {
                    continuation = properties?.Element(W + "vMerge");
                    break;
                }
                currentColumn += currentSpan;
            }
            if (continuation is null || Value(continuation) == "restart")
            {
                break;
            }
            count++;
        }
        return count;
    }
}
