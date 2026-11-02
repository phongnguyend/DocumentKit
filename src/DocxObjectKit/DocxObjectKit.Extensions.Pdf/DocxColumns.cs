using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private double[] columnWidths = [];
    private double[] columnOffsets = [];
    private int columnIndex;
    private double columnTop;
    private double flowBottom;

    private double FlowLeft => left + (columnOffsets.Length == 0 ? 0 : columnOffsets[columnIndex]);

    private double FlowWidth => columnWidths.Length == 0 ? result.Width - left - right : columnWidths[columnIndex];

    private void ConfigureColumns(XElement? section)
    {
        var columns = section?.Element(W + "cols");
        var available = result.Width - left - right;
        var explicitColumns = columns?.Elements(W + "col").ToArray() ?? [];
        var count = Math.Clamp((int)Number(columns, "num", 1), 1, 16);
        if (Value(columns, "equalWidth") is "0" or "false" && explicitColumns.Length > 0)
        {
            count = Math.Min(16, explicitColumns.Length);
        }
        var gap = Math.Clamp(Number(columns, "space", 720) / 20, 0, available / count);
        var widths = new double[count];
        var spaces = new double[count];
        for (var index = 0; index < count; index++)
        {
            var explicitColumn = Value(columns, "equalWidth") is "0" or "false" ? explicitColumns.ElementAtOrDefault(index) : null;
            widths[index] = Math.Max(1, explicitColumn is null ? (available - gap * (count - 1)) / count : Number(explicitColumn, "w", available * 20 / count) / 20);
            spaces[index] = index == count - 1 ? 0 : Math.Max(0, explicitColumn is null ? gap : Number(explicitColumn, "space", gap * 20) / 20);
        }
        var total = widths.Sum() + spaces.Sum();
        var scale = total > available ? available / total : 1;
        columnWidths = widths.Select(width => width * scale).ToArray();
        columnOffsets = new double[count];
        for (var index = 1; index < count; index++)
        {
            columnOffsets[index] = columnOffsets[index - 1] + (widths[index - 1] + spaces[index - 1]) * scale;
        }
        columnIndex = 0;
        columnTop = top;
        if (count > 1)
        {
            warnings.Add("Columns flow sequentially; end-of-section column balancing and separator rules are not implemented.");
        }
    }

    private void AdvanceFlow()
    {
        if (columnIndex + 1 < columnWidths.Length)
        {
            columnIndex++;
            y = columnTop;
        }
        else
        {
            page = NewPage();
        }
    }
}
