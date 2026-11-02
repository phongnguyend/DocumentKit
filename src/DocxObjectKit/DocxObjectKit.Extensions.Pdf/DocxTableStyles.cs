using System.Xml.Linq;
using System.Globalization;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private XElement? tableFormatting;

    private XElement TableFormatting(XElement? style, XElement? properties, int row, int rowCount, int column, int columnCount)
    {
        var chain = new List<XElement>();
        var seen = new HashSet<string>();
        while (style is not null && seen.Add(Value(style, "styleId") ?? ""))
        {
            chain.Add(style);
            var parent = Value(style.Element(W + "basedOn"));
            style = parent is not null && styles.TryGetValue(parent, out var inherited) ? inherited : null;
        }
        chain.Reverse();
        var merged = new XElement("format");
        void Apply(XElement item)
        {
            foreach (var property in item.Elements().Where(item => item.Name.LocalName is "pPr" or "rPr" or "tcPr" or "tblPr"))
            {
                MergeProperty(merged, property);
            }
        }
        foreach (var item in chain)
        {
            Apply(item);
        }
        var look = properties?.Element(W + "tblLook");
        _ = int.TryParse(Value(look), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var flags);
        bool Flag(string name, int mask)
        {
            var explicitValue = Value(look, name);
            return explicitValue is null ? (flags & mask) != 0 : explicitValue is not ("0" or "false" or "off");
        }
        var conditions = new List<string> { "wholeTable" };
        if (!Flag("noHBand", 0x200))
        {
            var bandRow = row - (Flag("firstRow", 0x20) ? 1 : 0);
            conditions.Add(bandRow % 2 == 0 ? "band1Horz" : "band2Horz");
        }
        if (!Flag("noVBand", 0x400))
        {
            conditions.Add(column % 2 == 0 ? "band1Vert" : "band2Vert");
        }
        if (column == 0 && Flag("firstColumn", 0x80))
        {
            conditions.Add("firstCol");
        }
        if (column == columnCount - 1 && Flag("lastColumn", 0x100))
        {
            conditions.Add("lastCol");
        }
        if (row == 0 && Flag("firstRow", 0x20))
        {
            conditions.Add("firstRow");
        }
        if (row == rowCount - 1 && Flag("lastRow", 0x40))
        {
            conditions.Add("lastRow");
        }
        foreach (var condition in conditions)
        {
            foreach (var item in chain.SelectMany(item => item.Elements(W + "tblStylePr")).Where(item => Value(item, "type") == condition))
            {
                Apply(item);
            }
        }
        return merged;
    }
}
