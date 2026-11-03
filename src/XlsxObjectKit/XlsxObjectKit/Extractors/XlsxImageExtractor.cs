using System.Xml.Linq;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Extractors;

public static class XlsxImageExtractor
{
    internal static readonly XNamespace D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    internal static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    internal static string RelationsPart(string part) => part[..(part.LastIndexOf('/') + 1)] + "_rels/" + part[(part.LastIndexOf('/') + 1)..] + ".rels";

    internal static XDocument Relations(XlsxObjectDocument document, string part) => document.PartNames.Contains(RelationsPart(part))
        ? document.ReadXml(RelationsPart(part)) : new(new XElement(Rel + "Relationships"));

    internal static string? Target(XlsxObjectDocument document, string part, string? id)
    {
        var relation = Relations(document, part).Root!.Elements().FirstOrDefault(item => (string?)item.Attribute("Id") == id);
        return relation is null || (string?)relation.Attribute("TargetMode") == "External" ? null : Resolve(part, (string)relation.Attribute("Target")!);
    }

    public static double ColumnPosition(XlsxSheet sheet, int column) => (column - 1) * (sheet.DefaultColumnWidth * 7 + 5)
        - sheet.HiddenColumns.Where(index => index < column).Sum(index => sheet.ColumnWidths.GetValueOrDefault(index, sheet.DefaultColumnWidth) * 7 + 5)
        + sheet.ColumnWidths.Where(item => item.Key < column).Sum(item => (item.Value - sheet.DefaultColumnWidth) * 7);

    public static double RowPosition(XlsxSheet sheet, int row) => ((row - 1) * sheet.DefaultRowHeight
        - sheet.HiddenRows.Where(index => index < row).Sum(index => sheet.RowHeights.GetValueOrDefault(index, sheet.DefaultRowHeight))
        + sheet.RowHeights.Where(item => item.Key < row).Sum(item => item.Value - sheet.DefaultRowHeight)) * 4 / 3;

    public static IReadOnlyList<XlsxImage> Extract(XlsxObjectDocument document, XlsxSheet sheet)
    {
        var result = new List<XlsxImage>();
        var reference = document.ReadXml(sheet.Part).Root!.Element(S + "drawing");
        var part = Target(document, sheet.Part, (string?)reference?.Attribute(R + "id"));
        if (part is null)
        {
            return result;
        }
        foreach (var anchor in document.ReadXml(part).Root!.Elements())
        {
            foreach (var picture in anchor.Elements(D + "pic"))
            {
                var properties = picture.Element(D + "nvPicPr")?.Element(D + "cNvPr");
                var imagePart = Target(document, part, (string?)picture.Descendants(A + "blip").FirstOrDefault()?.Attribute(R + "embed"));
                if (imagePart is null)
                {
                    continue;
                }
                var bytes = document.GetPart(imagePart);
                var mime = Mime(bytes);
                if (mime is null)
                {
                    continue;
                }
                var from = anchor.Element(D + "from");
                var to = anchor.Element(D + "to");
                var position = anchor.Element(D + "pos");
                var extent = anchor.Element(D + "ext");
                double Marker(XElement marker, bool horizontal) => (horizontal ? ColumnPosition(sheet, (int?)marker.Element(D + "col") + 1 ?? 1)
                    : RowPosition(sheet, (int?)marker.Element(D + "row") + 1 ?? 1)) + ((double?)marker.Element(D + (horizontal ? "colOff" : "rowOff")) ?? 0) / 9525;
                var x = from is null ? ((double?)position?.Attribute("x") ?? 0) / 9525 : Marker(from, true);
                var y = from is null ? ((double?)position?.Attribute("y") ?? 0) / 9525 : Marker(from, false);
                var width = to is null ? ((double?)extent?.Attribute("cx") ?? 0) / 9525 : Marker(to, true) - x;
                var height = to is null ? ((double?)extent?.Attribute("cy") ?? 0) / 9525 : Marker(to, false) - y;
                var transform = picture.Element(D + "spPr")?.Element(A + "xfrm");
                var editable = !picture.Descendants(A + "srcRect").Any() && transform?.Attribute("rot") is null
                    && transform?.Attribute("flipH") is null && transform?.Attribute("flipV") is null;
                if (width > 0 && height > 0 && new[] { x, y, width, height }.All(double.IsFinite))
                {
                    result.Add(new((string?)properties?.Attribute("id") ?? "", (string?)properties?.Attribute("name") ?? "Image",
                        $"data:{mime};base64,{Convert.ToBase64String(bytes)}", x, y, width, height, editable));
                }
            }
        }
        return result;
    }

    internal static string? Mime(byte[] bytes)
    {
        if (bytes.Length >= 24 && bytes.AsSpan().StartsWith(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }))
        {
            return "image/png";
        }
        if (bytes.Length >= 4 && bytes[0] == 255 && bytes[1] == 216 && bytes[2] == 255)
        {
            return "image/jpeg";
        }
        return null;
    }
}
