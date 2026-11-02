using System.Text.RegularExpressions;
using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private static bool IsLineShape(XElement drawing) => drawing.Descendants(A + "prstGeom").Any(geometry => (string?)geometry.Attribute("prst") == "line");

    private void LineShape(XElement drawing, double originX, double originY)
    {
        var geometry = drawing.Descendants(A + "prstGeom").First();
        var properties = geometry.Parent!;
        var transform = properties.Element(A + "xfrm");
        var extent = drawing.Element(WP + "extent") ?? transform?.Element(A + "ext");
        var width = Number(extent, "cx") / 12700;
        var height = Number(extent, "cy") / 12700;
        var stroke = properties.Element(A + "ln");
        if (stroke?.Element(A + "noFill") is not null)
        {
            return;
        }
        if (width < 0 || height < 0 || width > 20000 || height > 20000 || width != 0 && height != 0 || Number(transform, "rot") != 0)
        {
            warnings.Add("Diagonal or rotated DrawingML lines are not yet rendered.");
            return;
        }
        if (stroke?.Element(A + "prstDash") is { } dash && (string?)dash.Attribute("val") != "solid"
            || stroke?.Element(A + "headEnd") is { } head && (string?)head.Attribute("type") is not (null or "none")
            || stroke?.Element(A + "tailEnd") is { } tail && (string?)tail.Attribute("type") is not (null or "none"))
        {
            warnings.Add("DrawingML line dashes and arrowheads are approximated with solid lines.");
        }
        var colorNode = stroke?.Element(A + "solidFill")?.Elements().FirstOrDefault()
            ?? properties.Parent?.Elements().FirstOrDefault(element => element.Name.LocalName == "style")?.Element(A + "lnRef")?.Elements().FirstOrDefault();
        var color = (string?)colorNode?.Attribute("val") ?? "000000";
        if (colorNode?.Name == A + "schemeClr")
        {
            if (document.PartNames.Contains("word/theme/theme1.xml"))
            {
                theme ??= document.ReadXml("word/theme/theme1.xml");
            }
            var themeColor = theme?.Descendants(A + "clrScheme").FirstOrDefault()?.Element(A + color)?.Elements().FirstOrDefault();
            color = (string?)themeColor?.Attribute("lastClr") ?? (string?)themeColor?.Attribute("val") ?? "000000";
        }
        if (!Regex.IsMatch(color, "^[0-9a-fA-F]{6}$"))
        {
            warnings.Add("Unsupported DrawingML line color uses black.");
            color = "000000";
        }
        if (colorNode?.HasElements == true)
        {
            warnings.Add("DrawingML line color effects are not yet applied.");
        }
        var thickness = Math.Clamp(Number(stroke, "w", 12700) / 12700, 0, 1000);
        var lineX = DrawingPosition(drawing, "positionH", originX, width);
        var lineY = DrawingPosition(drawing, "positionV", originY, height);
        if (width > 0)
        {
            page.Rules.Add(new(lineX, lineY - thickness / 2, width, thickness, color, 0));
        }
        else if (height > 0)
        {
            page.Rules.Add(new(lineX - thickness / 2, lineY, thickness, height, color, 0));
        }
    }
}
