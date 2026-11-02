using System.Globalization;
using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private double DrawingPosition(XElement drawing, string axis, double fallback, double extent)
    {
        var horizontal = axis == "positionH";
        if ((string?)drawing.Attribute("simplePos") is "1" or "true")
        {
            return Number(drawing.Element(WP + "simplePos"), horizontal ? "x" : "y") / 12700;
        }
        var position = drawing.Element(WP + axis);
        if (position is null)
        {
            return fallback;
        }
        var relative = (string?)position.Attribute("relativeFrom");
        var origin = fallback;
        var length = 0d;
        if (relative == "page")
        {
            origin = 0;
            length = horizontal ? result.Width : result.Height;
        }
        else if (relative == "margin")
        {
            origin = horizontal ? left : top;
            length = horizontal ? result.Width - left - right : result.Height - top - bottom;
        }
        else if (horizontal && relative == "column")
        {
            origin = measuringTable ? fallback : FlowLeft;
            length = measuringTable ? result.Width - left - right : FlowWidth;
        }
        else if (horizontal && relative is "leftMargin" or "rightMargin")
        {
            origin = relative == "leftMargin" ? 0 : result.Width - right;
            length = relative == "leftMargin" ? left : right;
        }
        else if (!horizontal && relative is "topMargin" or "bottomMargin")
        {
            origin = relative == "topMargin" ? 0 : result.Height - bottom;
            length = relative == "topMargin" ? top : bottom;
        }
        else if (relative is not ("paragraph" or "character" or "line"))
        {
            warnings.Add("Drawing anchor '" + relative + "' uses approximate placement.");
        }
        if (position.Element(WP + "posOffset") is { } offset && double.TryParse(offset.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && double.IsFinite(number))
        {
            return origin + number / 12700;
        }
        return position.Element(WP + "align")?.Value switch
        {
            "center" => origin + (length - extent) / 2,
            "right" or "bottom" => origin + length - extent,
            "left" or "top" => origin,
            null => fallback,
            _ => Approximate()
        };
        double Approximate()
        {
            warnings.Add("Inside/outside drawing alignment is not implemented.");
            return origin;
        }
    }
}
