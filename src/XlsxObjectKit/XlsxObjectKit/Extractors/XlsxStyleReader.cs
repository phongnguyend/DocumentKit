using System.Globalization;
using System.Xml.Linq;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Extractors;

internal static class XlsxStyleReader
{
    private static readonly string[] Palette = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
        "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 " +
        "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF " +
        "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 " +
        "003366 339966 003300 333300 993300 993366 333399 333333").Split(' ');

    public static IReadOnlyList<XlsxStyle> Read(XlsxObjectDocument document, string? part, string? themePart)
    {
        if (part is null)
        {
            return [new(false, false, "#000000", "#ffffff", "general", "General")];
        }
        var root = document.ReadXml(part).Root!;
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var theme = themePart is null ? null : document.ReadXml(themePart);
        var scheme = theme?.Descendants(drawing + "clrScheme").FirstOrDefault();
        string[] themeNames = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
        string[] defaults = ["FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"];
        var indexed = root.Element(S + "colors")?.Element(S + "indexedColors")?.Elements().Select(item => (string?)item.Attribute("rgb") ?? "000000").ToArray() ?? Palette;
        string Color(XElement? node, string fallback)
        {
            var hex = (string?)node?.Attribute("rgb");
            if (hex is null && (int?)node?.Attribute("theme") is { } index && index >= 0 && index < themeNames.Length)
            {
                var color = scheme?.Element(drawing + themeNames[index])?.Elements().FirstOrDefault();
                hex = (string?)color?.Attribute(color.Name.LocalName == "sysClr" ? "lastClr" : "val") ?? defaults[index];
            }
            if (hex is null && (int?)node?.Attribute("indexed") is { } paletteIndex && paletteIndex >= 0 && paletteIndex < indexed.Length)
            {
                hex = indexed[paletteIndex];
            }
            if (hex?.Length == 8)
            {
                hex = hex[2..];
            }
            if (hex?.Length != 6 || !int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
            {
                return fallback;
            }
            var tint = Math.Clamp(Number(node, "tint", 0), -1, 1);
            return Tint(rgb, tint).ToLowerInvariant();
        }
        var fonts = root.Element(S + "fonts")?.Elements().ToArray() ?? [];
        var fills = root.Element(S + "fills")?.Elements().ToArray() ?? [];
        var borders = root.Element(S + "borders")?.Elements().ToArray() ?? [];
        var bases = root.Element(S + "cellStyleXfs")?.Elements().ToArray() ?? [];
        var result = new List<XlsxStyle>();
        foreach (var xf in root.Element(S + "cellXfs")?.Elements() ?? [])
        {
            var parent = bases.ElementAtOrDefault((int?)xf.Attribute("xfId") ?? 0);
            int Id(string name, string apply)
            {
                return (string?)xf.Attribute(apply) is "0" or "false"
                    ? (int?)parent?.Attribute(name) ?? 0
                    : (int?)xf.Attribute(name) ?? (int?)parent?.Attribute(name) ?? 0;
            }
            var font = fonts.ElementAtOrDefault(Id("fontId", "applyFont"));
            var fill = fills.ElementAtOrDefault(Id("fillId", "applyFill"));
            var border = borders.ElementAtOrDefault(Id("borderId", "applyBorder"));
            var format = Id("numFmtId", "applyNumberFormat");
            var alignment = new XElement(parent?.Element(S + "alignment") ?? new XElement(S + "alignment"));
            if ((string?)xf.Attribute("applyAlignment") is not ("0" or "false") && xf.Element(S + "alignment") is { } direct)
            {
                foreach (var attribute in direct.Attributes())
                {
                    alignment.SetAttributeValue(attribute.Name, attribute.Value);
                }
            }
            var fontName = (string?)font?.Element(S + "name")?.Attribute("val") ?? "Calibri";
            var fontScheme = (string?)font?.Element(S + "scheme")?.Attribute("val");
            if (fontScheme is "major" or "minor")
            {
                fontName = (string?)theme?.Descendants(drawing + (fontScheme + "Font")).FirstOrDefault()?.Element(drawing + "latin")?.Attribute("typeface") ?? fontName;
            }
            var edges = new Dictionary<string, XlsxBorder>();
            foreach (var side in new[] { "left", "right", "top", "bottom" })
            {
                var edge = border?.Element(S + side);
                if ((string?)edge?.Attribute("style") is { } line && line != "none")
                {
                    edges[side] = new(line, Color(edge.Element(S + "color"), "#000000"));
                }
            }
            var pattern = fill?.Element(S + "patternFill");
            result.Add(new XlsxStyle(On(font?.Element(S + "b")), On(font?.Element(S + "i")),
                Color(font?.Element(S + "color"), "#000000"),
                (string?)pattern?.Attribute("patternType") == "solid" ? Color(pattern.Element(S + "fgColor"), "#ffffff") : "#ffffff",
                (string?)alignment.Attribute("horizontal") ?? "general",
                (string?)root.Element(S + "numFmts")?.Elements().FirstOrDefault(item => (int?)item.Attribute("numFmtId") == format)?.Attribute("formatCode") ?? "")
            {
                NumberFormatId = format,
                FontName = fontName,
                FontSize = Math.Clamp(Number(font?.Element(S + "sz"), "val", 11), 1, 409),
                Underline = font?.Element(S + "u") is { } underline ? (string?)underline.Attribute("val") ?? "single" : "none",
                Strike = On(font?.Element(S + "strike")),
                WrapText = (string?)alignment.Attribute("wrapText") is "1" or "true",
                VerticalAlignment = (string?)alignment.Attribute("vertical") ?? "bottom",
                Indent = Math.Clamp((int?)alignment.Attribute("indent") ?? 0, 0, 250),
                TextRotation = (int?)alignment.Attribute("textRotation") ?? 0,
                Borders = edges
            });
        }
        return result.Count > 0 ? result : [new(false, false, "#000000", "#ffffff", "general", "General")];
    }

    private static bool On(XElement? element) => element is not null && (string?)element.Attribute("val") is not ("0" or "false");

    private static double Number(XElement? element, string attribute, double fallback) => double.TryParse((string?)element?.Attribute(attribute), NumberStyles.Float,
        CultureInfo.InvariantCulture, out var value) && double.IsFinite(value) ? value : fallback;

    // Spreadsheet tint changes HSL luminance, preserving hue and saturation.
    private static string Tint(int rgb, double tint)
    {
        var r = ((rgb >> 16) & 255) / 255d;
        var g = ((rgb >> 8) & 255) / 255d;
        var b = (rgb & 255) / 255d;
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var lightness = (max + min) / 2;
        var delta = max - min;
        var saturation = delta == 0 ? 0 : delta / (1 - Math.Abs(2 * lightness - 1));
        var hue = delta == 0 ? 0 : max == r ? ((g - b) / delta + 6) % 6 : max == g ? (b - r) / delta + 2 : (r - g) / delta + 4;
        lightness = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;
        var chroma = (1 - Math.Abs(2 * lightness - 1)) * saturation;
        var x = chroma * (1 - Math.Abs(hue % 2 - 1));
        var m = lightness - chroma / 2;
        var (red, green, blue) = hue switch
        {
            < 1 => (chroma, x, 0d),
            < 2 => (x, chroma, 0d),
            < 3 => (0d, chroma, x),
            < 4 => (0d, x, chroma),
            < 5 => (x, 0d, chroma),
            _ => (chroma, 0d, x)
        };
        return $"#{(int)Math.Round((red + m) * 255):X2}{(int)Math.Round((green + m) * 255):X2}{(int)Math.Round((blue + m) * 255):X2}";
    }
}
