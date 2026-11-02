using System.Xml.Linq;
using DocxObjectKit.Extensions.Pdf;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

internal static class LineShapeTests
{
    public static void Run(DocxObjectDocument source, Action<bool, string> check)
    {
        XNamespace shapes = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        XElement Line(bool vertical, bool hidden = false) => new(W + "p", new XElement(W + "r", new XElement(W + "drawing",
            new XElement(WP + "anchor", new XElement(WP + "positionH", new XAttribute("relativeFrom", "page"), new XElement(WP + "posOffset", 40 * 12700)),
                new XElement(WP + "positionV", new XAttribute("relativeFrom", "page"), new XElement(WP + "posOffset", 100 * 12700)),
                new XElement(WP + "extent", new XAttribute("cx", vertical ? 0 : 240 * 12700), new XAttribute("cy", vertical ? 120 * 12700 : 0)),
                new XElement(A + "graphic", new XElement(A + "graphicData", new XElement(shapes + "wsp",
                    new XElement(shapes + "spPr", new XElement(A + "prstGeom", new XAttribute("prst", "line")), new XElement(A + "ln", new XAttribute("w", 4 * 12700), hidden ? new XElement(A + "noFill") : null)),
                    new XElement(shapes + "style", new XElement(A + "lnRef", new XAttribute("idx", 1), new XElement(A + "schemeClr", new XAttribute("val", "accent2")))))))))));
        var xml = source.ReadXml("word/document.xml");
        xml.Root!.Element(W + "body")!.ReplaceNodes(Line(false), Line(true), Line(false, true), new XElement(W + "sectPr"));
        var theme = new XDocument(new XElement(A + "theme", new XElement(A + "themeElements", new XElement(A + "clrScheme", new XElement(A + "accent2", new XElement(A + "srgbClr", new XAttribute("val", "168A45")))))));
        var document = source.WithXml("word/document.xml", xml).WithXml("word/theme/theme1.xml", theme);
        using var engine = new DocxLayout(document);
        var result = engine.Build();
        var rules = result.Pages.SelectMany(page => page.Rules).ToArray();
        check(rules.Length == 2, "DrawingML lines must render once; noFill lines must stay hidden.");
        check(rules.Any(rule => rule.X == 40 && rule.Y == 98 && rule.Width == 240 && rule.Height == 4 && rule.Fill == "168A45"), "Horizontal lines must preserve theme color, thickness and anchor position.");
        check(rules.Any(rule => rule.X == 38 && rule.Y == 100 && rule.Width == 4 && rule.Height == 120), "Vertical lines must preserve their stroke geometry.");
        check(!result.Warnings.Any(warning => warning.Contains("Unsupported drawing")), "Supported line shapes must not be treated as missing pictures.");
        check(DocxPdfConverter.Convert(document).Bytes.Length > 100, "Line-only documents must export to PDF.");
    }
}
