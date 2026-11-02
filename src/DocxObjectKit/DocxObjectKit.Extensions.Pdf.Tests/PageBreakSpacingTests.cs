using System.Xml.Linq;
using DocxObjectKit.Extensions.Pdf;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

internal static class PageBreakSpacingTests
{
    public static void Run(DocxObjectDocument source, Action<bool, string> check)
    {
        XElement Section() => new(W + "sectPr",
            new XElement(W + "pgSz", new XAttribute(W + "w", 8000), new XAttribute(W + "h", 8000)),
            new XElement(W + "pgMar", new XAttribute(W + "top", 800), new XAttribute(W + "bottom", 800),
                new XAttribute(W + "left", 800), new XAttribute(W + "right", 800)));
        XElement Paragraph(string text, int before = 0) => new(W + "p",
            new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "before", before),
                new XAttribute(W + "after", 0), new XAttribute(W + "line", 400), new XAttribute(W + "lineRule", "exact"))),
            new XElement(W + "r", new XElement(W + "t", text)));
        DocxLayoutResult Layout(int mode, bool splitMark = false, bool continued = false)
        {
            var first = Paragraph("First", 600);
            var manual = Paragraph("");
            manual.Element(W + "r")!.Add(new XElement(W + "br", new XAttribute(W + "type", "page")));
            if (continued)
            {
                manual.Element(W + "r")!.Add(new XElement(W + "t", "Continued"));
            }
            var sectionEnd = Paragraph("");
            sectionEnd.Element(W + "pPr")!.Add(Section());
            var xml = source.ReadXml("word/document.xml");
            xml.Root!.Element(W + "body")!.ReplaceNodes(first, manual, Paragraph("After", 600), sectionEnd, Paragraph("Section", 600), Section());
            var settings = new XDocument(new XElement(W + "settings", new XElement(W + "compat",
                new XElement(W + "compatSetting", new XAttribute(W + "name", "compatibilityMode"),
                    new XAttribute(W + "uri", "http://schemas.microsoft.com/office/word"), new XAttribute(W + "val", mode)),
                splitMark ? new XElement(W + "splitPgBreakAndParaMark") : null)));
            using var engine = new DocxLayout(source.WithXml("word/document.xml", xml).WithXml("word/settings.xml", settings));
            return engine.Build();
        }

        var modern = Layout(15);
        check(modern.Pages.Count == 3, "Manual and section breaks must each advance one page.");
        check(Math.Abs(modern.Pages[0].Runs.Single(run => run.Text == "First").Y - 70) < .01,
            "Space before must be retained at the document start.");
        check(Math.Abs(modern.Pages[1].Runs.Single(run => run.Text == "After").Y - 40) < .01,
            "Modern Word layout must suppress space before at an ordinary page start without a phantom break line.");
        check(Math.Abs(modern.Pages[2].Runs.Single(run => run.Text == "Section").Y - 70) < .01,
            "Space before must be retained after a section break.");
        var legacy = Layout(14);
        check(Math.Abs(legacy.Pages[1].Runs.Single(run => run.Text == "After").Y - 70) < .01,
            "Legacy compatibility must preserve space before following a manual break.");
        var split = Layout(15, splitMark: true);
        check(Math.Abs(split.Pages[1].Runs.Single(run => run.Text == "After").Y - 90) < .01,
            "splitPgBreakAndParaMark must preserve an explicit paragraph-mark line after the break.");
        var continued = Layout(15, continued: true);
        check(Math.Abs(continued.Pages[1].Runs.Single(run => run.Text == "Continued").Y - 40) < .01,
            "Run content following a page break must remain on the next page.");
    }
}
