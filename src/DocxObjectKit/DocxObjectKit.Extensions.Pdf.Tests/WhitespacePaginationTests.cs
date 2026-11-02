using System.Xml.Linq;
using DocxObjectKit.Extensions.Pdf;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

internal static class WhitespacePaginationTests
{
    public static void Run(DocxObjectDocument source, Action<bool, string> check)
    {
        DocxLayoutResult Layout(string text, string alignment = "left", int pageHeight = 1600, int pageWidth = 4000)
        {
            XElement Paragraph(string value) => new(W + "p",
                new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", alignment)),
                    new XElement(W + "widowControl", new XAttribute(W + "val", 0)),
                    new XElement(W + "spacing", new XAttribute(W + "before", 0), new XAttribute(W + "after", 0),
                        new XAttribute(W + "line", 400), new XAttribute(W + "lineRule", "exact"))),
                new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value)));
            var xml = source.ReadXml("word/document.xml");
            xml.Root!.Element(W + "body")!.ReplaceNodes(Paragraph(text), Paragraph("Following"),
                new XElement(W + "sectPr", new XElement(W + "pgSz", new XAttribute(W + "w", pageWidth), new XAttribute(W + "h", pageHeight)),
                    new XElement(W + "pgMar", new XAttribute(W + "top", 400), new XAttribute(W + "bottom", 400),
                        new XAttribute(W + "left", 200), new XAttribute(W + "right", 200))));
            using var engine = new DocxLayout(source.WithXml("word/document.xml", xml));
            return engine.Build();
        }

        var plain = Layout("Body");
        var trailing = Layout("Body" + new string(' ', 100));
        check(plain.Pages.Count == 1 && trailing.Pages.Count == 1, "Overflowing trailing spaces must not add a line or page.");
        check(Math.Abs(plain.Pages[0].Runs.Single(run => run.Text == "Following").Y - trailing.Pages[0].Runs.Single(run => run.Text == "Following").Y) < .01,
            "Trailing spaces must not shift the following paragraph.");
        var interior = Layout("Body" + new string(' ', 100) + "Next", pageHeight: 3000);
        var runs = interior.Pages[0].Runs;
        check(Math.Abs(runs.Single(run => run.Text == "Next").Y - runs.Single(run => run.Text == "Body").Y - 20) < .01,
            "Overflowing inter-word spaces must still wrap subsequent text.");
        foreach (var alignment in new[] { "center", "right" })
        {
            var normal = Layout("Body", alignment);
            var padded = Layout("Body   ", alignment);
            check(Math.Abs(normal.Pages[0].Runs[0].X - padded.Pages[0].Runs[0].X) < .01,
                "Trailing spaces must not change " + alignment + " alignment.");
        }
        var widthReference = Layout("Borderline word");
        var textWidth = widthReference.Pages[0].Runs.Where(run => run.Id == "p0").Sum(run => run.Width);
        var roundedWidth = (int)Math.Round((textWidth + 20) * 20) - 1;
        check(Layout("Borderline word", pageWidth: roundedWidth).Pages.Count == 1,
            "Sub-pixel font-advance differences at a twip-rounded line edge must not create an extra page.");
        check(Layout("Borderline word", pageWidth: roundedWidth - 20).Pages.Count == 2,
            "A meaningful text overrun must still wrap instead of extending outside the margin.");
    }
}
