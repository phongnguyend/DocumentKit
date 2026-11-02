using System.Xml.Linq;
using DocxObjectKit.Extensions.Pdf;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

internal static class TableFlowTests
{
    public static void Run(DocxObjectDocument source, Action<bool, string> check)
    {
        XElement Paragraph(string prefix, int lines = 1, int advance = 16)
        {
            var run = new XElement(W + "r", new XElement(W + "rPr", new XElement(W + "sz", new XAttribute(W + "val", 20))));
            for (var index = 0; index < lines; index++)
            {
                if (index > 0)
                {
                    run.Add(new XElement(W + "br"));
                }
                run.Add(new XElement(W + "t", prefix + index));
            }
            return new(W + "p", new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "before", 0),
                new XAttribute(W + "after", 0), new XAttribute(W + "line", advance * 20), new XAttribute(W + "lineRule", "exact"))), run);
        }
        XElement Cell(XElement paragraph, string? merge = null) => new(W + "tc",
            new XElement(W + "tcPr", merge is null ? null : new XElement(W + "vMerge", new XAttribute(W + "val", merge))), paragraph);
        XElement Row(int minimum, params XElement[] cells) => new(W + "tr",
            new XElement(W + "trPr", new XElement(W + "trHeight", new XAttribute(W + "val", minimum * 20))), cells);
        XElement Table(params XElement[] rows) => new(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblCellMar", new[] { "top", "left", "bottom", "right" }
                .Select(side => new XElement(W + side, new XAttribute(W + "w", 0))))),
            new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 1600)),
                new XElement(W + "gridCol", new XAttribute(W + "w", 1600))), rows);
        DocxLayoutResult Layout(params XElement[] blocks)
        {
            var xml = source.ReadXml("word/document.xml");
            xml.Root!.Element(W + "body")!.ReplaceNodes(blocks, new XElement(W + "sectPr",
                new XElement(W + "pgSz", new XAttribute(W + "w", 4000), new XAttribute(W + "h", 4000)),
                new XElement(W + "pgMar", new[] { "top", "left", "bottom", "right" }.Select(side => new XAttribute(W + side, 400)))));
            using var engine = new DocxLayout(source.WithXml("word/document.xml", xml));
            return engine.Build();
        }
        var merged = Layout(Table(
            Row(20, Cell(Paragraph("Merged", 3), "restart"), Cell(Paragraph("RowA"))),
            Row(20, Cell(Paragraph(""), "continue"), Cell(Paragraph("RowB"))),
            Row(20, Cell(Paragraph(""), "continue"), Cell(Paragraph("RowC")))), Paragraph("After"));
        check(Math.Abs(merged.Pages[0].Runs.Single(run => run.Text == "RowB0").Y - 40) < .01,
            "Merged text must use its complete row span instead of enlarging the first row.");
        check(Math.Abs(merged.Pages[0].Runs.Single(run => run.Text == "After0").Y - 80) < .01,
            "Merged cell height must not be counted twice in table flow.");
        check(merged.Pages.SelectMany(page => page.Runs).Count(run => run.Text.StartsWith("Merged")) == 3,
            "All merged text lines must render exactly once.");
        var overflow = Layout(Table(Row(20, Cell(Paragraph("Tall", 6), "restart"), Cell(Paragraph("A"))),
            Row(20, Cell(Paragraph(""), "continue"), Cell(Paragraph("B"))),
            Row(20, Cell(Paragraph(""), "continue"), Cell(Paragraph("C")))), Paragraph("After"));
        check(Math.Abs(overflow.Pages[0].Runs.Single(run => run.Text == "After0").Y - 116) < .01,
            "The merged span must grow when its combined height cannot contain the text.");

        var independent = Layout(Paragraph("Filler", advance: 108), Table(Row(0, Cell(Paragraph("Left", 5)), Cell(Paragraph("Right", 4)))));
        check(independent.Pages[0].Runs.Count(run => run.Text.StartsWith("Left")) == 3
            && independent.Pages[0].Runs.Count(run => run.Text.StartsWith("Right")) == 2,
            "Each cell must apply widow/orphan control independently at a row split.");
        check(independent.Pages[1].Runs.Count(run => run.Text.StartsWith("Left")) == 2
            && independent.Pages[1].Runs.Count(run => run.Text.StartsWith("Right")) == 2,
            "Both cells must retain their remaining lines after independent pagination.");

        XElement Header()
        {
            var row = Row(0, Cell(Paragraph("Header")), Cell(Paragraph("Title")));
            row.Element(W + "trPr")!.Add(new XElement(W + "tblHeader"));
            return row;
        }
        var minimum = Layout(Paragraph("Filler", advance: 100), Table(Header(), Row(60, Cell(Paragraph("Body", 5)), Cell(Paragraph("Value")))));
        check(!minimum.Pages[0].Runs.Any(run => run.Text.StartsWith("Body")),
            "A row's declared minimum height must fit before placing its first fragment.");
        check(minimum.Pages[1].Runs.Any(run => run.Text == "Header0") && minimum.Pages[1].Runs.Count(run => run.Text.StartsWith("Body")) == 5,
            "Moving a minimum-height row must repeat its table header.");
        var headerBreak = Layout(Paragraph("Filler", advance: 150), Table(Header(), Row(0, Cell(Paragraph("Body", 12)), Cell(Paragraph("Value")))));
        check(!headerBreak.Pages[0].Runs.Any(run => run.Text == "Header0"), "A repeated header must not split into the previous page's remaining gap.");
        check(headerBreak.Pages.Skip(1).All(page => page.Runs.Any(run => run.Text == "Header0")),
            "A header moved intact must still repeat on later pages.");

        var image = source.ReadXml("word/document.xml").Descendants(WP + "inline").First();
        var picture = new XElement(image);
        picture.Element(WP + "extent")!.SetAttributeValue("cx", 40 * 12700);
        picture.Element(WP + "extent")!.SetAttributeValue("cy", 40 * 12700);
        picture.Element(WP + "effectExtent")?.Remove();
        var imageParagraph = Paragraph("");
        imageParagraph.Element(W + "pPr")!.Element(W + "spacing")!.SetAttributeValue(W + "lineRule", "auto");
        imageParagraph.Element(W + "pPr")!.Element(W + "spacing")!.SetAttributeValue(W + "line", 240);
        imageParagraph.Element(W + "r")!.ReplaceNodes(new XElement(W + "drawing", picture));
        imageParagraph.Add(new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), " ")));
        var pictureLayout = Layout(imageParagraph, Paragraph("After"));
        var renderedImage = pictureLayout.Pages[0].Images.Single();
        var gap = pictureLayout.Pages[0].Runs.Single(run => run.Text == "After0").Y - renderedImage.Y - renderedImage.Height;
        check(gap >= 0 && gap < 6, "An inline picture must reserve descender space, not an additional blank text line.");
        picture.Element(WP + "extent")!.AddAfterSelf(new XElement(WP + "effectExtent",
            new XAttribute("l", 4 * 12700), new XAttribute("r", 2 * 12700), new XAttribute("t", 6 * 12700), new XAttribute("b", 8 * 12700)));
        imageParagraph.Element(W + "pPr")!.Add(new XElement(W + "jc", new XAttribute(W + "val", "center")));
        var decorated = Layout(imageParagraph, Paragraph("After"));
        var decoratedImage = decorated.Pages[0].Images.Single();
        check(Math.Abs(decoratedImage.Y - renderedImage.Y - 6) < .01 && Math.Abs(decoratedImage.X - 81) < .01,
            "Inline image effect extents must participate in vertical placement and centering.");
        check(Math.Abs(decorated.Pages[0].Runs.Single(run => run.Text == "After0").Y
            - pictureLayout.Pages[0].Runs.Single(run => run.Text == "After0").Y - 14) < .01,
            "Both top and bottom image effect extents must reserve paragraph flow space.");
    }
}
