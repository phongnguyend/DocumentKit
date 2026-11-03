using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XlsxObjectKit.Application;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using XlsxObjectKit.Readers;
using XlsxObjectKit.Writers;
using static XlsxObjectKit.Models.XlsxNamespaces;

internal static class ImageTests
{
    internal static int Run()
    {
        var checks = 0;
        void Check(bool value, string message)
        {
            if (!value)
            {
                throw new Exception(message);
            }
            checks++;
        }
        var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jZ1kAAAAASUVORK5CYII=");
        var original = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("A1", "Text"), new("B1", "1+1", "formula")]);
        var document = XlsxImageEditor.Edit(original, "1", "add-image", bytes: png, x: 20, y: 30, width: 120, height: 80);
        XlsxImage Picture(XlsxObjectDocument value) => XlsxWorkbookExtractor.Extract(value).Sheets[0].Images.Single();
        var image = Picture(document);
        Check(image is { X: 20, Y: 30, Width: 120, Height: 80 }, "Added image geometry");
        var sheetXml = document.ReadXml("xl/worksheets/sheet1.xml");
        var drawingId = (string)sheetXml.Root!.Element(S + "drawing")!.Attribute(R + "id")!;
        var drawingPart = Resolve("xl/worksheets/sheet1.xml", (string)document.ReadXml("xl/worksheets/_rels/sheet1.xml.rels").Root!.Elements().Single(item => (string?)item.Attribute("Id") == drawingId).Attribute("Target")!);
        XNamespace d = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        var drawing = document.ReadXml(drawingPart);
        var anchor = drawing.Root!.Elements().Single();
        XElement Marker(string name, int row, int column) => new(d + name, new XElement(d + "col", column), new XElement(d + "colOff", 0), new XElement(d + "row", row), new XElement(d + "rowOff", 0));
        anchor.ReplaceWith(new XElement(d + "twoCellAnchor", Marker("from", 1, 1), Marker("to", 3, 3), anchor.Element(d + "pic"), new XElement(d + "clientData")));
        var anchored = Picture(document.WithXml(drawingPart, drawing));
        Check(Math.Abs(anchored.X - 64.01) < 0.001 && anchored.Y == 20 && Math.Abs(anchored.Width - 128.02) < 0.001 && anchored.Height == 40, "Two-cell anchors use worksheet dimensions");
        drawing = document.ReadXml(drawingPart);
        anchor = drawing.Root!.Elements().Single();
        anchor.ReplaceWith(new XElement(d + "oneCellAnchor", Marker("from", 2, 2), anchor.Element(d + "ext"), anchor.Element(d + "pic"), new XElement(d + "clientData")));
        Check(Picture(document.WithXml(drawingPart, drawing)) is { Y: 40, Width: 120, Height: 80 }, "One-cell anchors are extracted");
        Check(XlsxWorkbookExtractor.Extract(original).Sheets[0].Images.Count == 0, "Image editing is immutable");
        using (var stream = new MemoryStream(XlsxWriter.Write(document)))
        using (var package = SpreadsheetDocument.Open(stream, false))
        {
            var errors = new OpenXmlValidator().Validate(package).ToArray();
            Check(errors.Length == 0, "Image package schema: " + string.Join("; ", errors.Select(error => error.Description)));
        }
        document = XlsxImageEditor.Edit(document, "1", "transform-image", image.Id, x: 45, y: 65, width: 200, height: 150);
        Check(Picture(XlsxReader.Read(XlsxWriter.Write(document))) is { X: 45, Y: 65, Width: 200, Height: 150 }, "Image transform survives round trip");
        var firstId = image.Id;
        document = XlsxImageEditor.Edit(document, "1", "add-image", bytes: png);
        var second = XlsxWorkbookExtractor.Extract(document).Sheets[0].Images.Last();
        document = XlsxImageEditor.Edit(document, "1", "replace-image", firstId, png);
        Check(XlsxWorkbookExtractor.Extract(document).Sheets[0].Images.Last() == second, "Replacing image does not change another placement");
        document = XlsxImageEditor.Edit(document, "1", "delete-image", firstId);
        Check(Picture(document).Id == second.Id, "Delete removes only the selected picture");
        Check(XlsxWorkbookExtractor.Extract(document).Sheets[0].Cells.Single(cell => cell.Address == "B1").Formula == "1+1", "Image edits preserve formulas");
        try
        {
            XlsxImageEditor.Edit(document, "1", "transform-image", second.Id, width: double.NaN);
            throw new Exception("Invalid geometry accepted");
        }
        catch (ArgumentException)
        {
            checks++;
        }
        var xml = document.ReadXml("xl/worksheets/sheet1.xml");
        xml.Root!.Element(S + "sheetData")!.AddAfterSelf(new XElement(S + "sheetProtection", new XAttribute("sheet", 1)));
        try
        {
            XlsxImageEditor.Edit(document.WithXml("xl/worksheets/sheet1.xml", xml), "1", "delete-image", second.Id);
            throw new Exception("Protected sheet accepted");
        }
        catch (NotSupportedException)
        {
            checks++;
        }
        var bytes = WorkbookApplication.Edit(XlsxWriter.Write(document), new("transform-image", "1", ImageId: second.Id, X: 10, Y: 15, Width: 110, Height: 115));
        Check(Picture(XlsxReader.Read(bytes)) is { X: 10, Y: 15, Width: 110, Height: 115 }, "Application exposes image edits");
        return checks;
    }
}
