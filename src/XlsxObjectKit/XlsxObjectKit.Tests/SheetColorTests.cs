using System.Xml.Linq;
using XlsxObjectKit.Application;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Readers;
using XlsxObjectKit.Writers;
using static XlsxObjectKit.Models.XlsxNamespaces;

internal static class SheetColorTests
{
    internal static int Run()
    {
        var checks = 0;
        void Check(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(message);
            }
            checks++;
        }
        var source = XlsxEditor.Create();
        var xml = source.ReadXml("xl/worksheets/sheet1.xml");
        xml.Root!.AddFirst(new XElement(S + "sheetPr", new XAttribute("codeName", "Preserved"), new XElement(S + "pageSetUpPr", new XAttribute("fitToPage", true))));
        source = source.WithXml("xl/worksheets/sheet1.xml", xml);
        var colored = XlsxEditor.SetSheetColor(source, "1", "#16a34a");
        var reopened = XlsxReader.Read(XlsxWriter.Write(colored));
        Check(XlsxWorkbookExtractor.Extract(reopened).Sheets[0].TabColor == "#16a34a", "Sheet RGB color round trips");
        var properties = reopened.ReadXml("xl/worksheets/sheet1.xml").Root!.Element(S + "sheetPr")!;
        Check(properties.Elements().First().Name == S + "tabColor" && properties.Element(S + "pageSetUpPr") is not null, "Color uses schema order and preserves sheet properties");
        var cleared = XlsxEditor.SetSheetColor(reopened, "1", null);
        Check(XlsxWorkbookExtractor.Extract(cleared).Sheets[0].TabColor is null, "Clear sheet color");
        Check((string?)cleared.ReadXml("xl/worksheets/sheet1.xml").Root!.Element(S + "sheetPr")!.Attribute("codeName") == "Preserved", "Clearing retains other properties");
        var fixture = FormattingFixture.Create();
        xml = fixture.ReadXml("xl/worksheets/sheet1.xml");
        xml.Root!.AddFirst(new XElement(S + "sheetPr", new XElement(S + "tabColor", new XAttribute("theme", 4), new XAttribute("tint", 0.5))));
        fixture = fixture.WithXml("xl/worksheets/sheet1.xml", xml);
        Check(XlsxWorkbookExtractor.Extract(fixture).Sheets[0].TabColor == "#40ff40", "Theme tab color and tint resolved");
        xml.Descendants(S + "tabColor").Single().ReplaceAttributes(new XAttribute("indexed", 4));
        Check(XlsxWorkbookExtractor.Extract(fixture.WithXml("xl/worksheets/sheet1.xml", xml)).Sheets[0].TabColor == "#0000ff", "Indexed tab color resolved");
        var edited = WorkbookApplication.EditWithDiagnostics(XlsxWriter.Write(source), new("set-sheet-color", "1", TabColor: "#123456"));
        Check(edited.Warnings.Count == 0 && WorkbookApplication.Inspect(edited.Bytes).Sheets[0].TabColor == "#123456", "Application color action skips recalculation");
        var inactive = source.ReadXml("xl/worksheets/sheet1.xml");
        inactive.Root!.Add(new XElement(S + "sheetProtection", new XAttribute("formatCells", "0")));
        var book = source.ReadXml("xl/workbook.xml");
        book.Root!.AddFirst(new XElement(S + "workbookProtection", new XAttribute("lockStructure", "false")));
        var inactiveDocument = source.WithXml("xl/worksheets/sheet1.xml", inactive).WithXml("xl/workbook.xml", book);
        var allowed = XlsxEditor.SetSheetColor(inactiveDocument, "1", "#abcdef");
        Check(XlsxWorkbookExtractor.Extract(allowed).Sheets[0].TabColor == "#abcdef", "Inactive protection elements do not block editing");
        Check(XNode.DeepEquals(allowed.ReadXml("xl/worksheets/sheet1.xml").Root!.Element(S + "sheetProtection"), inactive.Root.Element(S + "sheetProtection")), "Inactive protection settings are preserved");
        foreach (var flag in new[] { "sheet", "objects", "scenarios" })
        {
            var active = new XDocument(inactive);
            active.Root!.Element(S + "sheetProtection")!.SetAttributeValue(flag, "1");
            try
            {
                XlsxEditor.SetSheetColor(source.WithXml("xl/worksheets/sheet1.xml", active), "1", "#abcdef");
                throw new Exception("Active worksheet protection accepted");
            }
            catch (NotSupportedException)
            {
                checks++;
            }
        }
        book.Root.Element(S + "workbookProtection")!.SetAttributeValue("lockStructure", "true");
        try
        {
            XlsxEditor.SetSheetColor(source.WithXml("xl/workbook.xml", book), "1", "#abcdef");
            throw new Exception("Active workbook protection accepted");
        }
        catch (NotSupportedException)
        {
            checks++;
        }
        try
        {
            XlsxEditor.SetSheetColor(source, "1", "red");
            throw new Exception("Invalid color accepted");
        }
        catch (ArgumentException)
        {
            checks++;
        }
        return checks;
    }
}
