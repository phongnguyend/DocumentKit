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

internal static class StructureTests
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
        void Reject(Action action)
        {
            try
            {
                action();
            }
            catch (Exception e) when (e is ArgumentException or NotSupportedException)
            {
                checks++;
                return;
            }
            throw new Exception("Unsafe insertion accepted");
        }
        var original = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("A1", "3", "number"), new("A3", "7", "number"),
            new("B3", "SUM($A$1:A3)", "formula"), new("C3", "IF(A3=7,\"A3\",LOG10(A3))", "formula")]);
        original = XlsxEditor.Dimensions(original, "1", "A3", 30, 18);
        original = XlsxEditor.MergeCells(original, "1", "D1:E3");
        original = XlsxEditor.Worksheet(original, "add-sheet", "1", "Other", 1);
        var otherId = XlsxWorkbookExtractor.Extract(original).Sheets[1].Id;
        original = XlsxEditor.SetCells(original, otherId, [new("A1", "'Sheet1'!$A$3+A3", "formula")]);
        var book = original.ReadXml("xl/workbook.xml");
        book.Root!.Add(new XElement(S + "definedNames", new XElement(S + "definedName", new XAttribute("name", "Amounts"), "Sheet1!$A$1:$A$3")));
        original = original.WithXml("xl/workbook.xml", book);
        var inserted = XlsxStructureEditor.Insert(original, "1", true, 2, 2);
        var info = XlsxWorkbookExtractor.Extract(inserted);
        Check(info.Sheets[0].Cells.Single(c => c.Address == "A5").Value == "7", "Rows shift values");
        Check(info.Sheets[0].Cells.Single(c => c.Address == "B5").Formula == "SUM($A$1:A5)", "Ranges expand and retain absolute markers");
        Check(info.Sheets[0].Cells.Single(c => c.Address == "C5").Formula == "IF(A5=7,\"A3\",LOG10(A5))", "Formula strings and function names remain intact");
        Check(info.Sheets[1].Cells.Single().Formula == "'Sheet1'!$A$5+A3", "Cross-sheet references shift without changing local references");
        Check(info.Sheets[0].Merges.Single() == "D1:E5" && info.Sheets[0].RowHeights[5] == 30, "Merge and row height shift");
        Check(inserted.ReadXml("xl/workbook.xml").Descendants(S + "definedName").Single().Value == "Sheet1!$A$1:$A$5", "Defined name shifts");
        var columns = XlsxStructureEditor.Insert(inserted, "1", false, 1, 2);
        var sheet = XlsxWorkbookExtractor.Extract(columns).Sheets[0];
        Check(sheet.Cells.Single(c => c.Address == "D5").Formula == "SUM($C$1:C5)", "Columns shift absolute and relative references");
        Check(sheet.ColumnWidths[3] == 18 && sheet.Merges.Single() == "F1:G5", "Column widths and merged ranges shift");
        Check(XlsxWorkbookExtractor.Extract(original).Sheets[0].Cells.Any(c => c.Address == "A3"), "Source remains immutable");
        var ranges = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("D5", "SUM(A:A)+SUM(1:3)", "formula")]);
        ranges = XlsxStructureEditor.Insert(ranges, "1", true, 2, 2);
        ranges = XlsxStructureEditor.Insert(ranges, "1", false, 1, 2);
        Check(XlsxWorkbookExtractor.Extract(ranges).Sheets[0].Cells.Single().Formula == "SUM(C:C)+SUM(1:5)", "Whole-row and whole-column formula references update");
        var filters = XlsxEditor.Create();
        var filterXml = filters.ReadXml("xl/worksheets/sheet1.xml");
        filterXml.Root!.Add(new XElement(S + "autoFilter", new XAttribute("ref", "A1:D5"),
            new XElement(S + "filterColumn", new XAttribute("colId", 2), new XElement(S + "filters", new XElement(S + "filter", new XAttribute("val", "Example"))))));
        filters = XlsxStructureEditor.Insert(filters.WithXml("xl/worksheets/sheet1.xml", filterXml), "1", false, 2);
        var updatedFilter = filters.ReadXml("xl/worksheets/sheet1.xml").Descendants(S + "autoFilter").Single();
        Check((string?)updatedFilter.Attribute("ref") == "A1:E5" && (int?)updatedFilter.Element(S + "filterColumn")!.Attribute("colId") == 3, "Filter column indices track inserted columns");
        var bytes = WorkbookApplication.Edit(XlsxWriter.Write(original), new("insert-rows", "1", Position: 2));
        Check(XlsxWorkbookExtractor.Extract(XlsxReader.Read(bytes)).Sheets[0].Cells.Single(c => c.Address == "B4").Value == "10", "Application recalculates adjusted formulas");
        using (var package = SpreadsheetDocument.Open(new MemoryStream(XlsxWriter.Write(columns)), false))
        {
            var errors = new OpenXmlValidator().Validate(package).ToArray();
            Check(errors.Length == 0, "Insertion package schema: " + string.Join("; ", errors.Select(e => e.Description)));
        }
        var last = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("XFD1048576", "Keep")]);
        Reject(() => XlsxStructureEditor.Insert(last, "1", true, 1));
        Reject(() => XlsxStructureEditor.Insert(last, "1", false, 1));
        Reject(() => XlsxStructureEditor.Insert(original, "1", true, 0));
        Reject(() => XlsxStructureEditor.Insert(original, "1", true, 1, 0));
        var complex = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("A1", "SUM(Sheet1:Other!B1)", "formula")]);
        Reject(() => XlsxStructureEditor.Insert(complex, "1", true, 1));
        var shared = original.ReadXml("xl/worksheets/sheet1.xml");
        shared.Descendants(S + "f").First().SetAttributeValue("t", "shared");
        Reject(() => XlsxStructureEditor.Insert(original.WithXml("xl/worksheets/sheet1.xml", shared), "1", true, 1));
        return checks;
    }
}
