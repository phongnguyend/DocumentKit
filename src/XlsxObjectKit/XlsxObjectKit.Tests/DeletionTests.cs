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

internal static class DeletionTests
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
        void Validate(XlsxObjectDocument document)
        {
            using var package = SpreadsheetDocument.Open(new MemoryStream(XlsxWriter.Write(document)), false);
            var errors = new OpenXmlValidator().Validate(package).ToArray();
            Check(errors.Length == 0, string.Join("; ", errors.Select(e => e.Description)));
        }
        var source = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("A1", "1", "number"), new("A2", "2", "number"),
            new("A3", "3", "number"), new("A4", "4", "number"), new("B5", "SUM($A$1:A4)", "formula"),
            new("C5", "$A$2", "formula"), new("D5", "SUM(A2:A3)", "formula"), new("E5", "IF(A4=4,\"A2\",0)", "formula")]);
        source = XlsxEditor.Dimensions(source, "1", "A4", 30, 18);
        source = XlsxEditor.MergeCells(source, "1", "G1:H4");
        source = XlsxEditor.MergeCells(source, "1", "J2:K3");
        source = XlsxEditor.Worksheet(source, "add-sheet", "1", "Other", 1);
        var otherId = XlsxWorkbookExtractor.Extract(source).Sheets[1].Id;
        source = XlsxEditor.SetCells(source, otherId, [new("A1", "Sheet1!A4", "formula"), new("B1", "Sheet1!A2", "formula")]);
        var deleted = XlsxStructureEditor.Delete(source, "1", true, 2, 2);
        var info = XlsxWorkbookExtractor.Extract(deleted);
        Check(info.Sheets[0].Cells.Single(c => c.Address == "A2").Value == "4", "Deletion shifts surviving rows up");
        Check(info.Sheets[0].Cells.Single(c => c.Address == "B3").Formula == "SUM($A$1:A2)", "Deletion contracts spanning ranges");
        Check(info.Sheets[0].Cells.Single(c => c.Address == "C3").Formula == "#REF!", "Deleted scalar reference becomes REF error");
        Check(info.Sheets[0].Cells.Single(c => c.Address == "D3").Formula == "SUM(#REF!)", "Entirely deleted range becomes REF error");
        Check(info.Sheets[0].Cells.Single(c => c.Address == "E3").Formula == "IF(A2=4,\"A2\",0)", "Strings are not rewritten");
        Check(info.Sheets[0].Merges.SequenceEqual(["G1:H2"]) && info.Sheets[0].RowHeights[2] == 30, "Deleted merges removed, spanning merges and dimensions shrink");
        Check(info.Sheets[1].Cells.Single(c => c.Address == "A1").Formula == "Sheet1!A2", "Cross-sheet reference follows moved cell");
        Check(info.Sheets[1].Cells.Single(c => c.Address == "B1").Formula == "#REF!", "Cross-sheet deleted reference becomes error");
        var shifted = XlsxStructureEditor.Delete(deleted, "1", false, 2, 2);
        Check(XlsxWorkbookExtractor.Extract(shifted).Sheets[0].Cells.Single(c => c.Address == "B3").Formula == "SUM(#REF!)", "Column deletion shifts surviving formulas left");
        Validate(shifted);
        var calculated = XlsxReader.Read(WorkbookApplication.Edit(XlsxWriter.Write(source), new("delete-rows", "1", Position: 2, Count: 2)));
        var calculatedCells = XlsxWorkbookExtractor.Extract(calculated).Sheets[0].Cells;
        Check(calculatedCells.Single(c => c.Address == "B3").Value == "5", "Application recalculates contracted SUM");
        Check(calculatedCells.Single(c => c.Address == "C3").Value == "#REF!", "Application calculates deleted reference error");
        Validate(calculated);
        Check(XlsxWorkbookExtractor.Extract(source).Sheets[0].Cells.Any(c => c.Address == "A4"), "Deletion does not mutate input");
        foreach (var formula in new[] { "SUM(A2:A5)", "SUM(A1:A3)", "SUM(A5:A2)", "SUM(2:5)" })
        {
            var doc = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("B6", formula, "formula")]);
            var result = XlsxWorkbookExtractor.Extract(XlsxStructureEditor.Delete(doc, "1", true, 2, 2)).Sheets[0].Cells.Single().Formula;
            var expected = formula switch
            {
                "SUM(A2:A5)" => "SUM(A2:A3)",
                "SUM(A1:A3)" => "SUM(A1:A1)",
                "SUM(A5:A2)" => "SUM(A3:A2)",
                _ => "SUM(2:3)"
            };
            Check(result == expected, "Partial deletion trims endpoints: " + formula);
        }
        var empty = XlsxStructureEditor.Delete(XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("XFD1048576", "last")]), "1", true, 1048576);
        Check(XlsxWorkbookExtractor.Extract(empty).Sheets[0].Cells.Count == 0, "Last worksheet row can be deleted");
        Validate(empty);
        return checks;
    }
}
