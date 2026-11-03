using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XlsxObjectKit.Application;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Readers;
using XlsxObjectKit.Writers;

internal static class VisibilityTests
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
        var source = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("B2", "Keep"), new("C3", "1+2", "formula")]);
        source = XlsxEditor.Dimensions(source, "1", "B2", 32, 22);
        var hidden = XlsxVisibilityEditor.SetHidden(source, "1", true, 2, 2, true);
        hidden = XlsxVisibilityEditor.SetHidden(hidden, "1", false, 2, 2, true);
        var sheet = XlsxWorkbookExtractor.Extract(XlsxReader.Read(XlsxWriter.Write(hidden))).Sheets[0];
        Check(sheet.HiddenRows.SequenceEqual([2, 3]) && sheet.HiddenColumns.SequenceEqual([2, 3]), "Hidden dimensions survive round trip");
        Check(sheet.RowHeights[2] == 32 && sheet.ColumnWidths[2] == 22, "Hiding preserves original sizes");
        Check(sheet.Cells.Single(c => c.Address == "B2").Value == "Keep", "Hiding retains cell contents");
        Check(sheet.Cells.Single(c => c.Address == "C3").Formula == "1+2", "Hiding retains formulas");
        Check(Math.Abs(XlsxImageExtractor.ColumnPosition(sheet, 4) - 64.01) < .001 && XlsxImageExtractor.RowPosition(sheet, 4) == 20, "Image positions account for hidden dimensions");
        var restored = XlsxReader.Read(WorkbookApplication.Edit(XlsxWriter.Write(hidden), new("unhide-rows", "1", Position: 1, Count: 1048576)));
        restored = XlsxVisibilityEditor.SetHidden(restored, "1", false, 2, 1, false);
        var visible = XlsxWorkbookExtractor.Extract(restored).Sheets[0];
        Check(visible.HiddenRows.Count == 0 && visible.HiddenColumns.SequenceEqual([3]), "Unhide supports selected ranges and all rows");
        Check(visible.RowHeights[2] == 32 && visible.ColumnWidths[2] == 22, "Unhide restores original dimensions");
        Check(XlsxWorkbookExtractor.Extract(source).Sheets[0].HiddenRows.Count == 0, "Visibility changes are immutable");
        using var package = SpreadsheetDocument.Open(new MemoryStream(XlsxWriter.Write(restored)), false);
        var errors = new OpenXmlValidator().Validate(package).ToArray();
        Check(errors.Length == 0, "Visibility XML schema: " + string.Join("; ", errors.Select(e => e.Description)));
        return checks;
    }
}
