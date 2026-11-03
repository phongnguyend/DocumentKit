using XlsxObjectKit.Editors;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using XlsxObjectKit.Readers;
using XlsxObjectKit.Writers;
using XlsxObjectKit.Calculation;

namespace XlsxObjectKit.Application;

public sealed record XlsxEditRequest(string Action, string SheetId, IReadOnlyList<XlsxCellEdit>? Cells = null,
    XlsxFormat? Format = null, string? Name = null, int Position = 0, double RowHeight = 21, double ColumnWidth = 14,
    string? SourceSheetId = null, string? SourceAddress = null, IReadOnlyList<string>? Addresses = null, string? ImageId = null, string? ImageBase64 = null,
    double X = 0, double Y = 0, double Width = 240, double Height = 160, string? Range = null, int Count = 1, string? Operator = null, double Value = 0, double SecondValue = 0,
    string Color = "#9c0006", string Background = "#ffc7ce", int RuleId = -1, string? TabColor = null);

public sealed record XlsxEditResult(byte[] Bytes, IReadOnlyList<string> Warnings);

public static class WorkbookApplication
{
    public static XlsxWorkbook Inspect(byte[] bytes) => XlsxWorkbookExtractor.Extract(XlsxReader.Read(bytes));

    private static byte[]? DecodeImage(string? value)
    {
        if (value is null)
        {
            return null;
        }
        if (value.Length > 7 * 1024 * 1024)
        {
            throw new ArgumentException("Image payload exceeds 5 MB.");
        }
        try
        {
            return Convert.FromBase64String(value);
        }
        catch (FormatException)
        {
            throw new ArgumentException("Invalid image payload.");
        }
    }

    public static byte[] Edit(byte[] bytes, XlsxEditRequest request) => EditWithDiagnostics(bytes, request).Bytes;

    public static XlsxEditResult EditWithDiagnostics(byte[] bytes, XlsxEditRequest request)
    {
        var document = XlsxCalculator.ExpandSharedFormulas(XlsxReader.Read(bytes));
        document = request.Action switch
        {
            "cells" => XlsxEditor.SetCells(document, request.SheetId, request.Cells ?? throw new ArgumentException("Provide cell edits."), request.Format),
            "resize-row" or "resize-column" => XlsxEditor.Dimensions(document, request.SheetId,
                request.Cells?.Single().Address ?? throw new ArgumentException("Choose a cell."), request.RowHeight, request.ColumnWidth,
                resizeRow: request.Action == "resize-row", resizeColumn: request.Action == "resize-column"),
            "dimensions" => XlsxEditor.Dimensions(document, request.SheetId, request.Cells?.Single().Address ?? throw new ArgumentException("Choose a cell."), request.RowHeight, request.ColumnWidth),
            "add-sheet" or "rename-sheet" or "delete-sheet" or "move-sheet" => XlsxEditor.Worksheet(document, request.Action, request.SheetId, request.Name, request.Position),
            "set-sheet-color" => XlsxEditor.SetSheetColor(document, request.SheetId, request.TabColor),
            "recalculate" => document,
            "add-conditional-rule" => XlsxConditionalFormatEditor.Edit(document, request.SheetId, request.Range, request.Operator, request.Value, request.SecondValue, request.Color, request.Background),
            "remove-conditional-rule" => request.RuleId < 0 ? throw new ArgumentException("Choose a conditional formatting rule.")
                : XlsxConditionalFormatEditor.Edit(document, request.SheetId, null, null, 0, 0, request.Color, request.Background, request.RuleId),
            "hide-rows" or "hide-columns" or "unhide-rows" or "unhide-columns" => XlsxVisibilityEditor.SetHidden(document, request.SheetId,
                request.Action.EndsWith("rows", StringComparison.Ordinal), request.Position, request.Count, request.Action.StartsWith("hide-", StringComparison.Ordinal)),
            "insert-rows" or "insert-columns" => XlsxStructureEditor.Insert(document, request.SheetId, request.Action == "insert-rows", request.Position, request.Count),
            "delete-rows" or "delete-columns" => XlsxStructureEditor.Delete(document, request.SheetId, request.Action == "delete-rows", request.Position, request.Count),
            "merge-cells" or "unmerge-cells" => XlsxEditor.MergeCells(document, request.SheetId, request.Range ?? throw new ArgumentException("Select a cell range."), request.Action == "unmerge-cells"),
            "add-image" or "replace-image" or "transform-image" or "delete-image" => XlsxImageEditor.Edit(document, request.SheetId, request.Action,
                request.ImageId, DecodeImage(request.ImageBase64), request.X, request.Y, request.Width, request.Height),
            "apply-format" => XlsxEditor.ApplyFormat(document, request.SourceSheetId ?? throw new ArgumentException("Choose a source worksheet."),
                request.SourceAddress ?? throw new ArgumentException("Choose a source cell."), request.SheetId, request.Addresses ?? throw new ArgumentException("Choose destination cells.")),
            _ => throw new ArgumentException("Unknown edit action.")
        };
        XlsxWorkbookExtractor.Extract(document);
        if (request.Action is "set-sheet-color" or "add-conditional-rule" or "remove-conditional-rule" or "apply-format" or "resize-row" or "resize-column" or "merge-cells" or "unmerge-cells" || request.Action.EndsWith("-image", StringComparison.Ordinal)
            || request.Action.StartsWith("hide-", StringComparison.Ordinal) || request.Action.StartsWith("unhide-", StringComparison.Ordinal))
        {
            return new(XlsxWriter.Write(document), []);
        }
        var calculation = XlsxCalculator.Calculate(document);
        return new(XlsxWriter.Write(calculation.Document), calculation.Warnings);
    }
}
