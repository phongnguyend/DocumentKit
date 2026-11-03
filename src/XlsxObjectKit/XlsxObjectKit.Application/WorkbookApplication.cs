using XlsxObjectKit.Editors;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using XlsxObjectKit.Readers;
using XlsxObjectKit.Writers;
using XlsxObjectKit.Calculation;

namespace XlsxObjectKit.Application;

public sealed record XlsxEditRequest(string Action, string SheetId, IReadOnlyList<XlsxCellEdit>? Cells = null,
    XlsxFormat? Format = null, string? Name = null, int Position = 0, double RowHeight = 21, double ColumnWidth = 14,
    string? SourceSheetId = null, string? SourceAddress = null, IReadOnlyList<string>? Addresses = null);

public sealed record XlsxEditResult(byte[] Bytes, IReadOnlyList<string> Warnings);

public static class WorkbookApplication
{
    public static XlsxWorkbook Inspect(byte[] bytes) => XlsxWorkbookExtractor.Extract(XlsxReader.Read(bytes));

    public static byte[] Edit(byte[] bytes, XlsxEditRequest request) => EditWithDiagnostics(bytes, request).Bytes;

    public static XlsxEditResult EditWithDiagnostics(byte[] bytes, XlsxEditRequest request)
    {
        var document = XlsxReader.Read(bytes);
        document = request.Action switch
        {
            "cells" => XlsxEditor.SetCells(document, request.SheetId, request.Cells ?? throw new ArgumentException("Provide cell edits."), request.Format),
            "dimensions" => XlsxEditor.Dimensions(document, request.SheetId, request.Cells?.Single().Address ?? throw new ArgumentException("Choose a cell."), request.RowHeight, request.ColumnWidth),
            "add-sheet" or "rename-sheet" or "delete-sheet" or "move-sheet" => XlsxEditor.Worksheet(document, request.Action, request.SheetId, request.Name, request.Position),
            "recalculate" => document,
            "apply-format" => XlsxEditor.ApplyFormat(document, request.SourceSheetId ?? throw new ArgumentException("Choose a source worksheet."),
                request.SourceAddress ?? throw new ArgumentException("Choose a source cell."), request.SheetId, request.Addresses ?? throw new ArgumentException("Choose destination cells.")),
            _ => throw new ArgumentException("Unknown edit action.")
        };
        XlsxWorkbookExtractor.Extract(document);
        if (request.Action == "apply-format")
        {
            return new(XlsxWriter.Write(document), []);
        }
        var calculation = XlsxCalculator.Calculate(document);
        return new(XlsxWriter.Write(calculation.Document), calculation.Warnings);
    }
}
