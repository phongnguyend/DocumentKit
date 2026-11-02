namespace PdfObjectKit.Models;

public sealed record PdfDocxResult(byte[] Bytes, int SourcePageCount, int ParagraphCount, int TableCount, int ImageCount, IReadOnlyList<string> Warnings);
