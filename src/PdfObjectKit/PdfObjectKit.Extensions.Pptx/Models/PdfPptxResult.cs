namespace PdfObjectKit.Models;

public sealed record PdfPptxResult(byte[] Bytes, int SlideCount, int TextBoxCount, int ImageCount, int ShapeCount, IReadOnlyList<string> Warnings);
