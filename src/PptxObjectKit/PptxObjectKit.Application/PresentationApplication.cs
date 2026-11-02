using PptxObjectKit.Editors;
using PptxObjectKit.Models;
using PptxObjectKit.Readers;
using PptxObjectKit.Writers;

namespace PptxObjectKit.Application;

public sealed record PptxEditRequest(string Kind, int SlideNumber = 1, string ObjectId = "", string Text = "", int[]? Order = null, string? ImageBase64 = null, double X = 36, double Y = 36, double Width = 180, double Height = 120, bool UpdateBounds = false, PptxOcrLine[]? OcrLines = null, bool ReplaceOcrImage = false);

public static class PresentationApplication
{
    public static byte[] Edit(byte[] bytes, PptxEditRequest request)
    {
        var source = PptxReader.Read(bytes);
        PptxObjectDocument edited = request.Kind switch
        {
            "add-image" or "replace-image" or "delete-image" or "resize-image" => PptxImageEditor.Edit(source, request.SlideNumber, request.ObjectId, request.Kind, request.ImageBase64 is null ? null : DecodeImage(request.ImageBase64), request.X, request.Y, request.Width, request.Height),
            "apply-image-ocr" => PptxImageOcrEditor.Apply(source, request.SlideNumber, request.ObjectId, request.OcrLines ?? [], request.ReplaceOcrImage),
            "transform-text" => PptxSlideEditor.TransformText(source, request.SlideNumber, request.ObjectId, request.X, request.Y, request.Width, request.Height),
            "text" => PptxSlideEditor.UpdateText(source, request.SlideNumber, request.ObjectId, request.Text),
            "reorder" => PptxSlideEditor.Reorder(source, request.Order ?? throw new ArgumentException("Slide order is required.")),
            "duplicate-slide" => PptxSlideEditor.Duplicate(source, request.SlideNumber),
            "insert-blank-slide" => PptxSlideEditor.InsertBlank(source, request.SlideNumber),
            "delete-slide" => PptxSlideEditor.Delete(source, request.SlideNumber),
            _ => throw new ArgumentException("Unknown edit kind.")
        };
        if (request.Kind == "replace-image" && request.UpdateBounds)
        {
            edited = PptxImageEditor.Edit(edited, request.SlideNumber, request.ObjectId, "resize-image", x: request.X, y: request.Y, width: request.Width, height: request.Height);
        }
        return PptxWriter.Write(edited);
    }
    private static byte[] DecodeImage(string value)
    {
        if (value.Length > 7 * 1024 * 1024)
        {
            throw new ArgumentException("Image exceeds the upload limit.");
        }
        try
        {
            return Convert.FromBase64String(value);
        }
        catch (FormatException exception)
        {
            throw new ArgumentException("Invalid image data.", exception);
        }
    }
}
