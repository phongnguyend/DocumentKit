using PptxObjectKit.Editors;
using PptxObjectKit.Models;
using PptxObjectKit.Readers;
using PptxObjectKit.Writers;

namespace PptxObjectKit.Application;

public sealed record PptxEditRequest(string Kind, int SlideNumber = 1, string ObjectId = "", string Text = "", int[]? Order = null);

public static class PresentationApplication
{
    public static byte[] Edit(byte[] bytes, PptxEditRequest request)
    {
        var source = PptxReader.Read(bytes);
        PptxObjectDocument edited = request.Kind switch
        {
            "text" => PptxSlideEditor.UpdateText(source, request.SlideNumber, request.ObjectId, request.Text),
            "reorder" => PptxSlideEditor.Reorder(source, request.Order ?? throw new ArgumentException("Slide order is required.")),
            "delete-slide" => PptxSlideEditor.Delete(source, request.SlideNumber),
            _ => throw new ArgumentException("Unknown edit kind.")
        };
        return PptxWriter.Write(edited);
    }
}
