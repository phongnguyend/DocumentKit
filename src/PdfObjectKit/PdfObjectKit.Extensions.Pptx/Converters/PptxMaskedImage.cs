using PdfObjectKit.Rendering;
using StbImageSharp;

namespace PdfObjectKit.Converters;

internal static class PptxMaskedImage
{
    internal static byte[] Decode(PdfCanvasCommand command, byte[] jpeg)
    {
        var pixels = (long)command.Width * command.Height;
        if (command.Width <= 0 || command.Height <= 0 || pixels > 16000000)
        {
            throw new InvalidDataException("Masked image exceeds the pixel limit.");
        }
        var alpha = Convert.FromBase64String(command.AlphaData!);
        if (alpha.Length != pixels || command.Matte is { } matte && (matte.Length != 3 || matte.Any(v => !double.IsFinite(v) || v < 0 || v > 1)))
        {
            throw new InvalidDataException("Invalid image soft mask or matte.");
        }
        using var stream = new MemoryStream(jpeg);
        var info = ImageInfo.FromStream(stream);
        if (info is null || info.Value.Width != command.Width || info.Value.Height != command.Height)
        {
            throw new InvalidDataException("JPEG dimensions do not match the soft mask.");
        }
        var image = ImageResult.FromMemory(jpeg, ColorComponents.RedGreenBlueAlpha);
        for (var pixel = 0; pixel < alpha.Length; pixel++)
        {
            var opacity = alpha[pixel];
            for (var channel = 0; channel < 3; channel++)
            {
                var value = image.Data[pixel * 4 + channel];
                image.Data[pixel * 4 + channel] = opacity == 0 ? (byte)0 : command.Matte is { } background
                    ? (byte)Math.Clamp(Math.Round((value - (255 - opacity) * background[channel]) * 255 / opacity), 0, 255) : value;
            }
            image.Data[pixel * 4 + 3] = opacity;
        }
        return PptxPng.Encode(image.Data, command.Width, command.Height);
    }
}
