using System.Buffers.Binary;
using SkiaSharp;

namespace DocxObjectKit.Extensions.Pdf;

// Portable subset of MS-EMF: solid filled paths, logical/device mapping and path clipping.
// Unsupported drawing records reject the image rather than returning an incomplete picture.
internal static class EmfImage
{
    public static bool IsEmf(byte[] bytes) => bytes.Length >= 88 && Read(bytes, 0) == 1 && Read(bytes, 40) == 0x464d4520;

    private static int Read(ReadOnlySpan<byte> bytes, int offset)
    {
        if (offset < 0 || offset > bytes.Length - 4)
        {
            throw new InvalidDataException("Truncated EMF record.");
        }
        return BinaryPrimitives.ReadInt32LittleEndian(bytes[offset..]);
    }

    public static byte[] Decode(byte[] bytes, double widthPoints, double heightPoints)
    {
        try
        {
            try
            {
                return DecodePortable(bytes, widthPoints, heightPoints);
            }
            catch (NotSupportedException)
            {
                return EmfDiagram.Decode(bytes, widthPoints, heightPoints);
            }
        }
        catch (NotSupportedException) when (OperatingSystem.IsWindows())
        {
            return WindowsEmfImage.Decode(bytes, widthPoints, heightPoints);
        }
    }

    private static byte[] DecodePortable(byte[] bytes, double widthPoints, double heightPoints)
    {
        if (!IsEmf(bytes) || Read(bytes, 48) != bytes.Length)
        {
            throw new InvalidDataException("Invalid EMF header or length.");
        }
        var left = Read(bytes, 8);
        var top = Read(bytes, 12);
        var width = (double)Read(bytes, 16) - left + 1;
        var height = (double)Read(bytes, 20) - top + 1;
        if (width <= 0 || height <= 0 || width > 1000000 || height > 1000000 || !double.IsFinite(widthPoints) || !double.IsFinite(heightPoints))
        {
            throw new InvalidDataException("Invalid EMF dimensions.");
        }
        var pixelsX = (int)Math.Clamp(Math.Ceiling(widthPoints * 3), 32, 4096);
        var pixelsY = (int)Math.Clamp(Math.Ceiling(heightPoints * 3), 32, 4096);
        using var bitmap = new SKBitmap(pixelsX, pixelsY, SKColorType.Rgba8888, SKAlphaType.Premul);
        using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.Transparent);
        canvas.Scale((float)(pixelsX / width), (float)(pixelsY / height));
        canvas.Translate(-left, -top);
        canvas.Save();
        using var path = new SKPath();
        using var paint = new SKPaint { IsAntialias = true, Style = SKPaintStyle.Fill, Color = SKColors.White };
        var brushes = new Dictionary<int, SKColor?>();
        SKColor? brush = SKColors.White;
        float windowX = 0, windowY = 0, windowWidth = 1, windowHeight = 1;
        float viewportX = 0, viewportY = 0, viewportWidth = 1, viewportHeight = 1;
        var fill = SKPathFillType.EvenOdd;
        var offset = 0;
        var records = 0;
        var points = 0;
        var ended = false;
        SKPoint Map(float x, float y) => new((x - windowX) * viewportWidth / windowWidth + viewportX, (y - windowY) * viewportHeight / windowHeight + viewportY);
        void ResetClip()
        {
            canvas.Restore();
            canvas.Save();
        }
        while (offset < bytes.Length)
        {
            var remaining = bytes.AsSpan(offset);
            var type = Read(remaining, 0);
            var size = Read(remaining, 4);
            if (size < 8 || size % 4 != 0 || size > remaining.Length || ++records > 100000)
            {
                throw new InvalidDataException("Invalid EMF record size or record limit exceeded.");
            }
            var record = remaining[..size];
            switch (type)
            {
                case 1:
                    if (offset != 0 || size < 88)
                    {
                        throw new InvalidDataException("Invalid EMF header position.");
                    }
                    break;
                case 14:
                    ended = true;
                    if (offset + size != bytes.Length)
                    {
                        throw new InvalidDataException("Unexpected data after EMF end record.");
                    }
                    break;
                case 9:
                    windowWidth = Read(record, 8);
                    windowHeight = Read(record, 12);
                    if (windowWidth == 0 || windowHeight == 0)
                    {
                        throw new InvalidDataException("Zero EMF window extent.");
                    }
                    break;
                case 10:
                    windowX = Read(record, 8);
                    windowY = Read(record, 12);
                    break;
                case 11:
                    viewportWidth = Read(record, 8);
                    viewportHeight = Read(record, 12);
                    break;
                case 12:
                    viewportX = Read(record, 8);
                    viewportY = Read(record, 12);
                    break;
                case 17:
                    if (Read(record, 8) is not (1 or 8))
                    {
                        throw new NotSupportedException("EMF mapping mode is not supported.");
                    }
                    break;
                case 19:
                    fill = Read(record, 8) switch
                    {
                        1 => SKPathFillType.EvenOdd,
                        2 => SKPathFillType.Winding,
                        _ => throw new NotSupportedException("EMF fill mode is unsupported.")
                    };
                    path.FillType = fill;
                    break;
                case 27:
                    path.MoveTo(Map(Read(record, 8), Read(record, 12)));
                    break;
                case 39:
                    var handle = Read(record, 8);
                    var style = Read(record, 12);
                    var color = Read(record, 16);
                    if (style is not (0 or 1))
                    {
                        throw new NotSupportedException("Patterned EMF brushes are unsupported.");
                    }
                    brushes[handle] = style == 1 ? null : new SKColor((byte)color, (byte)(color >> 8), (byte)(color >> 16));
                    break;
                case 37:
                    var selected = Read(record, 8);
                    if (selected < 0)
                    {
                        brush = (selected & int.MaxValue) switch
                        {
                            0 => SKColors.White,
                            1 => SKColors.LightGray,
                            2 => SKColors.Gray,
                            3 => SKColors.DarkGray,
                            4 => SKColors.Black,
                            5 => null,
                            _ => throw new NotSupportedException("EMF stock object is unsupported.")
                        };
                    }
                    else if (!brushes.TryGetValue(selected, out brush))
                    {
                        throw new InvalidDataException("EMF brush handle was not defined.");
                    }
                    break;
                case 40:
                    brushes.Remove(Read(record, 8));
                    break;
                case 59:
                    path.Reset();
                    path.FillType = fill;
                    break;
                case 60:
                    break;
                case 61:
                    path.Close();
                    break;
                case 62:
                    if (brush is { } active)
                    {
                        paint.Color = active;
                        canvas.DrawPath(path, paint);
                    }
                    path.Reset();
                    path.FillType = fill;
                    break;
                case 67:
                    var mode = Read(record, 8);
                    if (mode == 5)
                    {
                        ResetClip();
                    }
                    else if (mode != 1)
                    {
                        throw new NotSupportedException("EMF clipping combination is unsupported.");
                    }
                    canvas.ClipPath(path, SKClipOperation.Intersect, true);
                    path.Reset();
                    path.FillType = fill;
                    break;
                case 75:
                    if (Read(record, 8) != 0 || Read(record, 12) != 5)
                    {
                        throw new NotSupportedException("EMF region clipping is unsupported.");
                    }
                    ResetClip();
                    break;
                case 88:
                case 89:
                    var count = Read(record, 24);
                    if (count < 0 || count > (size - 28) / 4 || (type == 88 && count % 3 != 0) || count > 1000000 - points)
                    {
                        throw new InvalidDataException("Invalid EMF point array or point limit exceeded.");
                    }
                    points += count;
                    var mapped = new SKPoint[count];
                    for (var index = 0; index < count; index++)
                    {
                        mapped[index] = Map(BinaryPrimitives.ReadInt16LittleEndian(record[(28 + index * 4)..]), BinaryPrimitives.ReadInt16LittleEndian(record[(30 + index * 4)..]));
                    }
                    for (var index = 0; index < count; index += type == 88 ? 3 : 1)
                    {
                        if (type == 88)
                        {
                            path.CubicTo(mapped[index], mapped[index + 1], mapped[index + 2]);
                        }
                        else
                        {
                            path.LineTo(mapped[index]);
                        }
                    }
                    break;
                case 70:
                    // GDI comments carry non-rendering metadata or alternate representations.
                    var length = Read(record, 8);
                    if (length < 0 || length > size - 12)
                    {
                        throw new InvalidDataException("Invalid EMF comment length.");
                    }
                    if (length >= 4 && Read(record, 12) == 0x2b464d45)
                    {
                        throw new NotSupportedException("EMF+ drawing records are unsupported.");
                    }
                    break;
                case 13: // Brush origin has no effect on solid brushes.
                case 18: // Background mode, stretch mode and text alignment do not affect filled paths.
                case 21:
                case 22:
                    break;
                default:
                    throw new NotSupportedException("Unsupported EMF drawing record: " + type);
            }
            offset += size;
        }
        if (!ended || records != Read(bytes, 52))
        {
            throw new InvalidDataException("EMF record count or end marker is invalid.");
        }
        using var image = SKImage.FromBitmap(bitmap);
        using var png = image.Encode(SKEncodedImageFormat.Png, 100);
        return png.ToArray();
    }
}
