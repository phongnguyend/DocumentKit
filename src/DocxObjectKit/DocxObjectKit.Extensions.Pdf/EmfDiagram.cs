using System.Buffers.Binary;
using System.Numerics;
using System.Text;
using SkiaSharp;

namespace DocxObjectKit.Extensions.Pdf;

// Portable GDI record playback for embedded diagram previews. Unknown drawing operations fail closed.
internal static class EmfDiagram
{
    private sealed record Pen(SKColor Color, float Width, int Style);
    private sealed record Font(string Family, float Size, bool Bold, bool Italic);
    private sealed record State
    {
        public Matrix3x2 World { get; set; } = Matrix3x2.Identity;

        public Vector2 Window { get; set; }

        public Vector2 Viewport { get; set; }

        public Vector2 WindowSize { get; set; } = Vector2.One;

        public Vector2 ViewportSize { get; set; } = Vector2.One;

        public SKRect? Clip { get; set; }

        public SKColor? Brush { get; set; } = SKColors.White;

        public Pen? Pen { get; set; } = new(SKColors.Black, 1, 0);

        public Font Font { get; set; } = new("Arial", 12, false, false);

        public SKColor TextColor { get; set; } = SKColors.Black;

        public SKColor Background { get; set; } = SKColors.White;

        public int BackgroundMode { get; set; } = 1;

        public int Align { get; set; }

        public Vector2 Position { get; set; }

        public SKPathFillType Fill { get; set; } = SKPathFillType.EvenOdd;
    }

    private static int I(ReadOnlySpan<byte> data, int offset)
    {
        if (offset < 0 || offset > data.Length - 4)
        {
            throw new InvalidDataException("Truncated EMF diagram record.");
        }
        return BinaryPrimitives.ReadInt32LittleEndian(data[offset..]);
    }

    private static SKColor Color(int value) => new((byte)value, (byte)(value >> 8), (byte)(value >> 16));

    private static Matrix3x2 Matrix(ReadOnlySpan<byte> data)
    {
        var values = new float[6];
        for (var index = 0; index < values.Length; index++)
        {
            values[index] = BitConverter.Int32BitsToSingle(I(data, index * 4));
        }
        if (values.Any(value => !float.IsFinite(value) || Math.Abs(value) > 10000000))
        {
            throw new InvalidDataException("Invalid EMF transform.");
        }
        return new(values[0], values[1], values[2], values[3], values[4], values[5]);
    }

    private static SKMatrix Skia(Matrix3x2 value) => new(value.M11, value.M21, value.M31, value.M12, value.M22, value.M32, 0, 0, 1);

    public static byte[] Decode(byte[] bytes, double width, double height)
    {
        if (!EmfImage.IsEmf(bytes) || I(bytes, 48) != bytes.Length || !double.IsFinite(width) || !double.IsFinite(height))
        {
            throw new InvalidDataException("Invalid EMF diagram header.");
        }
        var bounds = new SKRect(I(bytes, 8), I(bytes, 12), (float)I(bytes, 16) + 1, (float)I(bytes, 20) + 1);
        if (bounds.Width <= 0 || bounds.Height <= 0)
        {
            throw new InvalidDataException("Invalid EMF diagram bounds.");
        }
        using var bitmap = new SKBitmap((int)Math.Clamp(width * 3, 32, 4096), (int)Math.Clamp(height * 3, 32, 4096));
        using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.Transparent);
        var baseMatrix = Matrix3x2.CreateTranslation(-bounds.Left, -bounds.Top) * Matrix3x2.CreateScale(bitmap.Width / bounds.Width, bitmap.Height / bounds.Height);
        var state = new State();
        var stack = new List<State>();
        var objects = new Dictionary<int, object>();
        var offset = 0;
        var records = 0;
        var ended = false;
        var dualMode = false;
        Matrix3x2 Mapping() => state.World * Matrix3x2.CreateTranslation(-state.Window) * Matrix3x2.CreateScale(state.ViewportSize / state.WindowSize) * Matrix3x2.CreateTranslation(state.Viewport);
        void Begin()
        {
            canvas.Save();
            canvas.SetMatrix(Skia(baseMatrix));
            if (state.Clip is { } clip)
            {
                canvas.ClipRect(clip);
            }
            canvas.SetMatrix(Skia(Mapping() * baseMatrix));
        }
        void Draw(SKPath path, bool filled)
        {
            Begin();
            path.FillType = state.Fill;
            if (filled && state.Brush is { } brush)
            {
                using var paint = new SKPaint { Color = brush, IsAntialias = true };
                canvas.DrawPath(path, paint);
            }
            if (state.Pen is { } pen)
            {
                using var paint = new SKPaint { Color = pen.Color, Style = SKPaintStyle.Stroke, StrokeWidth = Math.Max(.1f, pen.Width), IsAntialias = true,
                    StrokeCap = (pen.Style & 0xf00) == 0x200 ? SKStrokeCap.Square : (pen.Style & 0xf00) == 0x100 ? SKStrokeCap.Butt : SKStrokeCap.Round,
                    StrokeJoin = (pen.Style & 0xf000) == 0x2000 ? SKStrokeJoin.Miter : SKStrokeJoin.Round };
                canvas.DrawPath(path, paint);
            }
            canvas.Restore();
        }
        while (offset < bytes.Length)
        {
            var type = I(bytes, offset);
            var size = I(bytes, offset + 4);
            if (size < 8 || size % 4 != 0 || size > bytes.Length - offset || ++records > 100000)
            {
                throw new InvalidDataException("Invalid EMF diagram record size.");
            }
            var r = bytes.AsSpan(offset, size);
            switch (type)
            {
                case 1:
                    break;
                case 14:
                    ended = offset + size == bytes.Length;
                    break;
                case 9:
                    state.WindowSize = new(I(r, 8), I(r, 12));
                    if (state.WindowSize.X == 0 || state.WindowSize.Y == 0)
                    {
                        throw new InvalidDataException("Zero EMF window size.");
                    }
                    break;
                case 10:
                    state.Window = new(I(r, 8), I(r, 12));
                    break;
                case 11:
                    state.ViewportSize = new(I(r, 8), I(r, 12));
                    break;
                case 12:
                    state.Viewport = new(I(r, 8), I(r, 12));
                    break;
                case 17:
                    if (I(r, 8) is not (1 or 8))
                    {
                        throw new NotSupportedException("Unsupported EMF diagram mapping mode.");
                    }
                    break;
                case 18:
                    state.BackgroundMode = I(r, 8);
                    break;
                case 19:
                    state.Fill = I(r, 8) == 2 ? SKPathFillType.Winding : SKPathFillType.EvenOdd;
                    break;
                case 22:
                    state.Align = I(r, 8);
                    break;
                case 24:
                    state.TextColor = Color(I(r, 8));
                    break;
                case 25:
                    state.Background = Color(I(r, 8));
                    break;
                case 27:
                    state.Position = new(I(r, 8), I(r, 12));
                    break;
                case 30:
                    if (Mapping().M12 != 0 || Mapping().M21 != 0)
                    {
                        throw new NotSupportedException("Rotated EMF clipping rectangles are unsupported.");
                    }
                    var p1 = Vector2.Transform(new(I(r, 8), I(r, 12)), Mapping());
                    var p2 = Vector2.Transform(new(I(r, 16), I(r, 20)), Mapping());
                    var intersection = new SKRect(Math.Min(p1.X, p2.X), Math.Min(p1.Y, p2.Y), Math.Max(p1.X, p2.X), Math.Max(p1.Y, p2.Y));
                    if (state.Clip is { } oldClip)
                    {
                        intersection = new(Math.Max(intersection.Left, oldClip.Left), Math.Max(intersection.Top, oldClip.Top), Math.Min(intersection.Right, oldClip.Right), Math.Min(intersection.Bottom, oldClip.Bottom));
                        if (intersection.Width <= 0 || intersection.Height <= 0)
                        {
                            intersection = SKRect.Empty;
                        }
                    }
                    state.Clip = intersection;
                    break;
                case 33:
                    if (stack.Count >= 1024)
                    {
                        throw new InvalidDataException("Excessive EMF state depth.");
                    }
                    stack.Add(state with { });
                    break;
                case 34:
                    var restore = I(r, 8);
                    var target = restore < 0 ? stack.Count + restore : restore - 1;
                    if (target < 0 || target >= stack.Count)
                    {
                        throw new InvalidDataException("Invalid EMF restored state.");
                    }
                    state = stack[target];
                    stack.RemoveRange(target, stack.Count - target);
                    break;
                case 35:
                    state.World = Matrix(r[8..]);
                    break;
                case 36:
                    var transform = Matrix(r[8..]);
                    state.World = I(r, 32) switch
                    {
                        1 => Matrix3x2.Identity,
                        2 => transform * state.World,
                        3 => state.World * transform,
                        _ => throw new NotSupportedException("Unsupported EMF world transform mode.")
                    };
                    break;
                case 39:
                    objects[I(r, 8)] = I(r, 12) switch
                    {
                        0 => Color(I(r, 16)),
                        1 => SKColors.Transparent,
                        _ => throw new NotSupportedException("Unsupported EMF diagram brush.")
                    };
                    break;
                case 95:
                    var penStyle = I(r, 28);
                    if ((penStyle & 15) is not (0 or 5) || I(r, 36) != 0)
                    {
                        throw new NotSupportedException("Unsupported EMF diagram pen.");
                    }
                    objects[I(r, 8)] = new Pen(Color(I(r, 40)), I(r, 32), penStyle);
                    break;
                case 82:
                    if (size < 104 || I(r, 20) != 0 || I(r, 24) != 0)
                    {
                        throw new NotSupportedException("Unsupported EMF rotated font.");
                    }
                    objects[I(r, 8)] = new Font(Encoding.Unicode.GetString(r.Slice(40, 64)).Split('\0')[0], (float)Math.Clamp(Math.Abs((double)I(r, 12)), 1, 10000), I(r, 28) >= 600, r[32] != 0);
                    break;
                case 37:
                    var handle = I(r, 8);
                    if (handle < 0)
                    {
                        switch (handle & int.MaxValue)
                        {
                            case 0:
                                state.Brush = SKColors.White;
                                break;
                            case 4:
                                state.Brush = SKColors.Black;
                                break;
                            case 5:
                                state.Brush = null;
                                break;
                            case 6:
                                state.Pen = new(SKColors.White, 1, 0);
                                break;
                            case 7:
                                state.Pen = new(SKColors.Black, 1, 0);
                                break;
                            case 8:
                                state.Pen = null;
                                break;
                            case 10:
                            case 11:
                            case 12:
                            case 13:
                            case 14:
                            case 16:
                            case 17:
                                state.Font = new("Arial", 12, false, false);
                                break;
                            default:
                                throw new NotSupportedException("Unsupported EMF stock object.");
                        }
                    }
                    else if (objects.TryGetValue(handle, out var selected))
                    {
                        switch (selected)
                        {
                            case SKColor brush:
                                state.Brush = brush.Alpha == 0 ? null : brush;
                                break;
                            case Pen pen:
                                state.Pen = (pen.Style & 15) == 5 ? null : pen;
                                break;
                            case Font font:
                                state.Font = font;
                                break;
                        }
                    }
                    else
                    {
                        throw new InvalidDataException("Unknown EMF object handle.");
                    }
                    break;
                case 40:
                    objects.Remove(I(r, 8));
                    break;
                case 43:
                    using (var rectangle = new SKPath())
                    {
                        rectangle.AddRect(new(I(r, 8), I(r, 12), I(r, 16), I(r, 20)));
                        Draw(rectangle, true);
                    }
                    break;
                case 86:
                case 87:
                    var count = I(r, 24);
                    if (count < 1 || count > (size - 28) / 4)
                    {
                        throw new InvalidDataException("Invalid EMF polygon.");
                    }
                    using (var polygon = new SKPath())
                    {
                        for (var i = 0; i < count; i++)
                        {
                            var point = new SKPoint(BinaryPrimitives.ReadInt16LittleEndian(r[(28 + i * 4)..]), BinaryPrimitives.ReadInt16LittleEndian(r[(30 + i * 4)..]));
                            if (i == 0)
                            {
                                polygon.MoveTo(point);
                            }
                            else
                            {
                                polygon.LineTo(point);
                            }
                        }
                        if (type == 86)
                        {
                            polygon.Close();
                        }
                        Draw(polygon, type == 86);
                    }
                    break;
                case 75:
                    if (I(r, 12) != 5)
                    {
                        throw new NotSupportedException("Unsupported EMF region combination.");
                    }
                    if (I(r, 8) == 0)
                    {
                        state.Clip = null;
                    }
                    else if (size >= 64 && I(r, 8) == 48 && I(r, 16) == 32 && I(r, 20) == 1 && I(r, 24) == 1)
                    {
                        state.Clip = new(I(r, 48), I(r, 52), I(r, 56), I(r, 60));
                    }
                    else
                    {
                        throw new NotSupportedException("Unsupported EMF compound region.");
                    }
                    break;
                case 81:
                    if (I(r, 68) != 0xcc0020 || I(r, 64) != 0)
                    {
                        throw new NotSupportedException("Unsupported EMF bitmap raster operation.");
                    }
                    DrawBitmap(canvas, r, state, Mapping() * baseMatrix, baseMatrix);
                    break;
                case 84:
                    DrawText(canvas, r, state, Mapping() * baseMatrix, baseMatrix);
                    break;
                case 76:
                    if (I(r, 40) != 0xaa0029)
                    {
                        throw new NotSupportedException("Unsupported EMF bit-block operation.");
                    }
                    // DSTCOPY leaves the destination unchanged.
                    break;
                case 70:
                    if (I(r, 8) < 0 || I(r, 8) > size - 12)
                    {
                        throw new InvalidDataException("Invalid EMF comment length.");
                    }
                    if (I(r, 8) >= 4 && I(r, 12) == 0x2b464d45)
                    {
                        if (!dualMode)
                        {
                            // MS-EMFPLUS: Dual headers guarantee a complete parallel EMF record stream.
                            dualMode = records == 2 && I(r, 8) >= 32 && size >= 44 && I(r, 20) == 28 && I(r, 24) == 16 && (I(r, 16) & 0xffff) == 0x4001 && ((I(r, 16) >> 16) & 1) == 1;
                            if (!dualMode)
                            {
                                throw new NotSupportedException("EMF+ Only diagrams require an EMF+ renderer.");
                            }
                        }
                    }
                    break;
                case 28:
                    if (state.Clip is { } offsetClip)
                    {
                        var delta = Vector2.TransformNormal(new(I(r, 8), I(r, 12)), Mapping());
                        offsetClip.Offset(delta.X, delta.Y);
                        state.Clip = offsetClip;
                    }
                    break;
                case 13:
                case 21:
                case 58:
                case 98:
                    break;
                default:
                    throw new NotSupportedException("Unsupported EMF diagram record: " + type);
            }
            offset += size;
        }
        if (!ended || records != I(bytes, 52))
        {
            throw new InvalidDataException("Invalid EMF diagram termination.");
        }
        using var result = SKImage.FromBitmap(bitmap);
        using var png = result.Encode(SKEncodedImageFormat.Png, 100);
        return png.ToArray();
    }

    private static void DrawBitmap(SKCanvas canvas, ReadOnlySpan<byte> r, State state, Matrix3x2 matrix, Matrix3x2 baseMatrix)
    {
        var infoOffset = I(r, 48);
        var infoSize = I(r, 52);
        var bitsOffset = I(r, 56);
        var bitsSize = I(r, 60);
        if (infoOffset < 80 || infoSize < 40 || infoSize > r.Length - infoOffset || bitsOffset < 80 || bitsSize < 0 || bitsSize > r.Length - bitsOffset)
        {
            throw new InvalidDataException("Invalid EMF bitmap bounds.");
        }
        var bmp = new byte[checked(14 + infoSize + bitsSize)];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(2), bmp.Length);
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(10), 14 + infoSize);
        r.Slice(infoOffset, infoSize).CopyTo(bmp.AsSpan(14));
        r.Slice(bitsOffset, bitsSize).CopyTo(bmp.AsSpan(14 + infoSize));
        var bitmapWidth = I(r, infoOffset + 4);
        var bitmapHeight = Math.Abs((long)I(r, infoOffset + 8));
        if (bitmapWidth <= 0 || bitmapHeight <= 0 || bitmapWidth * bitmapHeight > 64000000)
        {
            throw new InvalidDataException("Excessive EMF bitmap dimensions.");
        }
        using var image = SKImage.FromEncodedData(bmp) ?? throw new NotSupportedException("Unsupported EMF bitmap encoding.");
        canvas.Save();
        canvas.SetMatrix(Skia(baseMatrix));
        if (state.Clip is { } clip)
        {
            canvas.ClipRect(clip);
        }
        canvas.SetMatrix(Skia(matrix));
        var source = SKRect.Create(I(r, 32), I(r, 36), I(r, 40), I(r, 44));
        var dest = SKRect.Create(I(r, 24), I(r, 28), I(r, 72), I(r, 76));
        canvas.DrawImage(image, source, dest, new SKSamplingOptions(SKFilterMode.Linear));
        canvas.Restore();
    }

    private static void DrawText(SKCanvas canvas, ReadOnlySpan<byte> r, State state, Matrix3x2 matrix, Matrix3x2 baseMatrix)
    {
        var count = I(r, 44);
        var start = I(r, 48);
        var flags = I(r, 52);
        var advances = I(r, 72);
        if (count < 0 || count > 100000 || start < 76 || count > (r.Length - start) / 2 || (flags & ~(2 | 4 | 16)) != 0 || (state.Align & 1) != 0)
        {
            throw new NotSupportedException("Unsupported or invalid EMF text record.");
        }
        if (advances != 0 && (advances < 76 || count > (r.Length - advances) / 4))
        {
            throw new InvalidDataException("Invalid EMF text advances.");
        }
        var glyphIndices = (flags & 16) != 0;
        var characters = new char[count];
        for (var index = 0; index < count; index++)
        {
            characters[index] = (char)BinaryPrimitives.ReadUInt16LittleEndian(r[(start + index * 2)..]);
        }
        var text = new string(characters);
        using var face = SKFontManager.Default.MatchFamily(state.Font.Family, new SKFontStyle(state.Font.Bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal, SKFontStyleWidth.Normal, state.Font.Italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright)) ?? SKTypeface.FromFamilyName("DejaVu Sans");
        // Calibri's glyph 3 is the unpainted space; other glyph IDs cannot be mapped to a substitute font.
        var spaceOnly = glyphIndices && state.Font.Family == "Calibri" && text.All(character => character == 3);
        if (glyphIndices && !spaceOnly && (!string.Equals(face.FamilyName, state.Font.Family, StringComparison.OrdinalIgnoreCase) || advances == 0))
        {
            throw new NotSupportedException("EMF glyph indices require the original font and explicit advances.");
        }
        using var font = new SKFont(face, state.Font.Size);
        using var paint = new SKPaint { Color = state.TextColor, IsAntialias = true };
        var totalWidth = advances == 0 ? font.MeasureText(text) : 0;
        if (advances != 0)
        {
            for (var index = 0; index < count; index++)
            {
                totalWidth += I(r, advances + index * 4);
            }
        }
        var x = (float)I(r, 36);
        var y = (float)I(r, 40);
        x -= (state.Align & 6) == 6 ? totalWidth / 2 : (state.Align & 2) != 0 ? totalWidth : 0;
        y -= (state.Align & 24) == 24 ? 0 : (state.Align & 8) != 0 ? font.Metrics.Descent : font.Metrics.Ascent;
        canvas.Save();
        canvas.SetMatrix(Skia(baseMatrix));
        if (state.Clip is { } clip)
        {
            canvas.ClipRect(clip);
        }
        canvas.SetMatrix(Skia(matrix));
        var rectangle = new SKRect(I(r, 56), I(r, 60), I(r, 64), I(r, 68));
        if ((flags & 4) != 0)
        {
            canvas.ClipRect(rectangle);
        }
        if ((flags & 2) != 0 || state.BackgroundMode == 2)
        {
            using var background = new SKPaint { Color = state.Background };
            canvas.DrawRect((flags & 2) != 0 ? rectangle : new SKRect(x, y + font.Metrics.Ascent, x + totalWidth, y + font.Metrics.Descent), background);
        }
        if (advances == 0)
        {
            canvas.DrawText(text, x, y, font, paint);
        }
        else
        {
            for (var i = 0; i < count; i++)
            {
                if (!glyphIndices)
                {
                    canvas.DrawText(text.Substring(i, 1), x, y, font, paint);
                }
                else if (!spaceOnly)
                {
                    using var glyph = font.GetGlyphPath(text[i]);
                    if (glyph is not null)
                    {
                        canvas.Save();
                        canvas.Translate(x, y);
                        canvas.DrawPath(glyph, paint);
                        canvas.Restore();
                    }
                }
                x += I(r, advances + i * 4);
            }
        }
        canvas.Restore();
    }
}
