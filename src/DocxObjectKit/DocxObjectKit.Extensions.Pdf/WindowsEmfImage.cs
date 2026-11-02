using System.Buffers.Binary;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using SkiaSharp;

namespace DocxObjectKit.Extensions.Pdf;

[SupportedOSPlatform("windows")]
internal static class WindowsEmfImage
{
    public static byte[] Decode(byte[] bytes, double width, double height)
    {
        if (!EmfImage.IsEmf(bytes) || BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(48)) != bytes.Length)
        {
            throw new InvalidDataException("Invalid EMF header.");
        }
        var offset = 0;
        var count = 0;
        while (offset < bytes.Length)
        {
            if (bytes.Length - offset < 8 || ++count > 100000)
            {
                throw new InvalidDataException("Invalid EMF record count.");
            }
            var size = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(offset + 4));
            if (size < 8 || size % 4 != 0 || size > bytes.Length - offset)
            {
                throw new InvalidDataException("Invalid EMF record length.");
            }
            offset += size;
        }
        var w = (int)Math.Clamp(Math.Ceiling(width * 3), 32, 4096);
        var h = (int)Math.Clamp(Math.Ceiling(height * 3), 32, 4096);
        var dc = CreateCompatibleDC(IntPtr.Zero);
        var metafile = IntPtr.Zero;
        var dib = IntPtr.Zero;
        var previous = IntPtr.Zero;
        try
        {
            var header = new BitmapHeader { Size = 40, Width = w, Height = -h, Planes = 1, Bits = 32 };
            dib = CreateDIBSection(dc, ref header, 0, out var pixels, IntPtr.Zero, 0);
            metafile = SetEnhMetaFileBits((uint)bytes.Length, bytes);
            if (dc == IntPtr.Zero || dib == IntPtr.Zero || metafile == IntPtr.Zero)
            {
                throw new NotSupportedException("Windows could not decode the EMF preview.");
            }
            previous = SelectObject(dc, dib);
            var data = new byte[w * h * 4];
            Array.Fill(data, (byte)255);
            Marshal.Copy(data, 0, pixels, data.Length);
            var bounds = new Rectangle { Right = w, Bottom = h };
            if (!PlayEnhMetaFile(dc, metafile, ref bounds))
            {
                throw new NotSupportedException("Windows could not render the EMF preview.");
            }
            GdiFlush();
            Marshal.Copy(pixels, data, 0, data.Length);
            for (var i = 3; i < data.Length; i += 4)
            {
                data[i] = 255;
            }
            using var bitmap = new SKBitmap(w, h, SKColorType.Bgra8888, SKAlphaType.Opaque);
            Marshal.Copy(data, 0, bitmap.GetPixels(), data.Length);
            using var image = SKImage.FromBitmap(bitmap);
            using var encoded = image.Encode(SKEncodedImageFormat.Png, 100);
            return encoded.ToArray();
        }
        finally
        {
            if (previous != IntPtr.Zero)
            {
                SelectObject(dc, previous);
            }
            if (metafile != IntPtr.Zero)
            {
                DeleteEnhMetaFile(metafile);
            }
            if (dib != IntPtr.Zero)
            {
                DeleteObject(dib);
            }
            if (dc != IntPtr.Zero)
            {
                DeleteDC(dc);
            }
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct BitmapHeader
    {
        public uint Size;
        public int Width, Height;
        public ushort Planes, Bits;
        public uint Compression, ImageSize;
        public int XResolution, YResolution;
        public uint Colors, ImportantColors;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Rectangle
    {
        public int Left, Top, Right, Bottom;
    }

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleDC(IntPtr dc);
    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateDIBSection(IntPtr dc, ref BitmapHeader header, uint usage, out IntPtr pixels, IntPtr section, uint offset);
    [DllImport("gdi32.dll")]
    private static extern IntPtr SetEnhMetaFileBits(uint size, byte[] bytes);
    [DllImport("gdi32.dll")]
    private static extern IntPtr SelectObject(IntPtr dc, IntPtr value);
    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PlayEnhMetaFile(IntPtr dc, IntPtr metafile, ref Rectangle rectangle);
    [DllImport("gdi32.dll")]
    private static extern bool DeleteEnhMetaFile(IntPtr value);
    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr value);
    [DllImport("gdi32.dll")]
    private static extern bool DeleteDC(IntPtr value);
    [DllImport("gdi32.dll")]
    private static extern bool GdiFlush();
}
