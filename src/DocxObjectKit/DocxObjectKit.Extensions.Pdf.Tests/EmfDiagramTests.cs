using System.Buffers.Binary;
using System.Text;
using DocxObjectKit.Extensions.Pdf;
using SkiaSharp;

internal static class EmfDiagramTests
{
    public static void Run(Action<bool, string> check)
    {
        var records = new List<byte[]>();
        byte[] Add(int type, params int[] values)
        {
            var record = new byte[8 + values.Length * 4];
            BinaryPrimitives.WriteInt32LittleEndian(record, type);
            BinaryPrimitives.WriteInt32LittleEndian(record.AsSpan(4), record.Length);
            for (var index = 0; index < values.Length; index++)
            {
                BinaryPrimitives.WriteInt32LittleEndian(record.AsSpan(8 + index * 4), values[index]);
            }
            records.Add(record);
            return record;
        }
        void Put(byte[] data, int offset, int value) => BinaryPrimitives.WriteInt32LittleEndian(data.AsSpan(offset), value);
        byte[] Pack()
        {
            var packed = records.SelectMany(record => record).ToArray();
            Put(packed, 48, packed.Length);
            Put(packed, 52, records.Count);
            return packed;
        }
        var header = Add(1, new int[20]);
        Put(header, 16, 99);
        Put(header, 20, 99);
        Put(header, 40, 0x464d4520);
        Put(header, 44, 0x10000);
        // Synthetic EMF+ Dual header: play the ordinary EMF stream only.
        var comment = Add(70, 32, 0x2b464d45, 0x00014001, 28, 16, unchecked((int)0xdbc01002), 1, 96, 96);
        Add(37, unchecked((int)0x80000008)); // Null pen.
        Add(39, 1, 0, 0x000000ff, 0); // Red brush.
        Add(37, 1);
        Add(33);
        Add(30, 10, 10, 40, 40);
        Add(43, 0, 0, 80, 80);
        Add(34, -1);
        Add(39, 2, 0, 0x00ff0000, 0); // Blue brush.
        Add(37, 2);
        Add(33);
        Add(36, BitConverter.SingleToInt32Bits(1), 0, 0, BitConverter.SingleToInt32Bits(1), BitConverter.SingleToInt32Bits(50), 0, 2);
        Add(43, 0, 10, 20, 30);
        Add(34, -1);
        var font = Add(82, new int[24]);
        Put(font, 8, 3);
        Put(font, 12, -12);
        Encoding.Unicode.GetBytes("DejaVu Sans\0padding").CopyTo(font, 40);
        Add(37, 3);
        Add(22, 24); // Baseline alignment.
        Add(24, 0);
        var text = Add(84, new int[23]);
        Put(text, 24, 2);
        Put(text, 28, BitConverter.SingleToInt32Bits(1));
        Put(text, 32, BitConverter.SingleToInt32Bits(1));
        Put(text, 36, 10);
        Put(text, 40, 80);
        Put(text, 44, 4);
        Put(text, 48, 76);
        Put(text, 72, 84);
        Encoding.Unicode.GetBytes("Node").CopyTo(text, 76);
        for (var index = 0; index < 4; index++)
        {
            Put(text, 84 + index * 4, 9);
        }
        var dib = Add(81, new int[32]);
        Put(dib, 24, 70);
        Put(dib, 28, 40);
        Put(dib, 40, 2);
        Put(dib, 44, 2);
        Put(dib, 48, 80);
        Put(dib, 52, 40);
        Put(dib, 56, 120);
        Put(dib, 60, 16);
        Put(dib, 68, 0xcc0020);
        Put(dib, 72, 20);
        Put(dib, 76, 20);
        Put(dib, 80, 40);
        Put(dib, 84, 2);
        Put(dib, 88, 2);
        Put(dib, 92, 1 | (24 << 16));
        Put(dib, 100, 16);
        dib[121] = dib[124] = dib[129] = dib[132] = 255;
        Add(14, 0, 0, 20);
        using var image = SKBitmap.Decode(EmfDiagram.Decode(Pack(), 100, 100));
        check(image.GetPixel(60, 60).Red == 255 && image.GetPixel(60, 60).Alpha == 255, "Portable EMF must preserve fills and clipping.");
        check(image.GetPixel(135, 60).Alpha == 0, "Portable EMF clip must exclude outside pixels.");
        check(image.GetPixel(180, 60).Blue == 255 && image.GetPixel(180, 60).Alpha == 255, "Portable EMF must apply transforms after restoring the clip.");
        check(Enumerable.Range(195, 48).Any(y => Enumerable.Range(30, 120).Any(x => image.GetPixel(x, y).Alpha > 0)), "Portable EMF must render positioned text.");
        check(image.GetPixel(240, 150).Green == 255 && image.GetPixel(240, 150).Alpha == 255, "Portable EMF must paint DIB bitmap records.");
        void Reject(string message)
        {
            try
            {
                EmfDiagram.Decode(Pack(), 100, 100);
                check(false, message);
            }
            catch (NotSupportedException)
            {
                check(true, message);
            }
            catch (InvalidDataException)
            {
                check(true, message);
            }
        }
        Put(comment, 16, 0x4001);
        Reject("EMF+ Only must not silently render an incomplete GDI stream.");
        Put(comment, 16, 0x14001);
        Put(text, 72, int.MaxValue);
        Reject("Invalid advance offsets must be rejected before reading text.");
        Put(text, 72, 84);
        Put(comment, 8, int.MaxValue);
        Reject("Invalid EMF comment sizes must be rejected.");
        Put(comment, 8, 32);
        Put(dib, 56, int.MaxValue);
        Reject("Invalid DIB offsets must be rejected before decoding.");
        Put(dib, 56, 120);
        Put(text, 0, 999);
        Reject("Unknown drawing operations must not silently disappear.");
    }
}
