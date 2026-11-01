using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace PdfObjectKit.Converters;

internal static class PptxPng
{
    internal static byte[] Encode(byte[] rgba, int width, int height)
    {
        if (width <= 0 || height <= 0 || (long)width * height > 16000000 || rgba.Length != (long)width * height * 4)
        {
            throw new InvalidDataException("Invalid RGBA image.");
        }
        using var output = new MemoryStream();
        output.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        void Chunk(string type, byte[] data)
        {
            var header = new byte[4];
            BinaryPrimitives.WriteInt32BigEndian(header, data.Length);
            output.Write(header);
            var tag = Encoding.ASCII.GetBytes(type);
            output.Write(tag);
            output.Write(data);
            var crc = uint.MaxValue;
            foreach (var value in tag.Concat(data))
            {
                crc ^= value;
                for (var bit = 0; bit < 8; bit++)
                {
                    crc = (crc >> 1) ^ ((crc & 1) != 0 ? 0xedb88320U : 0U);
                }
            }
            BinaryPrimitives.WriteUInt32BigEndian(header, ~crc);
            output.Write(header);
        }
        var ihdr = new byte[13];
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(0, 4), width);
        BinaryPrimitives.WriteInt32BigEndian(ihdr.AsSpan(4, 4), height);
        ihdr[8] = 8;
        ihdr[9] = 6;
        Chunk("IHDR", ihdr);
        using var compressed = new MemoryStream();
        using (var zlib = new ZLibStream(compressed, CompressionLevel.Optimal, true))
        {
            for (var row = 0; row < height; row++)
            {
                zlib.WriteByte(0);
                zlib.Write(rgba, row * width * 4, width * 4);
            }
        }
        Chunk("IDAT", compressed.ToArray());
        Chunk("IEND", []);
        return output.ToArray();
    }
}
