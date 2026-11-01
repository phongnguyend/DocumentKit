using System.Buffers.Binary;
using System.Text;
using PdfObjectKit.Readers;
using PdfObjectKit.Writers;

internal static class XrefPlaceholderTests
{
    public static void Run()
    {
        foreach (var stream in new[] { false, true })
        {
            var document = PdfReader.Read(Create(stream, false, false));
            if (document.Objects.Count != 2 || PdfReader.Read(PdfWriter.Write(document)).Objects.Count != 2)
            {
                throw new Exception("Unused zero-offset placeholders must be omitted and survive a round trip.");
            }
            foreach (var input in new[] { Create(stream, true, false), Create(stream, false, true) })
            {
                var rejected = false;
                try
                {
                    PdfReader.Read(input);
                }
                catch (InvalidDataException)
                {
                    rejected = true;
                }
                if (!rejected)
                {
                    throw new Exception("Referenced placeholders and nonzero object-header mismatches must be rejected.");
                }
            }
        }
        Console.WriteLine("Cross-reference placeholder and mismatch checks passed.");
    }

    private static byte[] Create(bool stream, bool referenced, bool wrongOffset)
    {
        using var output = new MemoryStream();
        void Write(string value)
        {
            output.Write(Encoding.ASCII.GetBytes(value));
        }
        Write("%PDF-1.7\n");
        var first = (int)output.Position;
        Write($"1 0 obj\n<< /Type /Catalog /Pages 2 0 R{(referenced ? " /Unused 3 0 R" : "")} >>\nendobj\n");
        var second = (int)output.Position;
        Write("2 0 obj\n<< /Type /Pages /Count 0 /Kids [] >>\nendobj\n");
        var xref = (int)output.Position;
        var offsets = new[] { 0, first, second, wrongOffset ? first : 0, xref };
        if (stream)
        {
            Write("4 0 obj\n<< /Type /XRef /Root 1 0 R /Size 5 /W [1 4 2] /Length 35 >>\nstream\n");
            for (var i = 0; i < offsets.Length; i++)
            {
                var entry = new byte[7];
                entry[0] = i == 0 ? (byte)0 : (byte)1;
                BinaryPrimitives.WriteInt32BigEndian(entry.AsSpan(1, 4), offsets[i]);
                output.Write(entry);
            }
            Write("\nendstream\nendobj\n");
        }
        else
        {
            Write("xref\n0 4\n0000000000 65535 f \n");
            foreach (var offset in offsets.Skip(1).Take(3))
            {
                Write($"{offset:0000000000} 00000 n \n");
            }
            Write("trailer\n<< /Root 1 0 R /Size 4 >>\n");
        }
        Write($"startxref\n{xref}\n%%EOF\n");
        return output.ToArray();
    }
}
