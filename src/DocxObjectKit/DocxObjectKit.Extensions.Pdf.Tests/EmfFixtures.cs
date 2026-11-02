using System.Buffers.Binary;

internal static class EmfFixtures
{
    public static byte[] Create(bool clip = false, bool curve = false, bool unsupported = false)
    {
        var records = new List<byte[]>();
        void Record(int type, params int[] data)
        {
            var bytes = new byte[8 + data.Length * 4];
            BinaryPrimitives.WriteInt32LittleEndian(bytes, type);
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(4), bytes.Length);
            for (var index = 0; index < data.Length; index++)
            {
                BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(8 + index * 4), data[index]);
            }
            records.Add(bytes);
        }
        int Point(short x, short y) => (ushort)x | y << 16;
        void Rectangle(short end)
        {
            Record(59);
            Record(27, 10, 10);
            Record(89, 0, 0, 99, 99, 3, Point(end, 10), Point(end, 90), Point(10, 90));
            Record(61);
            Record(60);
        }
        Record(1, new int[20]);
        Record(17, 8);
        Record(9, 100, 100);
        Record(11, 200, 200);
        Record(39, 1, 0, 0x003322dd, 0); // COLORREF: red DD, green 22, blue 33.
        Record(37, 1);
        if (clip)
        {
            Rectangle(50);
            Record(67, 5);
        }
        Rectangle(90);
        Record(62, 0, 0, 99, 99);
        if (curve)
        {
            Record(59);
            Record(27, 10, 50);
            Record(88, 0, 0, 99, 99, 3, Point(20, 0), Point(80, 0), Point(90, 50));
            Record(61);
            Record(60);
            Record(62, 0, 0, 99, 99);
        }
        if (unsupported)
        {
            Record(43, 0, 0, 10, 10);
        }
        Record(14, 0, 0, 20);
        var result = records.SelectMany(record => record).ToArray();
        BinaryPrimitives.WriteInt32LittleEndian(result.AsSpan(16), 199);
        BinaryPrimitives.WriteInt32LittleEndian(result.AsSpan(20), 199);
        BinaryPrimitives.WriteInt32LittleEndian(result.AsSpan(40), 0x464d4520);
        BinaryPrimitives.WriteInt32LittleEndian(result.AsSpan(44), 0x10000);
        BinaryPrimitives.WriteInt32LittleEndian(result.AsSpan(48), result.Length);
        BinaryPrimitives.WriteInt32LittleEndian(result.AsSpan(52), records.Count);
        return result;
    }
}
