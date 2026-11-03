using System.IO.Compression;
using XlsxObjectKit.Models;

namespace XlsxObjectKit.Writers;

public static class XlsxWriter
{
    public static byte[] Write(XlsxObjectDocument document)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
        {
            foreach (var name in document.PartNames)
            {
                using var stream = zip.CreateEntry(name, CompressionLevel.Optimal).Open();
                stream.Write(document.GetPart(name));
            }
        }
        return output.ToArray();
    }
}
