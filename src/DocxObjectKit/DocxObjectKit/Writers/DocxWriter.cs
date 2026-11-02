using System.IO.Compression;
using DocxObjectKit.Models;

namespace DocxObjectKit.Writers;

public static class DocxWriter
{
    public static byte[] Write(DocxObjectDocument document)
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
