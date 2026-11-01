using System.IO.Compression;
using PptxObjectKit.Models;

namespace PptxObjectKit.Writers;

public static class PptxWriter
{
    public static byte[] Write(PptxObjectDocument document)
    {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, true))
        {
            foreach (var (name, bytes) in document.Parts)
            {
                using var output = archive.CreateEntry(name, CompressionLevel.Optimal).Open();
                output.Write(bytes);
            }
        }
        return buffer.ToArray();
    }
}
