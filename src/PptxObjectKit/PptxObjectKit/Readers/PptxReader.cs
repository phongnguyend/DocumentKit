using System.IO.Compression;
using PptxObjectKit.Graphs;
using PptxObjectKit.Models;

namespace PptxObjectKit.Readers;

public static class PptxReader
{
    public static PptxObjectDocument Read(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        if (bytes.Length > 50 * 1024 * 1024)
        {
            throw new InvalidDataException("Presentations must be no larger than 50 MB.");
        }
        using var input = new MemoryStream(bytes);
        using var zip = new ZipArchive(input, ZipArchiveMode.Read);
        if (zip.Entries.Count > 10000)
        {
            throw new InvalidDataException("Too many package parts.");
        }
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        long total = 0;
        foreach (var entry in zip.Entries)
        {
            if (entry.FullName.EndsWith('/'))
            {
                continue;
            }
            if (entry.FullName.StartsWith('/') || entry.FullName.Contains('\\') || entry.FullName.Split('/').Any(segment => segment is ".." or "." or ""))
            {
                throw new InvalidDataException("Invalid package part path.");
            }
            total += entry.Length;
            if (entry.Length > 32 * 1024 * 1024 || total > 128 * 1024 * 1024)
            {
                throw new InvalidDataException("Expanded presentation exceeds package limits.");
            }
            using var stream = entry.Open();
            var data = new byte[checked((int)entry.Length)];
            stream.ReadExactly(data);
            if (stream.ReadByte() != -1 || !parts.TryAdd(entry.FullName, data))
            {
                throw new InvalidDataException("Invalid or duplicate package part.");
            }
        }
        var document = new PptxObjectDocument(parts);
        _ = document.Xml("[Content_Types].xml");
        _ = new PptxPackageGraph(document).Slides();
        return document;
    }
}
