using System.IO.Compression;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Readers;

public static class DocxReader
{
    public static DocxObjectDocument Read(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        if (bytes.Length > 50 * 1024 * 1024)
        {
            throw new InvalidDataException("Documents must be no larger than 50 MB.");
        }
        using var memory = new MemoryStream(bytes);
        using var zip = new ZipArchive(memory, ZipArchiveMode.Read);
        if (zip.Entries.Count > 10000)
        {
            throw new InvalidDataException("Too many package parts.");
        }
        long total = 0;
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries)
        {
            if (entry.FullName.EndsWith('/'))
            {
                continue;
            }
            total += entry.Length;
            if (entry.FullName.StartsWith('/') || entry.FullName.Contains('\\') || entry.FullName.Split('/').Any(segment => segment is ".." or "." or "") || entry.Length > 32 * 1024 * 1024 || total > 128 * 1024 * 1024)
            {
                throw new InvalidDataException("Invalid package path or expanded size limit exceeded.");
            }
            using var stream = entry.Open();
            var data = new byte[checked((int)entry.Length)];
            stream.ReadExactly(data);
            if (stream.ReadByte() != -1 || !parts.TryAdd(entry.FullName, data))
            {
                throw new InvalidDataException("Invalid or duplicate package part.");
            }
        }
        var document = new DocxObjectDocument(parts);
        if (document.ReadXml("word/document.xml").Root?.Element(W + "body") is null)
        {
            throw new InvalidDataException("Missing Word document body; only transitional DOCX is supported.");
        }
        var types = document.ReadXml("[Content_Types].xml");
        if (types.Descendants().Any(item => ((string?)item.Attribute("ContentType"))?.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase) == true))
        {
            throw new NotSupportedException("Macro-enabled documents are not supported.");
        }
        return document;
    }
}
