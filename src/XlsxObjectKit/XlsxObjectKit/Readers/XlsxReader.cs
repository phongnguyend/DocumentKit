using System.IO.Compression;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Readers;

public static class XlsxReader
{
    public static XlsxObjectDocument Read(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        if (bytes.Length > 50 * 1024 * 1024)
        {
            throw new InvalidDataException("Documents must be no larger than 50 MB.");
        }
        if (bytes.Length == 0)
        {
            throw new InvalidDataException("The uploaded workbook is empty. Choose the original XLSX file and try again.");
        }
        if (bytes.AsSpan().StartsWith(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }))
        {
            throw new NotSupportedException("This file uses an Office compound container, which can contain a password-encrypted or rights-managed workbook, or a legacy XLS file. Encrypted and rights-managed workbooks are not supported. Open it in Excel and, if your permissions allow, save an unprotected .xlsx copy before uploading. Renaming the extension does not convert the file.");
        }
        if (bytes.Length < 4 || bytes[0] != 'P' || bytes[1] != 'K')
        {
            throw new InvalidDataException("The uploaded file is not an XLSX ZIP package. Open the original workbook in Excel and use Save As > Excel Workbook (.xlsx); do not just rename its extension.");
        }
        using var memory = new MemoryStream(bytes);
        using var zip = OpenArchive(memory);
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
        var document = new XlsxObjectDocument(parts);
        if (document.ReadXml("xl/workbook.xml").Root?.Element(S + "sheets") is null)
        {
            throw new InvalidDataException("Missing workbook; only transitional XLSX is supported.");
        }
        var types = document.ReadXml("[Content_Types].xml");
        if (types.Descendants().Any(item => ((string?)item.Attribute("ContentType"))?.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase) == true))
        {
            throw new NotSupportedException("Macro-enabled documents are not supported.");
        }
        return document;
    }

    private static ZipArchive OpenArchive(Stream stream)
    {
        try
        {
            return new ZipArchive(stream, ZipArchiveMode.Read);
        }
        catch (InvalidDataException exception)
        {
            throw new InvalidDataException("The XLSX ZIP directory could not be read. The file may be incomplete or damaged. Download or save a fresh copy and try again.", exception);
        }
    }
}
