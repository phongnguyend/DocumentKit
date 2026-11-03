using System.Xml;
using System.Xml.Linq;

namespace XlsxObjectKit.Models;

public sealed class XlsxObjectDocument
{
    private readonly Dictionary<string, byte[]> parts;

    internal XlsxObjectDocument(Dictionary<string, byte[]> parts) => this.parts = parts;

    public IReadOnlyCollection<string> PartNames => parts.Keys;

    public byte[] GetPart(string name) => parts.TryGetValue(name, out var value) ? value.ToArray() : throw new InvalidDataException($"Missing part: {name}");

    public XDocument ReadXml(string name)
    {
        using var stream = new MemoryStream(GetPart(name));
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16 * 1024 * 1024 });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    public XlsxObjectDocument WithPart(string name, byte[] bytes)
    {
        var copy = parts.ToDictionary(item => item.Key, item => item.Value.ToArray(), StringComparer.Ordinal);
        copy[name] = bytes.ToArray();
        return new(copy);
    }

    public XlsxObjectDocument WithXml(string name, XDocument xml) => WithPart(name, System.Text.Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
}

public static class XlsxNamespaces
{
    public static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static string Resolve(string part, string target)
    {
        if (target.Contains('\\') || target.Contains(':'))
        {
            throw new InvalidDataException("Invalid relationship target.");
        }
        var uri = new Uri(new Uri("https://package/" + part), target);
        if (uri.Host != "package")
        {
            throw new InvalidDataException("External parts are not loaded.");
        }
        return Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/'));
    }
}
