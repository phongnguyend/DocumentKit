using System.Xml;
using System.Xml.Linq;

namespace DocxObjectKit.Models;

public sealed class DocxObjectDocument
{
    private readonly Dictionary<string, byte[]> parts;

    internal DocxObjectDocument(Dictionary<string, byte[]> parts) => this.parts = parts;

    public IReadOnlyCollection<string> PartNames => parts.Keys;

    public byte[] GetPart(string name) => parts.TryGetValue(name, out var value) ? value.ToArray() : throw new InvalidDataException($"Missing part: {name}");

    public XDocument ReadXml(string name)
    {
        using var stream = new MemoryStream(GetPart(name));
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 16 * 1024 * 1024 });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    public DocxObjectDocument WithPart(string name, byte[] bytes)
    {
        var copy = parts.ToDictionary(item => item.Key, item => item.Value.ToArray(), StringComparer.Ordinal);
        copy[name] = bytes.ToArray();
        return new(copy);
    }

    public DocxObjectDocument WithXml(string name, XDocument xml) => WithPart(name, System.Text.Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
}

public static class DocxNamespaces
{
    public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    public static readonly XNamespace PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    public static readonly XNamespace REL = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static string? Value(XElement? element, string name = "val") => (string?)element?.Attribute(W + name);

    public static double Number(XElement? element, string name, double fallback = 0) => double.TryParse((string?)element?.Attribute(name) ?? Value(element, name), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var value) && double.IsFinite(value) ? value : fallback;

    public static string Resolve(string part, string target)
    {
        if (target.Contains('\\') || target.Contains(':'))
        {
            throw new InvalidDataException("Invalid relationship target.");
        }
        var resolved = new Uri(new Uri("https://package/" + part), target);
        if (resolved.Host != "package")
        {
            throw new InvalidDataException("External relationship targets are not loaded.");
        }
        return Uri.UnescapeDataString(resolved.AbsolutePath.TrimStart('/'));
    }
}
