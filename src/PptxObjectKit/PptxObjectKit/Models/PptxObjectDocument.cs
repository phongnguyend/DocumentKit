using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace PptxObjectKit.Models;

/// <summary>A package-preserving presentation. Unmodified parts retain their original bytes.</summary>
public sealed class PptxObjectDocument
{
    internal Dictionary<string, byte[]> Parts { get; }

    internal PptxObjectDocument(Dictionary<string, byte[]> parts) => Parts = parts;

    public IReadOnlyCollection<string> PartNames => Parts.Keys;

    public byte[] GetPart(string name) => Parts.TryGetValue(name, out var bytes)
        ? bytes.ToArray() : throw new InvalidDataException($"Missing package part: {name}");

    internal XDocument Xml(string name)
    {
        using var input = new MemoryStream(GetPart(name));
        using var reader = XmlReader.Create(input, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 16 * 1024 * 1024
        });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    internal void SetXml(string name, XDocument xml) => Parts[name] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));

    public PptxObjectDocument Clone() => new(Parts.ToDictionary(pair => pair.Key, pair => pair.Value.ToArray(), StringComparer.Ordinal));
}

public sealed record PptxTextObject(string Id, string Name, string Text);

public sealed record PptxSlide(int Number, string Id, string PartName, string Title, List<PptxTextObject> TextObjects, int PictureCount, string Notes);

public sealed record PptxInspection(double Width, double Height, List<PptxSlide> Slides, List<string> Warnings);
