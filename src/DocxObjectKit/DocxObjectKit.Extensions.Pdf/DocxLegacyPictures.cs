using System.Globalization;
using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private void LegacyPictures(XElement container, double x, double available)
    {
        XNamespace vml = "urn:schemas-microsoft-com:vml";
        foreach (var source in container.Descendants(vml + "imagedata"))
        {
            var shape = source.Parent!;
            var style = ((string?)shape.Attribute("style") ?? "").Split(';')
                .Select(part => part.Split(':', 2)).Where(parts => parts.Length == 2)
                .GroupBy(parts => parts[0].Trim()).ToDictionary(group => group.Key, group => group.Last()[1].Trim());
            double Size(string name)
            {
                var text = style.GetValueOrDefault(name, "100pt");
                return text.EndsWith("pt", StringComparison.OrdinalIgnoreCase) && double.TryParse(text[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var value) && double.IsFinite(value) ? Math.Clamp(value, 1, 3000) : 100;
            }
            var picture = new XElement(WP + "inline",
                new XElement(WP + "extent", new XAttribute("cx", Size("width") * 12700), new XAttribute("cy", Size("height") * 12700)),
                new XElement(A + "blip", new XAttribute(R + "embed", (string?)source.Attribute(R + "id") ?? "")));
            var id = "legacy-" + imageIds.Count;
            imageIds[picture] = id;
            Image(picture, x, available);
            var index = page.Images.FindIndex(image => image.Id == id);
            if (index >= 0)
            {
                page.Images[index] = page.Images[index] with { ReadOnly = true, Name = "Embedded object preview" };
            }
        }
    }
}
