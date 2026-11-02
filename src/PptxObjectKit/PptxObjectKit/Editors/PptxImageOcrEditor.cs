using System.Xml.Linq;
using PptxObjectKit.Graphs;
using PptxObjectKit.Models;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Editors;

public sealed record PptxOcrLine(string Text, double X, double Y, double Width, double Height, double Baseline = 0);

public static class PptxImageOcrEditor
{
    public static PptxObjectDocument Apply(PptxObjectDocument source, int slideNumber, string objectId, IReadOnlyList<PptxOcrLine> lines, bool replaceImage = false)
    {
        if (source.PartNames.Any(path => path.StartsWith("_xmlsignatures/", StringComparison.Ordinal)))
        {
            throw new NotSupportedException("Editing signed presentations is not supported.");
        }
        if (lines.Count == 0 || lines.Count > 5000 || lines.Sum(line => (long)(line.Text?.Length ?? 0)) > 100000)
        {
            throw new ArgumentException("Provide 1 to 5000 OCR lines, with at most 100000 characters.");
        }
        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line.Text) || !new[] { line.X, line.Y, line.Width, line.Height, line.Baseline }.All(double.IsFinite) || line.X < 0 || line.Y < 0 || line.Width <= 0 || line.Height <= 0 || line.X + line.Width > 1.000001 || line.Y + line.Height > 1.000001)
            {
                throw new ArgumentException("OCR text requires finite normalized bounds inside the image.");
            }
        }
        var document = source.Clone();
        var slides = new PptxPackageGraph(document).Slides();
        if (slideNumber < 1 || slideNumber > slides.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(slideNumber));
        }
        var part = slides[slideNumber - 1].Part;
        var xml = document.Xml(part);
        var picture = xml.Descendants(P + "pic").SingleOrDefault(item => (string?)item.Element(P + "nvPicPr")?.Element(P + "cNvPr")?.Attribute("id") == objectId) ?? throw new ArgumentException("Image not found.");
        var crop = picture.Element(P + "blipFill")?.Element(A + "srcRect");
        if (crop?.Attributes().Any(attribute => double.TryParse(attribute.Value, out var value) && value != 0) == true)
        {
            throw new NotSupportedException("Applying OCR to cropped pictures is not supported yet. You can still copy the recognized text.");
        }
        var frame = picture.Element(P + "spPr")?.Element(A + "xfrm") ?? throw new NotSupportedException("Image has no explicit placement.");
        var width = (long?)frame.Element(A + "ext")?.Attribute("cx") ?? 0;
        var height = (long?)frame.Element(A + "ext")?.Attribute("cy") ?? 0;
        if (width <= 0 || height <= 0)
        {
            throw new ArgumentException("Image has invalid dimensions.");
        }
        var nextId = xml.Descendants(P + "cNvPr").Select(item => (long?)item.Attribute("id") ?? 0).DefaultIfEmpty().Max() + 1;
        var groupFrame = new XElement(frame);
        groupFrame.Add(new XElement(A + "chOff", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(A + "chExt", new XAttribute("cx", width), new XAttribute("cy", height)));
        var group = new XElement(P + "grpSp", new XElement(P + "nvGrpSpPr", new XElement(P + "cNvPr", new XAttribute("id", nextId++), new XAttribute("name", "OCR text")), new XElement(P + "cNvGrpSpPr"), new XElement(P + "nvPr")), new XElement(P + "grpSpPr", groupFrame));
        foreach (var line in lines)
        {
            var color = new XElement(A + "srgbClr", new XAttribute("val", "000000"));
            if (!replaceImage)
            {
                color.Add(new XElement(A + "alpha", new XAttribute("val", 0)));
            }
            var size = Math.Clamp((int)Math.Round(height / 12700d * line.Height * 80), 100, 400000);
            var id = nextId++;
            var transform = new XElement(A + "xfrm",
                new XElement(A + "off", new XAttribute("x", (long)Math.Round(line.X * width)), new XAttribute("y", (long)Math.Round(line.Y * height))),
                new XElement(A + "ext", new XAttribute("cx", Math.Max(1, (long)Math.Round(line.Width * width))), new XAttribute("cy", Math.Max(1, (long)Math.Round(line.Height * height)))));
            var run = new XElement(A + "r",
                new XElement(A + "rPr", new XAttribute("sz", size), new XElement(A + "solidFill", color), new XElement(A + "latin", new XAttribute("typeface", "Arial"))),
                new XElement(A + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), line.Text));
            var body = new XElement(P + "txBody",
                new XElement(A + "bodyPr", new XAttribute("wrap", "none"), new XAttribute("lIns", 0), new XAttribute("rIns", 0), new XAttribute("tIns", 0), new XAttribute("bIns", 0)),
                new XElement(A + "lstStyle"), new XElement(A + "p", run));
            group.Add(new XElement(P + "sp",
                new XElement(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", id), new XAttribute("name", "OCR line " + id)), new XElement(P + "cNvSpPr", new XAttribute("txBox", 1)), new XElement(P + "nvPr")),
                new XElement(P + "spPr", transform), body));
        }
        picture.AddAfterSelf(group);
        if (replaceImage)
        {
            picture.Remove();
        }
        document.SetXml(part, xml);
        return document;
    }
}
