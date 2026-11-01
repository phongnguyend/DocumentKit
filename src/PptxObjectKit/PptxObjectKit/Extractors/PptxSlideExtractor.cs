using System.Xml.Linq;
using PptxObjectKit.Graphs;
using PptxObjectKit.Models;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Extractors;

public static class PptxSlideExtractor
{
    internal static string Text(XElement body) => string.Join("\n", body.Elements(A + "p").Select(paragraph =>
        string.Concat(paragraph.Elements().Select(item => item.Name == A + "br" ? "\n" : item.Element(A + "t")?.Value ?? ""))));

    public static PptxInspection Extract(PptxObjectDocument document)
    {
        var graph = new PptxPackageGraph(document);
        var size = document.Xml(graph.PresentationPart()).Root!.Element(P + "sldSz");
        var width = (double?)size?.Attribute("cx") / 12700 ?? 720;
        var height = (double?)size?.Attribute("cy") / 12700 ?? 540;
        if (!double.IsFinite(width + height) || width <= 0 || height <= 0 || width > 10000 || height > 10000)
        {
            throw new InvalidDataException("Unsupported slide dimensions.");
        }
        var slides = new List<PptxSlide>();
        var warnings = new HashSet<string>
        {
            "Browser preview is approximate: fonts, themes, layouts, grouped shapes, charts, SmartArt, animations and media may differ from PowerPoint. Original package parts are preserved when saving."
        };
        foreach (var (id, path) in graph.Slides())
        {
            var xml = document.Xml(path);
            var objects = new List<PptxTextObject>();
            foreach (var shape in xml.Descendants(P + "sp"))
            {
                if (shape.Element(P + "txBody") is not { } body)
                {
                    continue;
                }
                var properties = shape.Element(P + "nvSpPr")?.Element(P + "cNvPr");
                objects.Add(new((string?)properties?.Attribute("id") ?? "", (string?)properties?.Attribute("name") ?? "Text", Text(body)));
            }
            var title = xml.Descendants(P + "sp").FirstOrDefault(shape => shape.Descendants(P + "ph").Any(ph => (string?)ph.Attribute("type") is "title" or "ctrTitle"));
            var titleText = title?.Element(P + "txBody") is { } titleBody ? Text(titleBody) : objects.FirstOrDefault()?.Text;
            var notes = "";
            var notesRelation = graph.Relationships(path).FirstOrDefault(rel => ((string?)rel.Attribute("Type"))?.EndsWith("/notesSlide", StringComparison.Ordinal) == true);
            if (notesRelation is not null)
            {
                var notesXml = document.Xml(Target(path, notesRelation));
                notes = string.Join("\n", notesXml.Descendants(P + "sp")
                    .Where(shape => shape.Descendants(P + "ph").Any(ph => (string?)ph.Attribute("type") == "body"))
                    .Elements(P + "txBody").Select(Text));
            }
            slides.Add(new(slides.Count + 1, (string?)id.Attribute("id") ?? "", path,
                string.IsNullOrWhiteSpace(titleText) ? $"Slide {slides.Count + 1}" : titleText.Split('\n')[0], objects, xml.Descendants(P + "pic").Count(), notes));
        }
        foreach (var slide in slides)
        {
            var xml = document.Xml(slide.PartName);
            foreach (var picture in xml.Descendants(P + "pic"))
            {
                var properties = picture.Element(P + "nvPicPr")?.Element(P + "cNvPr");
                var id = (string?)properties?.Attribute("id");
                var embed = (string?)picture.Element(P + "blipFill")?.Element(A + "blip")?.Attribute(R + "embed");
                var relationship = graph.Relationships(slide.PartName).FirstOrDefault(item => (string?)item.Attribute("Id") == embed && (string?)item.Attribute("TargetMode") != "External");
                if (id is null || relationship is null)
                {
                    continue;
                }
                var frame = picture.Element(P + "spPr")?.Element(A + "xfrm");
                double Value(string element, string attribute) => ((double?)frame?.Element(A + element)?.Attribute(attribute) ?? 0) / 12700;
                slide.Images.Add(new(id, (string?)properties?.Attribute("name") ?? "Image", Target(slide.PartName, relationship), Value("off", "x"), Value("off", "y"), Value("ext", "cx"), Value("ext", "cy"), picture.Ancestors(P + "grpSp").Any(), ((double?)frame?.Attribute("rot") ?? 0) / 60000));
            }
        }
        if (document.PartNames.Any(name => name.StartsWith("_xmlsignatures/", StringComparison.Ordinal)))
        {
            warnings.Add("This presentation is signed. Editing signed packages is disabled.");
        }
        return new(width, height, slides, warnings.ToList());
    }
}
