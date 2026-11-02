using System.Text;
using System.Text.RegularExpressions;
using PptxObjectKit.Extractors;
using PptxObjectKit.Graphs;
using PptxObjectKit.Models;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Extensions.Markdown;

public sealed record PptxMarkdownOptions(bool IncludeNotes = true, bool IncludeImages = true);

public static class PptxMarkdownConverter
{
    private static string Escape(string text) => Regex.Replace(text.Replace("\r", ""), @"([\\`*_{}\[\]()#+.!|<>~\-])", @"\$1");

    public static string Convert(PptxObjectDocument document, PptxMarkdownOptions? options = null)
    {
        options ??= new();
        var graph = new PptxPackageGraph(document);
        var result = new StringBuilder();
        foreach (var slide in PptxSlideExtractor.Extract(document).Slides)
        {
            if (result.Length > 0)
            {
                result.AppendLine("---").AppendLine();
            }
            result.AppendLine($"# Slide {slide.Number}: {Escape(slide.Title)}").AppendLine();
            var xml = document.Xml(slide.PartName);
            foreach (var body in xml.Descendants(P + "txBody"))
            {
                foreach (var paragraph in body.Elements(A + "p"))
                {
                    var text = PptxSlideExtractor.Text(new System.Xml.Linq.XElement(P + "txBody", new System.Xml.Linq.XElement(paragraph)));
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        var bullet = paragraph.Element(A + "pPr")?.Elements().Any(element => element.Name == A + "buChar" || element.Name == A + "buAutoNum") == true;
                        result.AppendLine((bullet ? "- " : "") + Escape(text)).AppendLine();
                    }
                }
            }
            foreach (var table in xml.Descendants(A + "tbl"))
            {
                var rows = table.Elements(A + "tr").Select(row => row.Elements(A + "tc")
                    .Select(cell => Escape(string.Join(" ", cell.Descendants(A + "t").Select(value => value.Value)))).ToList()).ToList();
                var columns = rows.Select(row => row.Count).DefaultIfEmpty().Max();
                if (columns > 0)
                {
                    result.AppendLine("| " + string.Join(" | ", Enumerable.Repeat("", columns)) + " |");
                    result.AppendLine("| " + string.Join(" | ", Enumerable.Repeat("---", columns)) + " |");
                    foreach (var row in rows)
                    {
                        result.AppendLine("| " + string.Join(" | ", row.Concat(Enumerable.Repeat("", columns - row.Count))) + " |");
                    }
                    result.AppendLine();
                }
            }
            if (options.IncludeImages)
            {
                var relations = graph.Relationships(slide.PartName).ToDictionary(rel => (string)rel.Attribute("Id")!);
                foreach (var picture in xml.Descendants(P + "pic"))
                {
                    var id = (string?)picture.Descendants(A + "blip").FirstOrDefault()?.Attribute(R + "embed");
                    if (id is not null && relations.TryGetValue(id, out var rel))
                    {
                        var path = Target(slide.PartName, rel);
                        var extension = Path.GetExtension(path).ToLowerInvariant();
                        var mime = extension switch { ".png" => "image/png", ".jpg" or ".jpeg" => "image/jpeg", _ => null };
                        if (mime is not null)
                        {
                            result.AppendLine($"![Slide image](data:{mime};base64,{System.Convert.ToBase64String(document.GetPart(path))})").AppendLine();
                        }
                    }
                }
            }
            if (options.IncludeNotes && !string.IsNullOrWhiteSpace(slide.Notes))
            {
                result.AppendLine("## Speaker notes").AppendLine().AppendLine(Escape(slide.Notes)).AppendLine();
            }
        }
        return result.ToString();
    }
}
