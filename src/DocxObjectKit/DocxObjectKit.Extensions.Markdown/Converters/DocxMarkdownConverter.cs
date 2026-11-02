using System.Text;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using DocxObjectKit.Models;
using DocxObjectKit.Editors;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Markdown;

public static class DocxMarkdownConverter
{
    public static string Convert(DocxObjectDocument document, bool images = true)
    {
        var output = new StringBuilder();
        var xml = document.ReadXml("word/document.xml");
        string Escape(string value) => Regex.Replace(value, @"([\\`*_{}\[\]<>|])", @"\$1");
        string Paragraph(XElement paragraph)
        {
            var text = new StringBuilder();
            foreach (var run in paragraph.Descendants(W + "r").Where(item => !item.Ancestors(W + "del").Any()))
            {
                var value = string.Concat(run.Elements().Select(item => item.Name == W + "t" ? Escape(item.Value) : item.Name == W + "br" ? "  \n" : item.Name == W + "tab" ? " " : ""));
                var properties = run.Element(W + "rPr");
                bool Enabled(string name) => properties?.Element(W + name) is { } item && Value(item) is not ("0" or "false" or "off");
                if (Enabled("b") && value.Length > 0)
                {
                    value = "**" + value + "**";
                }
                if (Enabled("i") && value.Length > 0)
                {
                    value = "*" + value + "*";
                }
                if (!Enabled("vanish"))
                {
                    text.Append(value);
                }
                if (images)
                {
                    foreach (var picture in run.Descendants().Where(item => item.Name == WP + "inline" || item.Name == WP + "anchor"))
                    {
                        try
                        {
                            var part = DocxEditor.ImagePart(document, picture);
                            var bytes = document.GetPart(part);
                            var mime = part.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ? "image/png" : part.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) || part.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) ? "image/jpeg" : null;
                            if (mime is not null)
                            {
                                text.Append("\n\n![Image](data:" + mime + ";base64," + System.Convert.ToBase64String(bytes) + ")\n\n");
                            }
                        }
                        catch (NotSupportedException)
                        {
                            text.Append("[External image omitted]");
                        }
                    }
                }
            }
            return text.ToString();
        }
        foreach (var block in xml.Root!.Element(W + "body")!.Elements())
        {
            if (block.Name == W + "p")
            {
                var style = Value(block.Element(W + "pPr")?.Element(W + "pStyle")) ?? "";
                var match = Regex.Match(style, @"^Heading([1-6])$", RegexOptions.IgnoreCase);
                var prefix = match.Success ? new string('#', int.Parse(match.Groups[1].Value)) + " " : block.Element(W + "pPr")?.Element(W + "numPr") is not null ? "- " : "";
                output.AppendLine(prefix + Paragraph(block)).AppendLine();
            }
            else if (block.Name == W + "tbl")
            {
                var rows = block.Elements(W + "tr").Select(row => row.Elements(W + "tc").Select(cell => string.Join(" ", cell.Elements(W + "p").Select(Paragraph)).Replace("\n", " ")).ToArray()).ToArray();
                var columns = rows.Select(row => row.Length).DefaultIfEmpty(0).Max();
                for (var i = 0; i < rows.Length; i++)
                {
                    output.AppendLine("| " + string.Join(" | ", rows[i].Concat(Enumerable.Repeat("", columns - rows[i].Length))) + " |");
                    if (i == 0)
                    {
                        output.AppendLine("| " + string.Join(" | ", Enumerable.Repeat("---", columns)) + " |");
                    }
                }
                output.AppendLine();
            }
        }
        return output.ToString();
    }
}
