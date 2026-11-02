using System.Xml.Linq;
using System.Text.RegularExpressions;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public static class DocxTocUpdater
{
    public static DocxObjectDocument Update(DocxObjectDocument document)
    {
        var xml = document.ReadXml("word/document.xml");
        var contents = xml.Descendants(W + "sdt").Where(item => item.Descendants(W + "docPartGallery").Any(gallery => (string?)gallery.Attribute(W + "val") == "Table of Contents")
            || item.Descendants(W + "instrText").Any(instruction => Regex.IsMatch(instruction.Value, @"^\s*TOC(?:\s|$)", RegexOptions.IgnoreCase))).ToArray();
        if (contents.Length != 1)
        {
            throw new NotSupportedException("Refresh requires one TOC content control. Multiple TOCs and unwrapped TOC fields are not yet supported.");
        }
        var toc = contents[0];
        var instruction = string.Concat(toc.Descendants(W + "instrText").Select(item => item.Value));
        var range = Regex.Match(instruction, "\\\\o\\s+\"([1-9])-([1-9])\"");
        var firstLevel = range.Success ? int.Parse(range.Groups[1].Value) : 1;
        var lastLevel = range.Success ? int.Parse(range.Groups[2].Value) : 3;
        if (Regex.Matches(instruction, @"\\([a-z])", RegexOptions.IgnoreCase).Any(match => !"ohzu".Contains(match.Groups[1].Value.ToLowerInvariant(), StringComparison.Ordinal)))
        {
            throw new NotSupportedException("Custom style, bookmark-range and caption TOCs are not yet regenerated.");
        }
        var styles = document.PartNames.Contains("word/styles.xml")
            ? document.ReadXml("word/styles.xml").Descendants(W + "style").GroupBy(item => (string?)item.Attribute(W + "styleId") ?? "").ToDictionary(group => group.Key, group => group.First())
            : new Dictionary<string, XElement>();
        int Level(XElement paragraph)
        {
            var properties = paragraph.Element(W + "pPr");
            var visited = new HashSet<string>();
            var styleName = (string?)properties?.Element(W + "pStyle")?.Attribute(W + "val");
            var explicitLevel = (string?)properties?.Element(W + "outlineLvl")?.Attribute(W + "val");
            while (explicitLevel is null && styleName is not null && visited.Add(styleName) && styles.TryGetValue(styleName, out var style))
            {
                explicitLevel = (string?)style.Element(W + "pPr")?.Element(W + "outlineLvl")?.Attribute(W + "val");
                var conventional = Regex.Match(styleName, "^Heading([1-9])$", RegexOptions.IgnoreCase);
                if (explicitLevel is null && conventional.Success)
                {
                    return int.Parse(conventional.Groups[1].Value);
                }
                styleName = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
            }
            return int.TryParse(explicitLevel, out var value) ? value + 1 : 0;
        }
        var headings = xml.Descendants(W + "p").Where(paragraph => !paragraph.Ancestors().Contains(toc) && !paragraph.Ancestors(W + "txbxContent").Any())
            .Select(paragraph => (Paragraph: paragraph, Level: Level(paragraph))).Where(item => item.Level >= firstLevel && item.Level <= lastLevel).ToArray();
        if (headings.Length == 0)
        {
            throw new NotSupportedException("No headings match this TOC's outline levels.");
        }
        var bookmarkId = xml.Descendants(W + "bookmarkStart").Select(item => int.TryParse((string?)item.Attribute(W + "id"), out var value) ? value : 0).DefaultIfEmpty().Max();
        var names = new HashSet<string>(xml.Descendants(W + "bookmarkStart").Select(item => (string?)item.Attribute(W + "name") ?? ""));
        var entries = new List<(string Name, string Text, int Level)>();
        foreach (var heading in headings)
        {
            var name = (string?)heading.Paragraph.Elements(W + "bookmarkStart").FirstOrDefault()?.Attribute(W + "name");
            if (name is null)
            {
                do
                {
                    name = "_DocxKitHeading" + ++bookmarkId;
                }
                while (!names.Add(name));
                var start = new XElement(W + "bookmarkStart", new XAttribute(W + "id", bookmarkId), new XAttribute(W + "name", name));
                if (heading.Paragraph.Element(W + "pPr") is { } properties)
                {
                    properties.AddAfterSelf(start);
                }
                else
                {
                    heading.Paragraph.AddFirst(start);
                }
                heading.Paragraph.Add(new XElement(W + "bookmarkEnd", new XAttribute(W + "id", bookmarkId)));
            }
            entries.Add((name, string.Concat(heading.Paragraph.Descendants(W + "t").Select(item => item.Value)), heading.Level));
        }
        var content = toc.Element(W + "sdtContent") ?? throw new InvalidDataException("TOC content is missing.");
        var titles = content.Elements(W + "p").TakeWhile(paragraph => (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == "TOCHeading").Select(paragraph => new XElement(paragraph)).ToArray();
        var pageNumbers = entries.ToDictionary(entry => entry.Name, _ => "1");
        var entryTitles = entries.ToDictionary(entry => entry.Name, entry => entry.Text);
        for (var pass = 0; pass < 6; pass++)
        {
            content.RemoveNodes();
            content.Add(titles.Select(title => new XElement(title)));
            foreach (var entry in entries)
            {
                var paragraph = new XElement(W + "p",
                    new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", "TOC" + entry.Level)),
                        new XElement(W + "ind", new XAttribute(W + "left", (entry.Level - firstLevel) * 240)),
                        new XElement(W + "tabs", new XElement(W + "tab", new XAttribute(W + "val", "right"), new XAttribute(W + "pos", 8640), new XAttribute(W + "leader", "dot")))),
                    new XElement(W + "hyperlink", new XAttribute(W + "anchor", entry.Name),
                        new XElement(W + "r", new XElement(W + "t", entryTitles[entry.Name])), new XElement(W + "r", new XElement(W + "tab")),
                        new XElement(W + "fldSimple", new XAttribute(W + "instr", " PAGEREF " + entry.Name + " \\h "), new XElement(W + "r", new XElement(W + "t", pageNumbers[entry.Name])))));
                content.Add(paragraph);
            }
            var first = content.Elements(W + "p").Skip(titles.Length).First();
            first.Element(W + "pPr")!.AddAfterSelf(
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
                new XElement(W + "r", new XElement(W + "instrText", " TOC \\o \"" + firstLevel + "-" + lastLevel + "\" \\h \\z \\u ")),
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))));
            content.Elements(W + "p").Last().Add(new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));
            var updated = document.WithXml("word/document.xml", xml);
            using var engine = new DocxLayout(updated);
            var layout = engine.Build();
            var ids = xml.Descendants(W + "p").Select((paragraph, index) => (paragraph, Id: "p" + index)).ToDictionary(item => item.paragraph, item => item.Id);
            var next = new Dictionary<string, string>();
            var nextTitles = new Dictionary<string, string>();
            foreach (var entry in entries)
            {
                var heading = xml.Descendants(W + "bookmarkStart").First(item => (string?)item.Attribute(W + "name") == entry.Name).Ancestors(W + "p").First();
                var target = layout.Pages.FirstOrDefault(page => page.Runs.Any(run => run.Id == ids[heading])) ?? throw new NotSupportedException("A TOC heading could not be laid out.");
                next[entry.Name] = layout.PageLabels?.GetValueOrDefault(target.Number) ?? target.Number.ToString(System.Globalization.CultureInfo.InvariantCulture);
                var marker = string.Concat(layout.Pages.SelectMany(page => page.Runs).Where(run => run.Id == ids[heading] && run.IsListMarker).Select(run => run.Text)).Trim();
                nextTitles[entry.Name] = marker.Length == 0 ? entry.Text : marker + " " + entry.Text;
            }
            if (next.All(item => pageNumbers[item.Key] == item.Value && entryTitles[item.Key] == nextTitles[item.Key]))
            {
                return updated;
            }
            pageNumbers = next;
            entryTitles = nextTitles;
        }
        throw new NotSupportedException("TOC pagination did not converge within six passes.");
    }
}
