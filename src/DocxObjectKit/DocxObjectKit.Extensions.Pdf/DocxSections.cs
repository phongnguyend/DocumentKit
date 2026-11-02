using System.Globalization;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private readonly XElement[] sections;
    private int sectionIndex;
    private readonly Dictionary<DocxPage, int> pageSections = new();
    private string currentPart = "word/document.xml";
    private readonly HashSet<string> absoluteDecorations = new();

    private bool IsAbsoluteDecoration(XElement drawing) => currentPart != "word/document.xml" && ((string?)drawing.Attribute("simplePos") is "1" or "true" || (string?)drawing.Element(WP + "positionV")?.Attribute("relativeFrom") is "page" or "margin" or "topMargin" or "bottomMargin");

    private static void NormalizeAlternates(XElement root)
    {
        XNamespace compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        foreach (var alternate in root.Descendants(compatibility + "AlternateContent").ToArray())
        {
            var branch = alternate.Element(compatibility + "Choice") ?? alternate.Element(compatibility + "Fallback");
            var children = branch?.Elements().ToArray() ?? [];
            foreach (var child in children)
            {
                child.Remove();
            }
            alternate.ReplaceWith(children);
        }
    }

    private void ApplySection(XElement section)
    {
        var size = section.Element(W + "pgSz");
        if (Math.Abs(Number(size, "w", result.Width * 20) / 20 - result.Width) > 1 || Math.Abs(Number(size, "h", result.Height * 20) / 20 - result.Height) > 1)
        {
            warnings.Add("Mixed page sizes use the first section's dimensions.");
        }
        var margins = section.Element(W + "pgMar");
        left = Math.Clamp(Number(margins, "left", 1440) / 20, 0, result.Width / 3);
        right = Math.Clamp(Number(margins, "right", 1440) / 20, 0, result.Width / 3);
        top = Math.Clamp(Number(margins, "top", 1440) / 20, 0, result.Height / 3);
        bottom = Math.Clamp(Number(margins, "bottom", 1440) / 20, 0, result.Height / 3);
        ConfigureColumns(section);
    }

    private string ResolveImage(XElement picture)
    {
        var id = (string?)picture.Descendants(A + "blip").FirstOrDefault()?.Attribute(R + "embed");
        var folder = currentPart[..(currentPart.LastIndexOf('/') + 1)];
        var relPart = folder + "_rels/" + currentPart[(currentPart.LastIndexOf('/') + 1)..] + ".rels";
        if (!document.PartNames.Contains(relPart))
        {
            throw new NotSupportedException("Drawing has no embedded image relationship.");
        }
        var relation = document.ReadXml(relPart).Root!.Elements().FirstOrDefault(item => (string?)item.Attribute("Id") == id);
        if (relation is null || (string?)relation.Attribute("TargetMode") == "External")
        {
            throw new NotSupportedException("External or missing pictures are not loaded.");
        }
        return Resolve(currentPart, (string?)relation.Attribute("Target") ?? "");
    }

    private void TextBoxes(XElement drawing, double originX, double originY, string anchorId)
    {
        var savedY = y;
        var savedStyle = previousParagraphStyle;
        var savedContext = previousContextual;
        previousParagraphStyle = null;
        var savedSpacing = pendingAfter;
        pendingAfter = 0;
        var flowEnd = y;
        var savedMeasuring = measuringTable;
        try
        {
            foreach (var content in drawing.Descendants(W + "txbxContent"))
            {
                var shape = content.Ancestors().FirstOrDefault(item => item.Attribute("style") is not null);
                var style = ((string?)shape?.Attribute("style") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries).Select(value => value.Split(':', 2)).Where(pair => pair.Length == 2).GroupBy(pair => pair[0]).ToDictionary(group => group.Key, group => group.Last()[1]);
                double Css(string key, double fallback)
                {
                    if (!style.TryGetValue(key, out var value))
                    {
                        return fallback;
                    }
                    var numeric = Regex.Match(value, @"^-?[0-9.]+").Value;
                    return double.TryParse(numeric, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number * (value.EndsWith("in", StringComparison.Ordinal) ? 72 : 1) : fallback;
                }
                var extent = drawing.Element(WP + "extent");
                var width = Math.Max(12, extent is null ? Css("width", result.Width - left - right) : Number(extent, "cx") / 12700);
                var boxX = originX + Css("margin-left", 0);
                var boxY = originY + Css("margin-top", 0);
                boxX = DrawingPosition(drawing, "positionH", boxX, width);
                boxY = DrawingPosition(drawing, "positionV", boxY, extent is null ? Css("height", 0) : Number(extent, "cy") / 12700);
                var wraps = drawing.Element(WP + "wrapNone") is null && currentPart == "word/document.xml";
                var height = extent is null ? Css("height", 0) : Number(extent, "cy") / 12700;
                if (wraps && !savedMeasuring && boxY + height > result.Height - bottom && y > top + 1)
                {
                    var previousLeft = FlowLeft;
                    var previousPage = page;
                    var anchorRuns = previousPage.Runs.Where(run => run.Id == anchorId).ToArray();
                    AdvanceFlow();
                    var deltaX = FlowLeft - previousLeft;
                    var deltaY = y - originY;
                    // A floating textbox and its anchor paragraph travel together.
                    previousPage.Runs.RemoveAll(run => run.Id == anchorId);
                    page.Runs.AddRange(anchorRuns.Select(run => run with { X = run.X + deltaX, Y = run.Y + deltaY, Baseline = run.Baseline + deltaY }));
                    boxX += deltaX;
                    boxY = y + Math.Max(0, boxY - originY);
                    flowEnd = y;
                }
                var bodyProperties = drawing.Descendants().FirstOrDefault(item => item.Name.LocalName == "bodyPr");
                var insetLeft = Number(bodyProperties, "lIns", 50800) / 12700;
                var insetRight = Number(bodyProperties, "rIns", 50800) / 12700;
                var insetTop = Number(bodyProperties, "tIns", 38100) / 12700;
                var insetBottom = Number(bodyProperties, "bIns", 38100) / 12700;
                var boxIds = content.Descendants(W + "p").Select(item => paragraphIds[item]).ToHashSet();
                y = boxY + insetTop;
                measuringTable = true;
                foreach (var paragraph in Blocks(content).Where(item => item.Name == W + "p"))
                {
                    if (IsAbsoluteDecoration(drawing))
                    {
                        absoluteDecorations.Add(paragraphIds[paragraph]);
                    }
                    Paragraph(paragraph, boxX + insetLeft, Math.Max(12, width - insetLeft - insetRight), true);
                }
                var spare = Math.Max(0, height - insetTop - insetBottom - (y - boxY - insetTop));
                var verticalShift = (string?)bodyProperties?.Attribute("anchor") switch
                {
                    "ctr" => spare / 2,
                    "b" => spare,
                    _ => 0
                };
                if (verticalShift > 0)
                {
                    for (var runIndex = 0; runIndex < page.Runs.Count; runIndex++)
                    {
                        var run = page.Runs[runIndex];
                        if (boxIds.Contains(run.Id))
                        {
                            page.Runs[runIndex] = run with { Y = run.Y + verticalShift, Baseline = run.Baseline + verticalShift };
                        }
                    }
                    for (var objectIndex = 0; objectIndex < page.TextObjects.Count; objectIndex++)
                    {
                        var item = page.TextObjects[objectIndex];
                        if (boxIds.Contains(item.Id))
                        {
                            page.TextObjects[objectIndex] = item with { Bounds = item.Bounds with { Y = item.Bounds.Y + verticalShift } };
                        }
                    }
                    y += verticalShift;
                }
                measuringTable = savedMeasuring;
                if (wraps)
                {
                    flowEnd = Math.Max(flowEnd, Math.Max(y + 3, boxY + height));
                }
                warnings.Add("Floating text boxes are rendered with approximate wrapping and placement.");
            }
        }
        finally
        {
            y = currentPart == "word/document.xml" ? flowEnd : savedY;
            measuringTable = savedMeasuring;
            pendingAfter = savedSpacing;
            previousParagraphStyle = savedStyle;
            previousContextual = savedContext;
        }
    }

    private void RenderPageDecorations()
    {
        if (sections.Length == 0 || !document.PartNames.Contains("word/_rels/document.xml.rels"))
        {
            return;
        }
        var relationships = document.ReadXml("word/_rels/document.xml.rels").Root!.Elements().ToArray();
        var references = new Dictionary<string, string>();
        var sectionReferences = new List<Dictionary<string, string>>();
        foreach (var section in sections)
        {
            foreach (var reference in section.Elements().Where(item => item.Name == W + "headerReference" || item.Name == W + "footerReference"))
            {
                var relation = relationships.FirstOrDefault(item => (string?)item.Attribute("Id") == (string?)reference.Attribute(R + "id"));
                if (relation is not null && (string?)relation.Attribute("TargetMode") != "External")
                {
                    references[reference.Name.LocalName + ":" + (Value(reference, "type") ?? "default")] = Resolve("word/document.xml", (string?)relation.Attribute("Target") ?? "");
                }
            }
            sectionReferences.Add(new(references));
        }
        var settings = document.PartNames.Contains("word/settings.xml") ? document.ReadXml("word/settings.xml").Root : null;
        var useEven = settings is not null && Enabled(settings, "evenAndOddHeaders");
        var pages = result.Pages.ToArray();
        long displayedPage = 0;
        var previousSection = -1;
        foreach (var target in pages)
        {
            var index = pageSections.GetValueOrDefault(target);
            var section = sections[Math.Min(index, sections.Length - 1)];
            displayedPage++;
            if (index != previousSection && int.TryParse(Value(section.Element(W + "pgNumType"), "start"), NumberStyles.None, CultureInfo.InvariantCulture, out var restart) && restart >= 0)
            {
                displayedPage = restart;
            }
            previousSection = index;
            ApplySection(section);
            var first = pages.First(item => pageSections.GetValueOrDefault(item) == index) == target;
            foreach (var kind in new[] { "header", "footer" })
            {
                var type = first && Enabled(section, "titlePg") ? "first" : useEven && displayedPage % 2 == 0 ? "even" : "default";
                if (!sectionReferences[index].TryGetValue(kind + "Reference:" + type, out var part) || !document.PartNames.Contains(part))
                {
                    continue;
                }
                var decoration = document.ReadXml(part);
                var ordinal = 0;
                foreach (var paragraph in decoration.Descendants(W + "p"))
                {
                    paragraphIds[paragraph] = "decoration-" + target.Number + "-" + kind + "-p" + ordinal++;
                }
                ordinal = 0;
                foreach (var picture in decoration.Descendants().Where(item => item.Name == WP + "inline" || item.Name == WP + "anchor"))
                {
                    imageIds[picture] = "decoration-" + target.Number + "-" + kind + "-i" + ordinal++;
                }
                NormalizeAlternates(decoration.Root!);
                foreach (var field in decoration.Descendants(W + "fldSimple"))
                {
                    var instruction = Value(field, "instr") ?? "";
                    if (!Regex.IsMatch(instruction, @"^\s*(PAGE|NUMPAGES)(?:\s|$)"))
                    {
                        continue;
                    }
                    var value = Regex.IsMatch(instruction, @"^\s*NUMPAGES(?:\s|$)") ? pages.Length : displayedPage;
                    var texts = field.Descendants(W + "t").ToArray();
                    if (texts.Length == 0)
                    {
                        field.Add(new XElement(W + "r", new XElement(W + "t", value.ToString(CultureInfo.InvariantCulture))));
                    }
                    else
                    {
                        texts[0].Value = value.ToString(CultureInfo.InvariantCulture);
                        foreach (var extra in texts.Skip(1))
                        {
                            extra.Value = "";
                        }
                    }
                }
                // Other cached field values are preserved except the current page number.
                var pageField = false;
                foreach (var run in decoration.Descendants(W + "r"))
                {
                    if (run.Element(W + "instrText") is { } instruction)
                    {
                        pageField = Regex.IsMatch(instruction.Value, @"^\s*PAGE(?:\s|$)");
                    }
                    if (pageField && run.Element(W + "t") is { } value)
                    {
                        value.Value = displayedPage.ToString(CultureInfo.InvariantCulture);
                        pageField = false;
                    }
                }
                var scratch = new DocxPage(0, "decoration", "", [], [], [], []);
                result.Pages.Add(scratch);
                var previousPage = page;
                var previousY = y;
                currentPart = part;
                measuringTable = true;
                page = scratch;
                y = 0;
                pendingAfter = 0;
                previousParagraphStyle = null;
                try
                {
                    foreach (var block in Blocks(decoration.Root!))
                    {
                        if (block.Name == W + "p")
                        {
                            Paragraph(block, left, result.Width - left - right);
                        }
                    }
                    var distance = Number(section.Element(W + "pgMar"), kind, 720) / 20;
                    var shift = kind == "header" ? distance : Math.Max(0, result.Height - distance - y);
                    target.Runs.AddRange(scratch.Runs.Select(run => run with { Y = run.Y + (absoluteDecorations.Contains(run.Id) ? 0 : shift), Baseline = run.Baseline + (absoluteDecorations.Contains(run.Id) ? 0 : shift) }));
                    target.Images.AddRange(scratch.Images.Select(image => image with { Y = image.Y + (absoluteDecorations.Contains(image.Id) ? 0 : shift), Grouped = true, ReadOnly = true }));
                }
                finally
                {
                    currentPart = "word/document.xml";
                    measuringTable = false;
                    result.Pages.Remove(scratch);
                    page = previousPage;
                    y = previousY;
                }
            }
        }
    }
}
