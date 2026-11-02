using System.Xml.Linq;
using System.Text.RegularExpressions;
using DocxObjectKit.Models;
using DocxObjectKit.Editors;
using SkiaSharp;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed record DocxBounds(double X, double Y, double Width, double Height, bool Grouped = false, double Rotation = 0);
public sealed record DocxTextObject(string Id, string Name, string Text, DocxBounds Bounds);
public sealed record DocxImageObject(string Id, string Name, string PartName, double X, double Y, double Width, double Height, bool Grouped = false, double Rotation = 0, bool ReadOnly = false, double CropLeft = 0, double CropTop = 0, double CropRight = 0, double CropBottom = 0, bool FlipH = false, bool FlipV = false);
public sealed record DocxGlyphRun(string Id, string Text, double X, double Y, double Baseline, double Width, double Height, double FontSize, string FontFamily, string Color, bool Bold, bool Italic, bool Underline, string? Background = null, string? NoteKey = null, bool IsListMarker = false);
public sealed record DocxBorder(double Width, string Color);

public sealed record DocxRule(double X, double Y, double Width, double Height, string? Fill = null, double StrokeWidth = .5, string Color = "888888", DocxBorder?[]? Edges = null);
public sealed record DocxPage(int Number, string Id, string Title, List<DocxTextObject> TextObjects, List<DocxImageObject> Images, List<DocxGlyphRun> Runs, List<DocxRule> Rules);
public sealed record DocxLayoutResult(double Width, double Height, List<DocxPage> Pages, string[] Warnings, Dictionary<string, string> Media, Dictionary<int, int>? PageNumbers = null, Dictionary<int, string>? PageLabels = null);

public sealed partial class DocxLayout : IDisposable
{
    private readonly DocxObjectDocument document;
    private readonly XDocument xml;
    private readonly Dictionary<string, XElement> styles;
    private readonly XElement? defaults;
    private readonly XElement? compatibility;
    private readonly Dictionary<string, SKTypeface> fonts = new();
    private readonly SKTypeface fallback;
    private readonly HashSet<string> warnings = new(StringComparer.Ordinal);
    private readonly Dictionary<XElement, string> paragraphIds;
    private readonly Dictionary<XElement, string> imageIds;
    private readonly DocxLayoutResult result;
    private double left;
    private double right;
    private double top;
    private double bottom;
    private DocxPage page;
    private double y;
    private double pendingAfter;
    private string? previousParagraphStyle;
    private bool previousContextual;
    private bool measuringTable;
    private readonly Dictionary<string, (byte[] Bytes, double Width, double Height)> decodedEmf = new();

    public DocxLayout(DocxObjectDocument document)
    {
        this.document = document;
        compatibility = document.PartNames.Contains("word/settings.xml")
            ? document.ReadXml("word/settings.xml").Root?.Element(W + "compat") : null;
        xml = document.ReadXml("word/document.xml");
        var styleXml = document.PartNames.Contains("word/styles.xml") ? document.ReadXml("word/styles.xml") : null;
        styles = styleXml?.Descendants(W + "style").Where(item => Value(item, "styleId") is not null).GroupBy(item => Value(item, "styleId")!).ToDictionary(group => group.Key, group => group.First()) ?? new();
        defaults = styleXml?.Root?.Element(W + "docDefaults");
        using var stream = typeof(DocxLayout).Assembly.GetManifestResourceStream("DocxObjectKit.Extensions.Pdf.Fonts.NotoSans-Regular.ttf")!;
        using var data = SKData.Create(stream);
        fallback = SKTypeface.FromData(data);
        paragraphIds = xml.Descendants(W + "p").Select((item, index) => (item, id: "p" + index)).ToDictionary(item => item.item, item => item.id);
        imageIds = DocxEditor.Pictures(xml).Select((item, index) => (item, id: "i" + index)).ToDictionary(item => item.item, item => item.id);
        NormalizeAlternates(xml.Root!);
        sections = xml.Descendants(W + "sectPr").ToArray();
        var section = sections.FirstOrDefault();
        var size = section?.Element(W + "pgSz");
        var width = Number(size, "w", 12240) / 20;
        var height = Number(size, "h", 15840) / 20;
        if (width < 72 || height < 72 || width > 3000 || height > 3000)
        {
            throw new NotSupportedException("Page dimensions must be between 72 and 3000 points.");
        }
        var margins = section?.Element(W + "pgMar");
        left = Math.Clamp(Number(margins, "left", 1440) / 20, 0, width / 3);
        right = Math.Clamp(Number(margins, "right", 1440) / 20, 0, width / 3);
        top = Math.Clamp(Number(margins, "top", 1440) / 20, 0, height / 3);
        bottom = Math.Clamp(Number(margins, "bottom", 1440) / 20, 0, height / 3);
        result = new(width, height, [], [], new());
        ConfigureColumns(section);
        page = NewPage();
        LoadNotes();
    }

    private void BuildBody(CancellationToken cancellationToken)
    {
        var blocks = Blocks(xml.Root!.Element(W + "body")!).ToArray();
        for (var blockIndex = 0; blockIndex < blocks.Length; blockIndex++)
        {
            var block = blocks[blockIndex];
            cancellationToken.ThrowIfCancellationRequested();
            if (block.Name == W + "p")
            {
                KeepParagraphsTogether(blocks, blockIndex);
                Paragraph(block, FlowLeft, FlowWidth);
            }
            else if (block.Name == W + "tbl")
            {
                Table(block);
            }
            else if (block.Name != W + "sectPr" && block.Name != W + "bookmarkStart" && block.Name != W + "bookmarkEnd")
            {
                warnings.Add("Unsupported body block: " + block.Name.LocalName);
            }
            if (block.Element(W + "pPr")?.Element(W + "sectPr") is not null && sectionIndex + 1 < sections.Length)
            {
                var nextSectionTop = Math.Max(y, flowBottom);
                sectionIndex++;
                ApplySection(sections[sectionIndex]);
                if (Value(sections[sectionIndex].Element(W + "type")) == "continuous")
                {
                    y = nextSectionTop;
                    columnTop = y;
                }
                else
                {
                    if (page.Runs.Count == 0 && page.Images.Count == 0 && page.Rules.Count == 0)
                    {
                        pageSections[page] = sectionIndex;
                        y = top;
                    }
                    else
                    {
                        page = NewPage();
                    }
                }
            }
        }
        foreach (var feature in new[] { "oMath", "ins", "del" })
        {
            if (xml.Descendants().Any(item => item.Name.LocalName == feature && (feature != "cols" || Number(item, "num", 1) > 1)))
            {
                warnings.Add($"Word feature '{feature}' is not fully rendered; its package data is preserved.");
            }
        }
    }

    private DocxPage NewPage()
    {
        if (result.Pages.Count >= 500)
        {
            throw new NotSupportedException("Layout exceeds 500 pages.");
        }
        var created = new DocxPage(result.Pages.Count + 1, "page-" + (result.Pages.Count + 1), "", [], [], [], []);
        result.Pages.Add(created);
        pageSections[created] = sectionIndex;
        columnIndex = 0;
        columnTop = top;
        flowBottom = top;
        y = top;
        pendingAfter = 0;
        return created;
    }

    private XElement Properties(XElement paragraph, string kind, XElement? run = null)
    {
        var merged = new XElement(W + kind);
        void Overlay(XElement? properties)
        {
            if (properties is null)
            {
                return;
            }
            foreach (var property in properties.Elements())
            {
                MergeProperty(merged, property);
            }
        }
        Overlay(defaults?.Element(W + (kind + "Default"))?.Element(W + kind));
        var chain = new List<XElement>();
        var name = Value(paragraph.Element(W + "pPr")?.Element(W + "pStyle")) ?? styles.Values.FirstOrDefault(item => Value(item, "type") == "paragraph" && Value(item, "default") == "1")?.Attribute(W + "styleId")?.Value;
        var visited = new HashSet<string>();
        while (name is not null && visited.Add(name) && styles.TryGetValue(name, out var style))
        {
            chain.Add(style);
            name = Value(style.Element(W + "basedOn"));
        }
        var explicitStyle = paragraph.Element(W + "pPr")?.Element(W + "pStyle") is not null;
        if (explicitStyle)
        {
            Overlay(tableFormatting?.Element(W + kind));
        }
        foreach (var style in chain.AsEnumerable().Reverse())
        {
            Overlay(style.Element(W + kind));
        }
        if (!explicitStyle)
        {
            Overlay(tableFormatting?.Element(W + kind));
        }
        // Direct pPr/rPr formats the paragraph mark, not the text runs.
        if (run is null)
        {
            Overlay(paragraph.Element(W + "pPr")?.Element(W + kind));
        }
        if (kind == "pPr")
        {
            Overlay(paragraph.Element(W + kind));
        }
        var runStyle = Value(run?.Element(W + "rPr")?.Element(W + "rStyle"));
        if (runStyle is not null && styles.TryGetValue(runStyle, out var characterStyle))
        {
            Overlay(characterStyle.Element(W + kind));
        }
        Overlay(run?.Element(W + kind));
        return merged;
    }

    private static bool Enabled(XElement properties, string name) => properties.Element(W + name) is { } item && Value(item) is not ("0" or "false" or "off" or "none");

    internal SKTypeface Face(string family, bool bold, bool italic)
    {
        var key = family + ":" + bold + ":" + italic;
        if (fonts.TryGetValue(key, out var cached))
        {
            return cached;
        }
        var face = SKFontManager.Default.MatchFamily(family, new SKFontStyle(bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal, SKFontStyleWidth.Normal, italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright));
        if (face is null || !string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase))
        {
            face?.Dispose();
            var substitute = family.ToLowerInvariant() switch
            {
                "arial" => "Liberation Sans",
                "times new roman" => "Liberation Serif",
                "courier new" => "Liberation Mono",
                "calibri" => "Carlito",
                "cambria" => "Caladea",
                _ => null
            };
            if (substitute is not null)
            {
                var replacement = SKFontManager.Default.MatchFamily(substitute, new SKFontStyle(bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal, SKFontStyleWidth.Normal, italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright));
                if (replacement is not null && string.Equals(replacement.FamilyName, substitute, StringComparison.OrdinalIgnoreCase))
                {
                    warnings.Add($"Font '{family}' substitutes metric-compatible '{substitute}'.");
                    fonts[key] = replacement;
                    return replacement;
                }
                replacement?.Dispose();
            }
            if (family != fallback.FamilyName)
            {
                warnings.Add($"Font '{family}' substitutes Noto Sans; wrapping may differ from Word.");
            }
            return fallback;
        }
        fonts[key] = face;
        return face;
    }

    internal SKFont Font(string family, bool bold, bool italic, double size) => new(Face(family, bold, italic), (float)size) { Embolden = bold && !Face(family, bold, italic).IsBold, SkewX = italic && !Face(family, bold, italic).IsItalic ? -.2f : 0 };

    private void Paragraph(XElement paragraph, double x, double width, bool inTable = false)
    {
        var properties = Properties(paragraph, "pPr");
        var numberingLevel = NumberingLevel(properties);
        foreach (var property in numberingLevel?.Element(W + "pPr")?.Elements() ?? [])
        {
            var inherited = new XElement(property);
            if (properties.Element(property.Name) is { } existing)
            {
                var container = new XElement("format", inherited);
                MergeProperty(container, existing);
                existing.ReplaceWith(inherited);
            }
            else
            {
                properties.Add(inherited);
            }
        }
        var spacing = properties.Element(W + "spacing");
        if (!measuringTable && Enabled(properties, "pageBreakBefore") && (page.Runs.Count > 0 || page.Images.Count > 0))
        {
            page = NewPage();
        }
        var paragraphStyleId = Value(paragraph.Element(W + "pPr")?.Element(W + "pStyle")) ?? "default";
        var sameStyle = previousParagraphStyle == paragraphStyleId;
        if (sameStyle && previousContextual)
        {
            y -= pendingAfter;
            pendingAfter = 0;
        }
        var before = sameStyle && Enabled(properties, "contextualSpacing") ? 0 : Number(spacing, "before") / 20;
        if (!measuringTable && currentPart == "word/document.xml" && properties.Element(W + "framePr") is null
            && Math.Abs(y - columnTop) < .01
            && Number(compatibility?.Elements(W + "compatSetting").FirstOrDefault(item => Value(item, "name") == "compatibilityMode"), "val") >= 15
            && result.Pages.TakeWhile(item => item != page).Any(item => pageSections.GetValueOrDefault(item) == sectionIndex))
        {
            before = 0;
        }
        y += Math.Max(0, before - pendingAfter);
        pendingAfter = 0;
        previousParagraphStyle = paragraphStyleId;
        previousContextual = Enabled(properties, "contextualSpacing");
        var frame = properties.Element(W + "framePr");
        var flowingY = y;
        if (frame is not null)
        {
            var horizontal = Value(frame, "hAnchor") ?? "margin";
            var vertical = Value(frame, "vAnchor") ?? "text";
            var originX = horizontal == "page" ? 0 : horizontal == "margin" ? left : x;
            var referenceWidth = horizontal == "page" ? result.Width : horizontal == "margin" ? result.Width - left - right : width;
            var explicitWidth = Number(frame, "w");
            width = Math.Max(12, explicitWidth > 0 ? explicitWidth / 20 : referenceWidth);
            x = originX + Number(frame, "x") / 20;
            var alignment = Value(frame, "xAlign");
            if (explicitWidth > 0)
            {
                x = alignment switch
                {
                    "right" => originX + referenceWidth - width,
                    "center" => originX + (referenceWidth - width) / 2,
                    "left" => originX,
                    _ => x
                };
            }
            else if (alignment is "right" or "center" or "left")
            {
                properties.Elements(W + "jc").Remove();
                properties.Add(new XElement(W + "jc", new XAttribute(W + "val", alignment)));
            }
            y = (vertical == "page" ? 0 : vertical == "margin" ? top : flowingY) + Number(frame, "y") / 20;
            if (currentPart != "word/document.xml" && vertical is "page" or "margin")
            {
                absoluteDecorations.Add(paragraphIds[paragraph]);
            }
        }
        var indent = properties.Element(W + "ind");
        x += Number(indent, "left") / 20;
        width = Math.Max(12, width - (Number(indent, "left") + Number(indent, "right")) / 20);
        var flowOrigin = FlowLeft;
        var flowWidth = FlowWidth;
        var id = paragraphIds[paragraph];
        var line = new List<DocxGlyphRun>();
        var floatingText = new List<XElement>();
        var offset = (Number(indent, "firstLine") - Number(indent, "hanging")) / 20;
        var paragraphStyle = Properties(paragraph, "rPr");
        var paragraphSize = Math.Clamp(Number(paragraphStyle.Element(W + "sz"), "val", 22) / 2, 1, 400);
        using var paragraphFont = Font(FontFamily(paragraphStyle), Enabled(paragraphStyle, "b"), Enabled(paragraphStyle, "i"), paragraphSize);
        var defaultAdvance = LineAdvance(paragraphFont, spacing, paragraphSize);
        var lineHeight = 0d;
        var completedLines = 0;
        var endsWithPageBreak = false;
        var endsWithInlineImage = false;
        var plannedBreak = measuringTable ? null : widowBreakAfter;
        if (!measuringTable)
        {
            widowBreakAfter = null;
        }
        void ContinueFlow(double requiredHeight)
        {
            var breakForWidow = plannedBreak == completedLines && line.Count == 0;
            while (!measuringTable && frame is null && (breakForWidow || y + requiredHeight > FlowLimit) && (y > columnTop || columnTop > top))
            {
                AdvanceFlow();
                breakForWidow = false;
                plannedBreak = null;
                var shift = FlowLeft - flowOrigin;
                x += shift;
                width = Math.Max(12, width + FlowWidth - flowWidth);
                for (var index = 0; index < line.Count; index++)
                {
                    line[index] = line[index] with { X = line[index].X + shift };
                }
                flowOrigin = FlowLeft;
                flowWidth = FlowWidth;
            }
        }
        void Flush(bool force = false)
        {
            if (line.Count == 0 && !force)
            {
                return;
            }
            if (lineHeight == 0)
            {
                lineHeight = defaultAdvance;
            }
            ContinueFlow(lineHeight);
            measuredLines?.Add((y, y + lineHeight));
            var alignment = Value(properties.Element(W + "jc"));
            var visibleWidth = line.LastOrDefault(piece => piece.Text.Any(character => character != ' ')) is { } lastVisible
                ? lastVisible.X + lastVisible.Width - x : 0;
            var shift = alignment is "center" ? Math.Max(0, (width - visibleWidth) / 2) : alignment is "right" or "end" ? Math.Max(0, width - visibleWidth) : 0;
            var baseline = line.Select(piece => piece.Baseline).DefaultIfEmpty(0).Max();
            foreach (var piece in line)
            {
                page.Runs.Add(piece with { X = piece.X + shift, Y = y + baseline - piece.Baseline, Baseline = y + baseline });
            }
            y += lineHeight;
            completedLines++;
            if (!measuringTable && frame is null)
            {
                flowBottom = Math.Max(flowBottom, y);
            }
            line.Clear();
            offset = 0;
            lineHeight = 0;
        }
        void Add(string text, XElement run, bool isMarker = false, string? noteKey = null)
        {
            var style = Properties(paragraph, "rPr", run);
            if (Enabled(style, "vanish"))
            {
                return;
            }
            var size = Math.Clamp(Number(style.Element(W + "sz"), "val", 22) / 2, 1, 400);
            var bold = Enabled(style, "b");
            var italic = Enabled(style, "i");
            var family = FontFamily(style);
            var face = Face(family, bold, italic);
            var color = Value(style.Element(W + "color")) ?? "000000";
            if (!Regex.IsMatch(color, "^[0-9a-fA-F]{6}$"))
            {
                color = "000000";
            }
            using var font = Font(family, bold, italic, size);
            if (!font.ContainsGlyphs(text))
            {
                warnings.Add("Some characters are unavailable in the selected font and may not render correctly.");
            }
            if (Regex.IsMatch(text, @"[\u0590-\u0dff]"))
            {
                warnings.Add("Complex-script shaping and bidirectional text are not yet supported.");
            }
            var advance = isMarker ? defaultAdvance : LineAdvance(font, spacing, size);
            foreach (var word in Regex.Split(text.Replace("\t", "    "), @"( +)").Where(item => item.Length > 0))
            {
                var measured = font.MeasureText(word);
                if (word.All(character => character == ' ') && offset + measured > width)
                {
                    // Trailing spaces may extend past the line edge without creating a
                    // new line. Keep their advance so following text still wraps.
                    if (line.Count > 0)
                    {
                        offset += measured;
                    }
                    continue;
                }
                // Font advances accumulate fractional-point differences across runs.
                // Bound the fit tolerance to two twips instead of wrapping a whole
                // word for a sub-pixel overrun.
                const double fitTolerance = .1;
                if (offset + measured > width + fitTolerance && line.Count > 0)
                {
                    Flush();
                }
                if (line.Count == 0)
                {
                    ContinueFlow(Math.Max(lineHeight, advance));
                }
                if (offset == 0 && string.IsNullOrWhiteSpace(word))
                {
                    continue;
                }
                var starts = measured > width + fitTolerance ? System.Globalization.StringInfo.ParseCombiningCharacters(word) : [0];
                var chunks = starts.Select((start, i) => word.Substring(start, (i + 1 < starts.Length ? starts[i + 1] : word.Length) - start));
                foreach (var chunk in chunks)
                {
                    if (line.Count == 0)
                    {
                        ContinueFlow(Math.Max(lineHeight, advance));
                    }
                    measured = font.MeasureText(chunk);
                    if (offset + measured > width + fitTolerance && line.Count > 0)
                    {
                        Flush();
                        ContinueFlow(Math.Max(lineHeight, advance));
                    }
                    var height = Math.Max(size * 1.15, font.Metrics.Descent - font.Metrics.Ascent);
                    line.Add(new(id, chunk, x + offset, 0, -font.Metrics.Ascent, measured, height, size, face.FamilyName, color, bold, italic, Enabled(style, "u"), Value(style.Element(W + "shd"), "fill"), noteKey, isMarker && noteKey is null));
                    offset += measured;
                    if (!isMarker)
                    {
                        lineHeight = Math.Max(lineHeight, advance);
                    }
                }
            }
        }
        if (numberingLevel is not null)
        {
            var markerRun = new XElement(W + "r", numberingLevel.Element(W + "rPr") is { } markerProperties ? new XElement(markerProperties) : null);
            var marker = ListMarker(properties, numberingLevel);
            var markerStyle = Properties(paragraph, "rPr", markerRun);
            using var markerFont = Font(FontFamily(markerStyle), Enabled(markerStyle, "b"), Enabled(markerStyle, "i"), paragraphSize);
            if (Value(numberingLevel.Element(W + "numFmt")) == "bullet" && !markerFont.ContainsGlyphs(marker))
            {
                var fallbackProperties = markerRun.Element(W + "rPr");
                if (fallbackProperties is null)
                {
                    fallbackProperties = new XElement(W + "rPr");
                    markerRun.Add(fallbackProperties);
                }
                fallbackProperties.Elements(W + "rFonts").Remove();
                fallbackProperties.Add(new XElement(W + "rFonts", new XAttribute(W + "ascii", fallback.FamilyName), new XAttribute(W + "hAnsi", fallback.FamilyName)));
            }
            Add(marker, markerRun, true);
            var suffix = Value(numberingLevel.Element(W + "suff")) ?? "tab";
            if (suffix == "space")
            {
                Add(" ", markerRun, true);
            }
            else if (suffix == "tab")
            {
                var stop = properties.Element(W + "tabs")?.Elements(W + "tab")
                    .Where(item => Value(item) == "num" && Number(item, "pos") / 20 - Number(indent, "left") / 20 >= offset)
                    .OrderBy(item => Number(item, "pos")).FirstOrDefault();
                offset = stop is null ? Math.Max(0, offset + paragraphFont.MeasureText(" ")) : Number(stop, "pos") / 20 - Number(indent, "left") / 20;
            }
        }
        foreach (var run in paragraph.Descendants(W + "r").Where(item => item.Ancestors(W + "p").FirstOrDefault() == paragraph && !item.Ancestors(W + "del").Any()))
        {
            foreach (var node in run.Elements())
            {
                if (node.Name != W + "rPr" && node.Name != W + "lastRenderedPageBreak" && node.Name != W + "instrText"
                    && (node.Name != W + "t" || node.Value.Length > 0))
                {
                    endsWithPageBreak = false;
                    if (node.Name != W + "t" || node.Value.Any(character => character != ' '))
                    {
                        endsWithInlineImage = false;
                    }
                }
                if (node.Name == W + "t")
                {
                    Add(node.Value, run);
                }
                else if (node.Name == W + "footnoteReference" || node.Name == W + "endnoteReference")
                {
                    var key = node.Name.LocalName.Replace("Reference", "") + ":" + Value(node, "id");
                    if (notes.TryGetValue(key, out var note))
                    {
                        Add("[" + note.Number + "]", run, true, key);
                    }
                    else
                    {
                        warnings.Add("A note reference has no matching note definition.");
                    }
                }
                else if (node.Name == W + "tab")
                {
                    var tab = properties.Element(W + "tabs")?.Elements(W + "tab").Where(item => Value(item) is not ("clear" or "bar") && Number(item, "pos") / 20 - Number(indent, "left") / 20 > offset + .1).OrderBy(item => Number(item, "pos")).FirstOrDefault();
                    var stop = tab is null ? (Math.Floor(offset / 36) + 1) * 36 : Number(tab, "pos") / 20 - Number(indent, "left") / 20;
                    var following = string.Concat(paragraph.Descendants().SkipWhile(item => item != node).Skip(1).TakeWhile(item => item.Name != W + "tab" && item.Name != W + "br").Where(item => item.Name == W + "t").Select(item => item.Value));
                    var tabStyle = Properties(paragraph, "rPr", run);
                    var tabSize = Math.Clamp(Number(tabStyle.Element(W + "sz"), "val", 22) / 2, 1, 400);
                    using var tabFont = Font(FontFamily(tabStyle), Enabled(tabStyle, "b"), Enabled(tabStyle, "i"), tabSize);
                    var tailWidth = tabFont.MeasureText(following);
                    var adjustment = Value(tab) == "right" ? tailWidth : Value(tab) == "center" ? tailWidth / 2 : 0;
                    var targetOffset = Math.Max(offset, Math.Min(width - tailWidth, stop - adjustment));
                    var leader = Value(tab, "leader") switch
                    {
                        "dot" => ".",
                        "hyphen" => "-",
                        "underscore" => "_",
                        _ => ""
                    };
                    if (leader.Length > 0)
                    {
                        var count = Math.Clamp((int)((targetOffset - offset) / Math.Max(.1f, tabFont.MeasureText(leader))), 0, 2000);
                        Add(string.Concat(Enumerable.Repeat(leader, count)), run);
                    }
                    offset = targetOffset;
                }
                else if (node.Name == W + "br")
                {
                    Flush(true);
                    if (!measuringTable && Value(node, "type") is "page" or "column")
                    {
                        if (Value(node, "type") == "page")
                        {
                            page = NewPage();
                            endsWithPageBreak = true;
                        }
                        else
                        {
                            AdvanceFlow();
                        }
                        x += FlowLeft - flowOrigin;
                        width = Math.Max(12, width + FlowWidth - flowWidth);
                        flowOrigin = FlowLeft;
                        flowWidth = FlowWidth;
                    }
                }
                else if (node.Name == W + "object" || node.Name == W + "pict")
                {
                    if (node.Descendants().Any(item => item.Name.LocalName == "imagedata"))
                    {
                        Flush();
                        LegacyPictures(node, x, width);
                    }
                    else
                    {
                        floatingText.Add(node);
                    }
                }
                else if (node.Name == W + "drawing")
                {
                    foreach (var picture in node.Elements())
                    {
                        if (picture.Descendants(W + "txbxContent").Any() || IsLineShape(picture))
                        {
                            floatingText.Add(picture);
                        }
                        else
                        {
                            Flush();
                            var imageCount = page.Images.Count;
                            var imagePage = page;
                            Image(picture, x, width, Value(properties.Element(W + "jc")));
                            endsWithInlineImage = picture.Name == WP + "inline"
                                && (page != imagePage || page.Images.Count > imageCount);
                            if (!measuringTable && frame is null)
                            {
                                x += FlowLeft - flowOrigin;
                                width = Math.Max(12, width + FlowWidth - flowWidth);
                                flowOrigin = FlowLeft;
                                flowWidth = FlowWidth;
                            }
                        }
                    }
                }
            }
        }
        var paragraphEndedBeforeBreak = endsWithPageBreak && (compatibility is null || !Enabled(compatibility, "splitPgBreakAndParaMark"));
        if (endsWithInlineImage && line.Count == 0 && Value(spacing, "lineRule") is not ("exact" or "atLeast"))
        {
            // The inline picture already occupies the line above its baseline.
            // Reserve only the paragraph mark's descent, not another text line.
            lineHeight = Math.Max(.1, paragraphFont.Metrics.Descent * Math.Max(.01, Number(spacing, "line", 240) / 240));
        }
        Flush(!paragraphEndedBeforeBreak);
        pendingAfter = paragraphEndedBeforeBreak ? 0 : Number(spacing, "after", 0) / 20;
        y += pendingAfter;
        var anchorY = page.Runs.Where(run => run.Id == id).Select(run => run.Y).DefaultIfEmpty(flowingY).Min();
        foreach (var drawing in floatingText)
        {
            if (IsLineShape(drawing))
            {
                LineShape(drawing, x, page.Runs.Any(run => run.Id == id) ? anchorY : y - pendingAfter - defaultAdvance);
            }
            else
            {
                TextBoxes(drawing, x, anchorY, id);
            }
        }
        foreach (var target in result.Pages)
        {
            var runs = target.Runs.Where(item => item.Id == id).ToArray();
            if (runs.Length > 0)
            {
                var minX = runs.Min(item => item.X);
                var minY = runs.Min(item => item.Y);
                target.TextObjects.Add(new(id, "Paragraph", string.Concat(paragraph.Descendants().Where(item => item.Ancestors(W + "p").FirstOrDefault() == paragraph && !item.Ancestors(W + "del").Any()).Select(item => item.Name == W + "t" ? item.Value : item.Name == W + "br" ? "\n" : item.Name == W + "tab" ? "\t" : "")), new(minX, minY, runs.Max(item => item.X + item.Width) - minX, runs.Max(item => item.Y + item.Height) - minY, inTable || result.Pages.Count(item => item.Runs.Any(run => run.Id == id)) > 1)));
            }
        }
        if (frame is not null)
        {
            y = flowingY;
        }
    }

    private void Image(XElement picture, double x, double available, string? alignment = null)
    {
        if (!imageIds.TryGetValue(picture, out var id) || !picture.Descendants(A + "blip").Any())
        {
            warnings.Add("Unsupported drawing content is omitted from layout.");
            return;
        }
        try
        {
            var part = ResolveImage(picture);
            var bytes = document.GetPart(part);
            if (EmfImage.IsEmf(bytes))
            {
                var requestedWidth = Number(picture.Element(WP + "extent"), "cx", 1270000) / 12700;
                var requestedHeight = Number(picture.Element(WP + "extent"), "cy", 1270000) / 12700;
                if (!decodedEmf.TryGetValue(part, out var decoded) || decoded.Width < requestedWidth || decoded.Height < requestedHeight)
                {
                    requestedWidth = Math.Max(requestedWidth, decoded.Width);
                    requestedHeight = Math.Max(requestedHeight, decoded.Height);
                    decoded = (EmfImage.Decode(bytes, requestedWidth, requestedHeight), requestedWidth, requestedHeight);
                    decodedEmf[part] = decoded;
                    result.Media[part] = "data:image/png;base64," + Convert.ToBase64String(decoded.Bytes);
                }
                bytes = decoded.Bytes;
            }
            using var imageData = SKData.CreateCopy(bytes);
            using var codec = SKCodec.Create(imageData);
            if (codec is null || codec.Info.Width * (long)codec.Info.Height > 64000000)
            {
                warnings.Add("A picture format is unsupported or too large to decode: " + part);
                return;
            }
            var width = Math.Max(1, Number(picture.Element(WP + "extent"), "cx", codec.Info.Width * 12700) / 12700);
            var height = Math.Max(1, Number(picture.Element(WP + "extent"), "cy", codec.Info.Height * 12700) / 12700);
            var imageY = y;
            if (picture.Name == WP + "inline")
            {
                var effects = picture.Element(WP + "effectExtent");
                var effectLeft = Math.Max(0, Number(effects, "l") / 12700);
                var effectRight = Math.Max(0, Number(effects, "r") / 12700);
                var effectTop = Math.Max(0, Number(effects, "t") / 12700);
                var effectBottom = Math.Max(0, Number(effects, "b") / 12700);
                var outerWidth = width + effectLeft + effectRight;
                var outerHeight = height + effectTop + effectBottom;
                var scale = Math.Min(1, Math.Min(available / outerWidth, (result.Height - top - bottom) / outerHeight));
                width *= scale;
                height *= scale;
                outerWidth *= scale;
                outerHeight *= scale;
                if (!measuringTable && y + outerHeight > FlowLimit)
                {
                    var previousLeft = FlowLeft;
                    AdvanceFlow();
                    x += FlowLeft - previousLeft;
                    var columnScale = Math.Min(1, FlowWidth / outerWidth);
                    width *= columnScale;
                    height *= columnScale;
                    outerWidth *= columnScale;
                    outerHeight *= columnScale;
                    scale *= columnScale;
                    available = FlowWidth;
                }
                x += effectLeft * scale + (alignment == "center" ? Math.Max(0, (available - outerWidth) / 2)
                    : alignment is "right" or "end" ? Math.Max(0, available - outerWidth) : 0);
                imageY = y + effectTop * scale;
                y += outerHeight;
            }
            else
            {
                x = DrawingPosition(picture, "positionH", x, width);
                imageY = DrawingPosition(picture, "positionV", y, height);
            }
            var crop = picture.Descendants(A + "srcRect").FirstOrDefault();
            double Crop(string side) => double.TryParse((string?)crop?.Attribute(side), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var value) && double.IsFinite(value) ? Math.Clamp(value / 100000, -10, 1) : 0;
            var cropLeft = Crop("l");
            var cropRight = Crop("r");
            var cropTop = Crop("t");
            var cropBottom = Crop("b");
            if (1 - cropLeft - cropRight <= .00001 || 1 - cropTop - cropBottom <= .00001)
            {
                warnings.Add("An image has an invalid crop rectangle and is omitted.");
                return;
            }
            var transform = picture.Descendants(A + "xfrm").FirstOrDefault();
            var rotation = double.TryParse((string?)transform?.Attribute("rot"), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var angle) && double.IsFinite(angle) ? angle / 60000 % 360 : 0;
            var flipH = (string?)transform?.Attribute("flipH") is "1" or "true";
            var flipV = (string?)transform?.Attribute("flipV") is "1" or "true";
            var mime = codec.EncodedFormat switch
            {
                SKEncodedImageFormat.Png => "image/png",
                SKEncodedImageFormat.Jpeg => "image/jpeg",
                SKEncodedImageFormat.Webp => "image/webp",
                SKEncodedImageFormat.Gif => "image/gif",
                SKEncodedImageFormat.Bmp => "image/bmp",
                _ => "application/octet-stream"
            };
            result.Media.TryAdd(part, "data:" + mime + ";base64," + Convert.ToBase64String(bytes));
            if (IsAbsoluteDecoration(picture))
            {
                absoluteDecorations.Add(id);
            }
            page.Images.Add(new(id, (string?)picture.Element(WP + "docPr")?.Attribute("name") ?? "Image", part, x, imageY, width, height, Rotation: rotation, CropLeft: cropLeft, CropTop: cropTop, CropRight: cropRight, CropBottom: cropBottom, FlipH: flipH, FlipV: flipV));
        }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException)
        {
            warnings.Add(exception.Message);
        }
    }

    public void Dispose()
    {
        foreach (var font in fonts.Values)
        {
            font.Dispose();
        }
        fallback.Dispose();
    }
}
