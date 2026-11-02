using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using PdfObjectKit.Models;
using PdfObjectKit.Rendering;

namespace PdfObjectKit.Converters;

internal sealed class DocxPackage : IDisposable
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private readonly MemoryStream buffer = new();
    private readonly ZipArchive zip;
    private readonly XElement body = new(W + "body");
    private readonly List<XElement> relationships = [Relation("styles", "styles", "styles.xml")];
    private long packageBytes;

    internal int ParagraphCount { get; private set; }

    internal int TableCount { get; private set; }

    internal int ImageCount { get; private set; }

    internal DocxPackage()
    {
        zip = new(buffer, ZipArchiveMode.Create, true);
    }

    private static XElement Value(string name, object value) => new(W + name, new XAttribute(W + "val", value));

    private static XElement Relation(string id, string type, string target) => new(Rel + "Relationship",
        new XAttribute("Id", id), new XAttribute("Type", R.NamespaceName + "/" + type), new XAttribute("Target", target));

    internal void PageBreak() => body.Add(new XElement(W + "p", new XElement(W + "r", new XElement(W + "br", new XAttribute(W + "type", "page")))));

    internal void Paragraph(PdfContentBlock block)
    {
        var text = block.Kind == PdfContentBlockKind.ListItem ? (block.ListNumber is null ? "• " : block.ListNumber + ". ") + block.Text : block.Text;
        var paragraph = TextParagraph(text, block.SourceLines);
        if (block.Kind == PdfContentBlockKind.Heading)
        {
            paragraph.AddFirst(new XElement(W + "pPr", Value("pStyle", "Heading" + Math.Clamp(block.HeadingLevel, 1, 6))));
        }
        body.Add(paragraph);
        ParagraphCount++;
    }

    private static XElement TextParagraph(string text, IEnumerable<PdfParagraphLine> lines)
    {
        var paragraph = new XElement(W + "p");
        var offset = 0;
        foreach (var run in lines.SelectMany(line => line.Runs))
        {
            if (string.IsNullOrEmpty(run.Text))
            {
                continue;
            }
            var start = text.IndexOf(run.Text, offset, StringComparison.Ordinal);
            if (start < 0)
            {
                continue;
            }
            if (start > offset)
            {
                paragraph.Add(TextRun(text[offset..start], run));
            }
            paragraph.Add(TextRun(run.Text, run));
            offset = start + run.Text.Length;
        }
        if (offset < text.Length || !paragraph.HasElements)
        {
            paragraph.Add(TextRun(text[offset..], null));
        }
        return paragraph;
    }

    private static XElement TextRun(string text, PdfTextRun? source)
    {
        var properties = new XElement(W + "rPr");
        if (source is not null)
        {
            var font = Regex.Replace(source.FontName ?? "Arial", "^[A-Z]{6}\\+", "");
            properties.Add(new XElement(W + "rFonts", new XAttribute(W + "ascii", font), new XAttribute(W + "hAnsi", font)));
            if (font.Contains("Bold", StringComparison.OrdinalIgnoreCase))
            {
                properties.Add(new XElement(W + "b"));
            }
            if (font.Contains("Italic", StringComparison.OrdinalIgnoreCase) || font.Contains("Oblique", StringComparison.OrdinalIgnoreCase))
            {
                properties.Add(new XElement(W + "i"));
            }
            var size = double.IsFinite(source.FontSize) ? (int)Math.Clamp(Math.Round(source.FontSize * 2), 2, 400) : 24;
            properties.Add(Value("sz", size), Value("szCs", size));
        }
        var safe = string.Concat(text.EnumerateRunes().Where(rune => rune.Value is 9 or 10 or 13 || rune.Value >= 32 && rune.Value is not (0xfffe or 0xffff)));
        return new(W + "r", properties, new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), safe.Replace('\n', ' ').Replace('\r', ' ')));
    }

    internal void Table(PdfTable source)
    {
        var columns = source.Rows.Select(row => row.Cells.Sum(cell => Math.Max(1, cell.ColumnSpan))).DefaultIfEmpty(1).Max();
        var table = new XElement(W + "tbl",
            new XElement(W + "tblPr", new XElement(W + "tblW", new XAttribute(W + "w", 5000), new XAttribute(W + "type", "pct")),
                new XElement(W + "tblBorders", new[] { "top", "left", "bottom", "right", "insideH", "insideV" }.Select(edge =>
                    new XElement(W + edge, new XAttribute(W + "val", "single"), new XAttribute(W + "sz", 4), new XAttribute(W + "color", "B8C2CC"))))),
            new XElement(W + "tblGrid", Enumerable.Range(0, columns).Select(_ => new XElement(W + "gridCol", new XAttribute(W + "w", 9360 / columns)))));
        foreach (var row in source.Rows)
        {
            var output = new XElement(W + "tr");
            if (row.IsHeader)
            {
                output.Add(new XElement(W + "trPr", new XElement(W + "tblHeader")));
            }
            foreach (var cell in row.Cells)
            {
                var properties = new XElement(W + "tcPr");
                if (cell.ColumnSpan > 1)
                {
                    properties.Add(Value("gridSpan", cell.ColumnSpan));
                }
                output.Add(new XElement(W + "tc", properties, TextParagraph(cell.Text, cell.SourceLines)));
            }
            table.Add(output);
        }
        body.Add(table, new XElement(W + "p"));
        TableCount++;
    }

    internal void Image(PdfObjectDocument document, int page, PdfImageOccurrence source)
    {
        var command = PdfCanvasExporter.ExportImage(document, page, source.ImageNumber);
        var bytes = Convert.FromBase64String(command.Data ?? throw new InvalidDataException("Missing image bytes."));
        var extension = "png";
        if (command.Kind == "rgba")
        {
            bytes = DocxPng.Encode(bytes, command.Width, command.Height);
        }
        else if (command.Kind == "jpeg")
        {
            if (command.AlphaData is not null)
            {
                bytes = DocxMaskedImage.Decode(command, bytes);
            }
            else
            {
                extension = "jpg";
            }
        }
        else
        {
            throw new NotSupportedException("Unsupported image encoding.");
        }
        // Fit full source pixels within the Word text area, preserving their aspect ratio.
        var width = source.Corners.Count == 4 ? Math.Sqrt(Math.Pow(source.Corners[1].X - source.Corners[0].X, 2) + Math.Pow(source.Corners[1].Y - source.Corners[0].Y, 2)) : command.Width * .75;
        width = double.IsFinite(width) && width > 0 ? Math.Min(width, 468) : 468;
        var height = width * command.Height / command.Width;
        if (height > 680)
        {
            width *= 680 / height;
            height = 680;
        }
        var cx = (long)Math.Round(width * 12700);
        var cy = (long)Math.Round(height * 12700);
        var id = ImageCount + 1;
        var name = $"image{id}.{extension}";
        Write("word/media/" + name, bytes);
        ImageCount = id;
        relationships.Add(Relation("image" + id, "image", "media/" + name));
        var picture = new XElement(Pic + "pic",
            new XElement(Pic + "nvPicPr", new XElement(Pic + "cNvPr", new XAttribute("id", id), new XAttribute("name", name)), new XElement(Pic + "cNvPicPr")),
            new XElement(Pic + "blipFill", new XElement(A + "blip", new XAttribute(R + "embed", "image" + id)), new XElement(A + "stretch", new XElement(A + "fillRect"))),
            new XElement(Pic + "spPr", new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(A + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy))),
                new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))));
        body.Add(new XElement(W + "p", new XElement(W + "r", new XElement(W + "drawing",
            new XElement(Wp + "inline", new XElement(Wp + "extent", new XAttribute("cx", cx), new XAttribute("cy", cy)),
                new XElement(Wp + "docPr", new XAttribute("id", id), new XAttribute("name", name)),
                new XElement(A + "graphic", new XElement(A + "graphicData", new XAttribute("uri", Pic.NamespaceName), picture)))))));
    }

    internal byte[] Finish()
    {
        body.Add(new XElement(W + "sectPr", new XElement(W + "pgSz", new XAttribute(W + "w", 12240), new XAttribute(W + "h", 15840)),
            new XElement(W + "pgMar", new XAttribute(W + "top", 1080), new XAttribute(W + "right", 1440), new XAttribute(W + "bottom", 1080),
                new XAttribute(W + "left", 1440), new XAttribute(W + "header", 0), new XAttribute(W + "footer", 0), new XAttribute(W + "gutter", 0))));
        Xml("word/document.xml", new XElement(W + "document", new XAttribute(XNamespace.Xmlns + "w", W), new XAttribute(XNamespace.Xmlns + "r", R), body));
        var styles = new XElement(W + "styles", new XElement(W + "docDefaults",
            new XElement(W + "rPrDefault", new XElement(W + "rPr", Value("sz", 24))),
            new XElement(W + "pPrDefault", new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "after", 120))))));
        for (var level = 1; level <= 6; level++)
        {
            styles.Add(new XElement(W + "style", new XAttribute(W + "type", "paragraph"), new XAttribute(W + "styleId", "Heading" + level),
                Value("name", "heading " + level), new XElement(W + "pPr", new XElement(W + "keepNext"), Value("outlineLvl", level - 1)),
                new XElement(W + "rPr", new XElement(W + "b"))));
        }
        Xml("word/styles.xml", styles);
        Xml("word/_rels/document.xml.rels", new XElement(Rel + "Relationships", relationships));
        Xml("_rels/.rels", new XElement(Rel + "Relationships", Relation("document", "officeDocument", "word/document.xml")));
        Xml("[Content_Types].xml", new XElement(Ct + "Types",
            new XElement(Ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
            new XElement(Ct + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png")),
            new XElement(Ct + "Default", new XAttribute("Extension", "jpg"), new XAttribute("ContentType", "image/jpeg")),
            new XElement(Ct + "Override", new XAttribute("PartName", "/word/document.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")),
            new XElement(Ct + "Override", new XAttribute("PartName", "/word/styles.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"))));
        zip.Dispose();
        return buffer.ToArray();
    }

    private void Xml(string name, XElement element) => Write(name, Encoding.UTF8.GetBytes(element.ToString(SaveOptions.DisableFormatting)));

    private void Write(string name, byte[] bytes)
    {
        packageBytes += bytes.Length;
        if (packageBytes > 128L * 1024 * 1024)
        {
            throw new NotSupportedException("Word export exceeds the 128 MB uncompressed package limit. Export fewer pages.");
        }
        using var stream = zip.CreateEntry(name, CompressionLevel.Optimal).Open();
        stream.Write(bytes);
    }

    public void Dispose()
    {
        zip.Dispose();
        buffer.Dispose();
    }
}
