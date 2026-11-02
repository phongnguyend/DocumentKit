using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using PptxObjectKit.Readers;
using PptxObjectKit.Extensions.Pdf;
using PdfObjectKit.Readers;
using PdfObjectKit.Extractors;
using PdfObjectKit.Rendering;
using PptxObjectKit.Editors;
using SkiaSharp;

var checks = 0;
void Check(bool value, string message)
{
    if (!value)
    {
        throw new Exception(message);
    }
    checks++;
}
void Reject(Action action)
{
    try
    {
        action();
    }
    catch (Exception exception) when (exception is ArgumentException or NotSupportedException or OperationCanceledException or InvalidDataException)
    {
        checks++;
        return;
    }
    throw new Exception("Expected invalid/unsupported conversion to be rejected.");
}
var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "sample.pptx"));
var document = PptxReader.Read(bytes);
var before = document.PartNames.ToDictionary(name => name, document.GetPart);
var pdf = PptxPdfConverter.Convert(document);
Check(pdf.PageCount == 2 && Encoding.ASCII.GetString(pdf.Bytes, 0, 5) == "%PDF-", "A valid PDF with two pages is produced.");
var parsed = PdfReader.Read(pdf.Bytes);
Check(PdfTextSearcher.Search(parsed, "PowerPoint").Matches.Count > 0, "Text is searchable, not a page screenshot.");
Check(PdfTextSearcher.Search(parsed, "Second slide").Matches.Count > 0, "All slides retain searchable text.");
Check(PdfParagraphExtractor.ExtractSelectionLines(parsed, 1).Count > 0, "PDF text exposes selectable lines.");
Check(before.All(item => item.Value.SequenceEqual(document.GetPart(item.Key))), "Conversion does not mutate the PPTX.");
var reversed = PptxPdfConverter.Convert(document, new() { Slides = [2, 1] });
Check(PdfTextSearcher.Search(PdfReader.Read(reversed.Bytes), "Second slide").Matches[0].PageNumber == 1, "Explicit slide order is respected.");
Check(PptxPdfConverter.Convert(document, new() { Slides = [2] }).PageCount == 1, "Single-slide export.");
Reject(() => PptxPdfConverter.Convert(document, new() { Slides = [] }));
Reject(() => PptxPdfConverter.Convert(document, new() { Slides = [0] }));
Reject(() => PptxPdfConverter.Convert(document, new() { Slides = [3] }));
Reject(() => PptxPdfConverter.Convert(document, new() { Slides = [1, 1] }));
Reject(() => PptxPdfConverter.Convert(document, cancellationToken: new CancellationToken(true)));

XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
PptxObjectKit.Models.PptxObjectDocument Change(Action<XDocument> edit, Action<XDocument>? editDefaults = null)
{
    using var memory = new MemoryStream();
    memory.Write(bytes);
    using (var zip = new ZipArchive(memory, ZipArchiveMode.Update, true))
    {
        var entry = zip.GetEntry("ppt/slides/slide1.xml")!;
        XDocument xml;
        using (var input = entry.Open())
        {
            xml = XDocument.Load(input);
        }
        edit(xml);
        entry.Delete();
        using var output = zip.CreateEntry("ppt/slides/slide1.xml").Open();
        xml.Save(output);
        if (editDefaults is not null)
        {
            var presentationEntry = zip.GetEntry("ppt/presentation.xml")!;
            XDocument presentation;
            using (var input = presentationEntry.Open())
            {
                presentation = XDocument.Load(input);
            }
            editDefaults(presentation);
            presentationEntry.Delete();
            using var presentationOutput = zip.CreateEntry("ppt/presentation.xml").Open();
            presentation.Save(presentationOutput);
        }
    }
    return PptxReader.Read(memory.ToArray());
}
var hidden = Change(xml => xml.Root!.SetAttributeValue("show", "0"));
Check(PptxPdfConverter.Convert(hidden, new() { IncludeHiddenSlides = false }).PageCount == 1, "Hidden slides can be excluded.");
Reject(() => PptxPdfConverter.Convert(hidden, new() { Slides = [1], IncludeHiddenSlides = false }));
var unsupported = Change(xml =>
{
    var props = xml.Descendants(p + "spPr").First();
    props.Elements(a + "prstGeom").Remove();
    props.Add(new XElement(a + "prstGeom", new XAttribute("prst", "star24")));
});
Check(PptxPdfConverter.Convert(unsupported).Warnings.Any(warning => warning.Contains("star24")), "Unsupported geometry is reported.");
Reject(() => PptxPdfConverter.Convert(unsupported, new() { FailOnWarnings = true }));
var missingFont = Change(xml =>
{
    foreach (var run in xml.Descendants(a + "r"))
    {
        var style = run.Element(a + "rPr");
        if (style is null)
        {
            style = new XElement(a + "rPr");
            run.AddFirst(style);
        }
        style.Elements(a + "latin").Remove();
        style.Add(new XElement(a + "latin", new XAttribute("typeface", "MissingFontForNativePdfTest")));
    }
});
var substituted = PptxPdfConverter.Convert(missingFont);
Check(substituted.Warnings.Any(warning => warning.Contains("MissingFontForNativePdfTest")), "Missing fonts report substitution.");
Check(PdfTextSearcher.Search(PdfReader.Read(substituted.Bytes), "PowerPoint").Matches.Count > 0, "Bundled fallback font retains Unicode extraction.");
var unicode = Change(xml => xml.Descendants(a + "t").First().Value = "Ti\u1ebfng Vi\u1ec7t caf\u00e9");
Check(PdfTextSearcher.Search(PdfReader.Read(PptxPdfConverter.Convert(unicode).Bytes), "caf\u00e9").Matches.Count > 0, "Unicode survives PDF text extraction.");
var invalid = Change(xml => xml.Descendants(a + "ext").First().SetAttributeValue("cx", "NaN"));
Reject(() => PptxPdfConverter.Convert(invalid));
var updated = PptxSlideEditor.UpdateText(document, 1, "2", "Current edited text");
Check(PdfTextSearcher.Search(PdfReader.Read(PptxPdfConverter.Convert(updated).Bytes), "Current edited text").Matches.Count > 0, "Current edits are exported.");
var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
var withImage = PptxImageEditor.Edit(document, 1, "", "add-image", png, 20, 30, 40, 50);
var imageCommands = PdfCanvasExporter.Export(PdfReader.Read(PptxPdfConverter.Convert(withImage).Bytes), 1).Commands;
Check(imageCommands.Any(command => command.Kind is "jpeg" or "rgba"), "Pictures remain image objects inside a vector PDF.");
var colored = Change(xml =>
{
    var style = xml.Descendants(a + "rPr").First();
    style.Add(new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "FF0000"))));
    xml.Descendants(a + "xfrm").First().SetAttributeValue("rot", 90 * 60000);
});
var commands = PdfCanvasExporter.Export(PdfReader.Read(PptxPdfConverter.Convert(colored).Bytes), 1).Commands;
Check(commands.Any(command => command.Kind == "rg" && command.Numbers.SequenceEqual(new double[] { 1, 0, 0 })), "Text color survives export.");
Check(commands.Any(command => command.Kind == "cm" && Math.Abs(command.Numbers[1]) > .9), "Rotation survives export.");
var table = Change(xml =>
{
    var transform = new XElement(p + "xfrm",
        new XElement(a + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
        new XElement(a + "ext", new XAttribute("cx", 5080000), new XAttribute("cy", 1270000)));
    var grid = new XElement(a + "tblGrid",
        new XElement(a + "gridCol", new XAttribute("w", 2540000)),
        new XElement(a + "gridCol", new XAttribute("w", 2540000)));
    var row = new XElement(a + "tr", new XAttribute("h", 1270000), Cell("TableLeft"), Cell("TableRight"));
    var content = new XElement(a + "tbl", new XElement(a + "tblPr"), grid, row);
    var graphic = new XElement(p + "graphicFrame", transform, new XElement(a + "graphic", new XElement(a + "graphicData", content)));
    xml.Descendants(p + "spTree").First().Add(graphic);
    XElement Cell(string text) => new(a + "tc", new XElement(a + "txBody", new XElement(a + "bodyPr"), new XElement(a + "p", new XElement(a + "r", new XElement(a + "rPr", new XAttribute("sz", 1800)), new XElement(a + "t", text)))), new XElement(a + "tcPr", new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "00FF00")))));
});
var tablePdf = PdfReader.Read(PptxPdfConverter.Convert(table).Bytes);
Check(PdfTextSearcher.Search(tablePdf, "TableLeft").Matches.Count == 1 && PdfTextSearcher.Search(tablePdf, "TableRight").Matches.Count == 1, "Table cells retain searchable text.");
Check(PdfCanvasExporter.Export(tablePdf, 1).Commands.Any(command => command.Kind == "rg" && command.Numbers.SequenceEqual(new double[] { 0, 1, 0 })), "Table cell fill is preserved.");
// Diagram pixels must not be degraded by implicit quality-zero JPEG encoding.
using var bitmap = new SKBitmap(32, 32);
for (var y = 0; y < bitmap.Height; y++)
{
    for (var x = 0; x < bitmap.Width; x++)
    {
        bitmap.SetPixel(x, y, (x + y) % 2 == 0 ? SKColors.Red : SKColors.White);
    }
}
using var encoded = bitmap.Encode(SKEncodedImageFormat.Png, 100);
var diagram = PptxImageEditor.Edit(document, 1, "", "add-image", encoded.ToArray());
var diagramPdf = PdfReader.Read(PptxPdfConverter.Convert(diagram).Bytes);
var diagramImage = PdfCanvasExporter.ExportImage(diagramPdf, 1, 1);
Check(diagramImage.Kind == "rgba" && diagramImage.Width == 32 && diagramImage.Height == 32, "PNG diagrams retain source resolution and lossless encoding.");
var pixels = Convert.FromBase64String(diagramImage.Data!);
for (var y = 0; y < 32; y++)
{
    for (var x = 0; x < 32; x++)
    {
        var expected = bitmap.GetPixel(x, y);
        var index = (y * 32 + x) * 4;
        if (pixels[index] != expected.Red || pixels[index + 1] != expected.Green || pixels[index + 2] != expected.Blue || pixels[index + 3] != 255)
        {
            throw new Exception("Lossless image pixels changed during PDF export.");
        }
    }
}
Check(true, "Every exported diagram pixel matches the source.");
var gradient = Change(xml =>
{
    var body = xml.Descendants(p + "txBody").First();
    body.Element(a + "lstStyle")!.Add(new XElement(a + "lvl1pPr", new XElement(a + "defRPr",
        new XElement(a + "gradFill", new XElement(a + "gsLst",
            new XElement(a + "gs", new XAttribute("pos", 0), new XElement(a + "srgbClr", new XAttribute("val", "800080"))),
            new XElement(a + "gs", new XAttribute("pos", 100000), new XElement(a + "srgbClr", new XAttribute("val", "FF0000")))),
            new XElement(a + "lin", new XAttribute("ang", 0))))));
    xml.Descendants(p + "sp").First().Add(new XElement(p + "style", new XElement(a + "fontRef", new XAttribute("idx", "minor"), new XElement(a + "srgbClr", new XAttribute("val", "000000")))));
});
var gradientBytes = PptxPdfConverter.Convert(gradient).Bytes;
var gradientPdf = PdfReader.Read(gradientBytes);
Check(PdfTextSearcher.Search(gradientPdf, "PowerPoint").Matches.Count > 0, "Gradient text remains searchable.");
bool HasShading(PdfObjectKit.Models.PdfValue value) => value.Entries?.ContainsKey("ShadingType") == true
    || (value.Entries?.Values.Any(HasShading) ?? false)
    || (value.Items?.Any(HasShading) ?? false);
Check(gradientPdf.Objects.Any(item => HasShading(item.Value)), "Inherited gradient text emits a PDF gradient instead of black text.");
File.WriteAllBytes(Path.Combine(AppContext.BaseDirectory, "gradient.pdf"), gradientBytes);
double BaselineDistance(int reduction, int? percentage = null, int? points = null)
{
    var source = Change(xml =>
    {
        var body = xml.Descendants(p + "txBody").First();
        body.RemoveNodes();
        body.Add(new XElement(a + "bodyPr", new XElement(a + "normAutofit", new XAttribute("lnSpcReduction", reduction))), new XElement(a + "lstStyle"));
        foreach (var text in new[] { "SpacingFirst", "SpacingSecond" })
        {
            var properties = new XElement(a + "pPr", new XElement(a + "buNone"),
                new XElement(a + "defRPr", new XAttribute("sz", 2000)),
                new XElement(a + "spcAft", new XElement(a + "spcPts", new XAttribute("val", 500))));
            if (percentage is not null || points is not null)
            {
                properties.Add(new XElement(a + "lnSpc", new XElement(a + (points is not null ? "spcPts" : "spcPct"), new XAttribute("val", points ?? percentage!.Value))));
            }
            body.Add(new XElement(a + "p", properties, new XElement(a + "r", new XElement(a + "rPr", new XAttribute("sz", 2000)), new XElement(a + "t", text))));
        }
    });
    var selections = PdfParagraphExtractor.ExtractSelectionLines(PdfReader.Read(PptxPdfConverter.Convert(source).Bytes), 1);
    return Math.Abs(selections.Single(item => item.Text.Trim() == "SpacingFirst").Baseline - selections.Single(item => item.Text.Trim() == "SpacingSecond").Baseline);
}
Check(Math.Abs((BaselineDistance(10000) - 5) / (BaselineDistance(0) - 5) - .9) < .001, "Autofit reduces default line spacing without reducing paragraph spacing.");
Check(Math.Abs((BaselineDistance(10000, 120000) - 5) / (BaselineDistance(0, 120000) - 5) - 11.0 / 12) < .001, "Autofit subtracts percentage points from explicit proportional line spacing.");
Check(Math.Abs(BaselineDistance(10000, points: 3000) - BaselineDistance(0, points: 3000)) < .001, "Autofit preserves explicit fixed point line spacing.");
var symbols = Change(xml =>
{
    var run = xml.Descendants(a + "r").First();
    run.Element(a + "rPr")!.Add(new XElement(a + "sym", new XAttribute("typeface", "Wingdings")));
    run.Element(a + "t")!.Value = "\uf0e8Unit testing!? \uf0e8 hard to replace";
});
var symbolResult = PptxPdfConverter.Convert(symbols);
var symbolPdf = PdfReader.Read(symbolResult.Bytes);
Check(PdfTextSearcher.Search(symbolPdf, "Unit testing").Matches.Count == 1 && PdfTextSearcher.Search(symbolPdf, "hard to replace").Matches.Count == 1, "Wingdings symbols do not change adjacent Latin text.");
Check(PdfTextSearcher.Search(symbolPdf, "\u2794").Matches.Count == 2 || symbolResult.Warnings.Any(item => item.Contains("U+2794")), "Wingdings arrows use searchable Unicode glyphs or explicitly report unavailable system glyphs.");
var literalPua = Change(xml => xml.Descendants(a + "t").First().Value = "\uf0e8");
Check(PdfTextSearcher.Search(PdfReader.Read(PptxPdfConverter.Convert(literalPua).Bytes), "\u2794").Matches.Count == 0, "PUA characters outside Wingdings runs are not remapped.");
var styleReference = Change(xml =>
{
    xml.Descendants(p + "sp").First().Add(new XElement(p + "style", new XElement(a + "fontRef", new XAttribute("idx", "minor"), new XElement(a + "srgbClr", new XAttribute("val", "0000FF")))));
}, presentation =>
{
    presentation.Root!.Elements(p + "defaultTextStyle").Remove();
    presentation.Root.Add(new XElement(p + "defaultTextStyle", new XElement(a + "lvl1pPr", new XElement(a + "defRPr", new XElement(a + "solidFill", new XElement(a + "srgbClr", new XAttribute("val", "FFFFFF")))))));
});
var styleCommands = PdfCanvasExporter.Export(PdfReader.Read(PptxPdfConverter.Convert(styleReference).Bytes), 1).Commands;
Check(styleCommands.Any(command => command.Kind == "rg" && command.Numbers.SequenceEqual(new double[] { 0, 0, 1 })), "Shape font references override global text fill defaults.");
using var fallbackStream = typeof(PptxPdfConverter).Assembly.GetManifestResourceStream("PptxObjectKit.Extensions.Pdf.Fonts.NotoSans-Regular.ttf")!;
using var fallbackData = SKData.Create(fallbackStream);
using var fallbackFace = SKTypeface.FromData(fallbackData);
using var wrappingFont = new SKFont(fallbackFace, 20);
var wrapping = Change(xml =>
{
    var shape = xml.Descendants(p + "sp").First();
    shape.Descendants(a + "ext").First().SetAttributeValue("cx", (long)((wrappingFont.MeasureText("AA too") + wrappingFont.MeasureText(".") / 2) * 12700));
    var body = shape.Element(p + "txBody")!;
    body.RemoveNodes();
    body.Add(new XElement(a + "bodyPr", new XAttribute("lIns", 0), new XAttribute("rIns", 0)), new XElement(a + "lstStyle"));
    XElement Run(string text) => new(a + "r", new XElement(a + "rPr", new XAttribute("sz", 2000), new XElement(a + "latin", new XAttribute("typeface", "MissingFontForNativePdfTest"))), new XElement(a + "t", text));
    body.Add(new XElement(a + "p", new XElement(a + "pPr", new XElement(a + "buNone")), Run("AA too"), Run(".")));
});
var wrappedLines = PdfParagraphExtractor.ExtractSelectionLines(PdfReader.Read(PptxPdfConverter.Convert(wrapping).Bytes), 1);
Check(wrappedLines.Any(line => line.Text.Trim() == "too.") && wrappedLines.All(line => line.Text.Trim() != "."), "Punctuation split into a separate run stays with its word when wrapping.");
var sadFace = Change(xml =>
{
    var run = xml.Descendants(a + "r").First();
    run.Element(a + "rPr")!.Add(new XElement(a + "sym", new XAttribute("typeface", "Wingdings")));
    run.Element(a + "t")!.Value = "\uf04c";
});
var sadResult = PptxPdfConverter.Convert(sadFace);
var sadPdf = PdfReader.Read(sadResult.Bytes);
if (sadResult.Warnings.All(warning => !warning.Contains("U+2639")))
{
    Check(PdfTextSearcher.Search(sadPdf, "\u2639").Matches.Count == 1, "Wingdings sad face remains searchable.");
    if (SKFontManager.Default.FontFamilies.Contains("Segoe UI Symbol"))
    {
        Check(PdfCanvasExporter.Export(sadPdf, 1).Commands.All(command => command.Kind is not ("rgba" or "jpeg")), "Text symbol fallback does not rasterize a colored emoji.");
    }
}
File.WriteAllBytes(Path.Combine(AppContext.BaseDirectory, "converted.pdf"), pdf.Bytes);
Console.WriteLine($"Passed {checks} native PDF conversion checks.");
