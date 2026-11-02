using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PdfObjectKit.Converters;
using PdfObjectKit.Editors;
using PdfObjectKit.Graphs;
using PdfObjectKit.Readers;

static void Check(bool condition, string message)
{
    if (!condition)
    {
        throw new Exception(message);
    }
}

static XDocument Validate(byte[] bytes)
{
    using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
    var errors = new OpenXmlValidator().Validate(document).ToArray();
    Check(errors.Length == 0, string.Join("\n", errors.Select(error => error.Description)));
    using var zip = new ZipArchive(new MemoryStream(bytes));
    using var stream = zip.GetEntry("word/document.xml")!.Open();
    return XDocument.Load(stream);
}

var pdf = PdfReader.Read(PdfFixture.Create(null, new("Editable & safe <text>", Y: 700), new("Second page", Y: 700)));
PdfImageEditor.InsertImage(pdf, 1, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="), 30, 300, 100, 50);
var original = PdfConverter.ToJson(pdf);
var result = PdfDocxConverter.Convert(pdf);
Check(result.SourcePageCount == 2 && result.ParagraphCount == 2 && result.ImageCount == 1, "Content counts");
Check(PdfConverter.ToJson(pdf) == original, "Source document unchanged");
XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
var xml = Validate(result.Bytes);
Check(xml.Descendants(w + "t").Any(text => text.Value == "Editable & safe <text>"), "Escaped editable text");
Check(xml.Descendants(w + "drawing").Count() == 1, "Embedded editable image");
Check(xml.Descendants(w + "br").Count() == 1, "Source page boundary");
Check(xml.Descendants(w + "sz").Any(size => (string?)size.Attribute(w + "val") == "24"), "Font size preserved");
var ordered = PdfDocxConverter.Convert(pdf, new() { PageNumbers = [2, 1], IncludeImages = false, PageBreaks = false });
var orderedXml = Validate(ordered.Bytes);
Check(orderedXml.Descendants(w + "t").First().Value == "Second page", "Page ordering");
Check(!orderedXml.Descendants(w + "drawing").Any() && !orderedXml.Descendants(w + "br").Any(), "Export options");
Check(Validate(PdfDocxConverter.Convert(pdf, new() { PageNumbers = [2] }).Bytes).Descendants(w + "t").All(text => !text.Value.Contains("Editable")), "Page selection");
foreach (var pages in new[] { Array.Empty<int>(), new[] { 0 }, new[] { 3 }, new[] { 1, 1 } })
{
    try
    {
        PdfDocxConverter.Convert(pdf, new() { PageNumbers = pages });
        throw new Exception("Invalid pages accepted");
    }
    catch (ArgumentException)
    {
    }
}
using var cancelled = new CancellationTokenSource();
cancelled.Cancel();
try
{
    PdfDocxConverter.Convert(pdf, cancellationToken: cancelled.Token);
    throw new Exception("Cancelled export succeeded");
}
catch (OperationCanceledException)
{
}
var tablePdf = PdfReader.Read(PdfFixture.Create(null, new FixturePage()));
new PdfGraph(tablePdf).Pages()[0].Streams[0].Value.Base64 = Convert.ToBase64String(Encoding.ASCII.GetBytes(
    "BT /F1 12 Tf 1 0 0 1 30 700 Tm (Item) Tj 1 0 0 1 240 700 Tm (Amount) Tj " +
    "1 0 0 1 30 680 Tm (Alpha) Tj 1 0 0 1 240 680 Tm (100) Tj " +
    "1 0 0 1 30 660 Tm (Beta) Tj 1 0 0 1 240 660 Tm (200) Tj ET"));
var table = PdfDocxConverter.Convert(tablePdf);
Check(table.TableCount == 1, "Heuristic editable table");
Check(Validate(table.Bytes).Descendants(w + "tc").Count() == 6, "Table cells");
var styled = PdfReader.Read(PdfFixture.Create(null, new FixturePage()));
var styledGraph = new PdfGraph(styled);
var styledPage = styledGraph.Pages()[0];
styledGraph.Get(styledPage.Resources, "Font")!.Entries!["F2"] = PdfObjectKit.Primitives.PdfPrimitives.Dict(
    ("Type", PdfObjectKit.Primitives.PdfPrimitives.Name("Font")),
    ("Subtype", PdfObjectKit.Primitives.PdfPrimitives.Name("Type1")),
    ("BaseFont", PdfObjectKit.Primitives.PdfPrimitives.Name("Helvetica-BoldOblique")));
styledPage.Streams[0].Value.Base64 = Convert.ToBase64String(Encoding.ASCII.GetBytes(
    "BT /F2 24 Tf 1 0 0 1 30 730 Tm (Export heading) Tj " +
    "/F1 12 Tf 1 0 0 1 30 680 Tm (Normal ) Tj /F2 12 Tf (emphasis) Tj " +
    "/F1 12 Tf 1 0 0 1 30 640 Tm (1. First list item) Tj ET"));
var styledXml = Validate(PdfDocxConverter.Convert(styled).Bytes);
Check(styledXml.Descendants(w + "pStyle").Any(), "Heading style inferred");
var emphasis = styledXml.Descendants(w + "r").Single(run => run.Element(w + "t")?.Value == "emphasis");
Check(emphasis.Descendants(w + "b").Any() && emphasis.Descendants(w + "i").Any(), "Mixed run bold and italic preserved");
Check(string.Concat(styledXml.Descendants(w + "t").Select(text => text.Value)).Contains("1. First list item"), "List marker retained once");
var blank = PdfReader.Read(PdfFixture.Create(null, new FixturePage()));
Validate(PdfDocxConverter.Convert(blank).Bytes);
Check(typeof(PdfDocxConverter).Assembly.GetName().Name == "PdfObjectKit.Extensions.Docx", "Separate optional extension");
Console.WriteLine("DOCX export tests passed: package validation, text, tables, images, page options, cancellation and source preservation.");
