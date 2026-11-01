using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PdfObjectKit.Converters;
using PdfObjectKit.Editors;
using PdfObjectKit.Graphs;
using PdfObjectKit.Readers;
using static PdfObjectKit.Primitives.PdfPrimitives;

internal static class PptxTests
{
    public static void Run()
    {
        Check(typeof(PdfPptxConverter).Assembly.GetName().Name == "PdfObjectKit.Extensions.Pptx", "PPTX converter is owned by the optional extension");
        Check(typeof(PdfObjectKit.Options.PdfPptxOptions).Assembly == typeof(PdfPptxConverter).Assembly
            && typeof(PdfObjectKit.Models.PdfPptxResult).Assembly == typeof(PdfPptxConverter).Assembly, "PPTX public types live together in the extension");
        Check(typeof(PdfReader).Assembly.GetReferencedAssemblies().All(reference => reference.Name != "StbImageSharp"
            && reference.Name != "PdfObjectKit.Extensions.Pptx"), "Core assembly does not depend on PPTX or its image decoder");
        var document = PdfReader.Read(PdfFixture.Create(null, new("Editable title", Width: 400, Height: 240, X: 30, Y: 180, FontSize: 24), new("Page two", Width: 240, Height: 400)));
        var graph = new PdfGraph(document);
        var page = graph.Pages()[0];
        var content = page.Streams[0].Value;
        var original = Encoding.ASCII.GetString(Convert.FromBase64String(content.Base64!));
        content.Base64 = Convert.ToBase64String(Encoding.ASCII.GetBytes("0 0 1 rg 20 20 100 40 re f\n1 0 0 rg\n" + original + "\nBT /F1 12 Tf 3 Tr 1 0 0 1 20 60 Tm (Hidden OCR) Tj ET\n"));
        PdfImageEditor.InsertImage(document, 1, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="), 200, 50, 50, 60);
        var before = PdfConverter.ToJson(document);
        var result = PdfPptxConverter.Convert(document);
        Check(result.SlideCount == 2 && result.TextBoxCount == 2 && result.ImageCount == 1 && result.ShapeCount == 1, "Editable object counts");
        Check(PdfConverter.ToJson(document) == before, "Export does not mutate PDF");
        Validate(result.Bytes);
        using var archive = new ZipArchive(new MemoryStream(result.Bytes));
        XDocument Xml(string name)
        {
            using var stream = archive.GetEntry(name)!.Open();
            return XDocument.Load(stream);
        }
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var slide = Xml("ppt/slides/slide1.xml");
        Check(slide.Descendants(a + "t").Single().Value == "Editable title", "Visible Unicode text stored as text, hidden text excluded");
        Check(slide.Descendants(a + "rPr").Single().Element(a + "solidFill")!.Element(a + "srgbClr")!.Attribute("val")!.Value == "FF0000", "Text fill color");
        Check(slide.Descendants(p + "spTree").Single().Elements().Skip(2).Select(e => e.Name.LocalName).SequenceEqual(new[] { "sp", "sp", "pic" }), "Drawing order retained");
        Check(slide.Descendants(p + "pic").Count() == 1 && archive.GetEntry("ppt/media/image1.png") is not null, "Editable picture and media part");
        var size = Xml("ppt/presentation.xml").Descendants(p + "sldSz").Single();
        Check((long)size.Attribute("cx")! == 400 * 12700 && (long)size.Attribute("cy")! == 240 * 12700, "Slide dimensions follow first page");
        var textOnly = PdfPptxConverter.Convert(document, new() { IncludeImages = false, IncludePaths = false, PageNumbers = [2, 1], SlideWidth = 960, SlideHeight = 540 });
        Check(textOnly.ImageCount == 0 && textOnly.ShapeCount == 0 && textOnly.TextBoxCount == 2, "Content options and page order");
        Validate(textOnly.Bytes);
        PdfPageEditor.RotatePage(document, 1, 90);
        var rotated = PdfPptxConverter.Convert(document, new() { PageNumbers = [1] });
        Validate(rotated.Bytes);
        using (var zip = new ZipArchive(new MemoryStream(rotated.Bytes)))
        using (var stream = zip.GetEntry("ppt/slides/slide1.xml")!.Open())
        {
            var xml = XDocument.Load(stream);
            Check(xml.Descendants(a + "xfrm").Any(x => (int?)x.Attribute("rot") == 5400000), "Rotated page transforms editable objects");
        }
        page = new PdfGraph(document).Pages()[0];
        page.Dictionary.Entries!["CropBox"] = Array(Num(10), Num(10), Num(390), Num(230));
        page.Dictionary.Entries["UserUnit"] = Num(2);
        Validate(PdfPptxConverter.Convert(document).Bytes);
        var clipped = PdfReader.Read(PdfFixture.Create(null, new FixturePage("Clipped")));
        new PdfGraph(clipped).Pages()[0].Streams[0].Value.Base64 = Convert.ToBase64String(Encoding.ASCII.GetBytes("0 0 100 100 re W n\n" + original));
        Check(PdfPptxConverter.Convert(clipped).Warnings.Any(w => w.Contains("Clipping")), "Clipping limitation reported");
        Reject(() => PdfPptxConverter.Convert(document, new() { PageNumbers = [0] }));
        Reject(() => PdfPptxConverter.Convert(document, new() { PageNumbers = [1, 1] }));
        Reject(() => PdfPptxConverter.Convert(document, new() { SlideWidth = 100 }));
        Reject(() => PdfPptxConverter.Convert(document, new() { SlideWidth = double.NaN, SlideHeight = 100 }));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        try
        {
            PdfPptxConverter.Convert(document, cancellationToken: cancelled.Token);
            throw new Exception("Cancelled export succeeded");
        }
        catch (OperationCanceledException)
        {
        }
        if (Environment.GetEnvironmentVariable("PDFOBJECTKIT_PPTX_FIXTURE") is { Length: > 0 } destination)
        {
            File.WriteAllBytes(destination, result.Bytes);
        }
        CheckLineGrouping();
        PptxGraphicsTests.Run();
        Console.WriteLine("Editable PPTX schema, text, colors, shapes, images, geometry, options, cancellation and immutability checks passed.");
    }

    private static void CheckLineGrouping()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        (PdfObjectKit.Models.PdfObjectDocument Document, PdfObjectKit.Models.PdfPptxResult Result, XDocument Slide) Export(string content, bool rotate = false)
        {
            var document = PdfReader.Read(PdfFixture.Create(null, new FixturePage("Line", Width: 400, Height: 240)));
            var graph = new PdfGraph(document);
            var page = graph.Pages()[0];
            graph.Get(page.Resources, "Font")!.Entries!["F2"] = Dict(("Type", Name("Font")), ("Subtype", Name("Type1")), ("BaseFont", Name("Helvetica-Bold")));
            page.Streams[0].Value.Base64 = Convert.ToBase64String(Encoding.ASCII.GetBytes(content));
            if (rotate)
            {
                PdfPageEditor.RotatePage(document, 1, 90);
            }
            var result = PdfPptxConverter.Convert(document);
            Validate(result.Bytes);
            using var zip = new ZipArchive(new MemoryStream(result.Bytes));
            using var stream = zip.GetEntry("ppt/slides/slide1.xml")!.Open();
            return (document, result, XDocument.Load(stream));
        }
        const string mixed = "BT /F1 12 Tf 1 0 0 rg 1 0 0 1 20 200 Tm [(Hello) -300] TJ /F2 16 Tf 0 0 1 rg (World) Tj ET";
        var line = Export(mixed);
        Check(line.Result.TextBoxCount == 1, "Mixed-format runs become one text box");
        var body = line.Slide.Descendants(p + "txBody").Single();
        var text = string.Concat(body.Descendants(a + "t").Select(t => t.Value));
        Check(text == "Hello World", "Geometric word spacing retained in grouped line");
        Check(PdfObjectKit.Extractors.PdfParagraphExtractor.ExtractSelectionLines(line.Document, 1).Single().Text == text, "Export matches UI line selection text");
        var runs = body.Descendants(a + "rPr").ToArray();
        Check(runs.Length == 2 && (int)runs[0].Attribute("sz")! == 1200 && (int)runs[1].Attribute("sz")! == 1600 && (int)runs[1].Attribute("b")! == 1, "Grouped runs retain font sizes and bold");
        Check(runs[0].Descendants(a + "srgbClr").Single().Attribute("val")!.Value == "FF0000" && runs[1].Descendants(a + "srgbClr").Single().Attribute("val")!.Value == "0000FF", "Grouped runs retain individual colors");
        var rotated = Export(mixed, rotate: true);
        Check(rotated.Result.TextBoxCount == 1 && (int)rotated.Slide.Descendants(p + "spPr").Single().Element(a + "xfrm")!.Attribute("rot")! == 5400000, "Rotated pages retain line grouping and rotation");
        Check(Export("BT /F1 12 Tf 20 200 Td (Hel) Tj (lo) Tj ET").Result.TextBoxCount == 1, "Adjacent fragments merge without requiring whitespace");
        Check(Export("BT /F1 12 Tf 20 200 Td [(Hel) 6 (lo)] TJ ET").Result.TextBoxCount == 1, "Small negative kerning does not split a line");
        using (var package = new PptxPackage())
        {
            var writer = new PptxSlideWriter(package, new(1, 0, 0, -1, 0, 240), 400, 240, new(), _ => { });
            writer.Write([new PdfObjectKit.Rendering.PdfCanvasCommand("text")
            {
                Numbers = [1, 0, 0, 1, 20, 200, 12, 1, 0],
                FontName = "Wingdings",
                Glyphs = [new("\u25aa", 6, 6)]
            }], default);
            var bytes = package.Finish(400, 240);
            Validate(bytes);
            using var zip = new ZipArchive(new MemoryStream(bytes));
            using var stream = zip.GetEntry("ppt/slides/slide1.xml")!.Open();
            var xml = XDocument.Load(stream);
            Check(xml.Descendants(a + "t").Single().Value == "\u25aa"
                && xml.Descendants(a + "latin").Single().Attribute("typeface")!.Value == "Arial", "Decoded Unicode bullet does not reuse symbol-font encoding");
        }
        const string scoped = "q 0 0 400 240 re W n BT /F1 12 Tf 20 200 Td (Hello) Tj ET Q q 0 0 400 240 re W n 1 0 0 rg BT /F2 12 Tf 50 200 Td (World) Tj ET Q";
        var scopedLine = Export(scoped);
        Check(scopedLine.Result.TextBoxCount == 1, "State saves and repeated page clips do not split a line");
        Check(string.Concat(scopedLine.Slide.Descendants(a + "t").Select(t => t.Value)) == "Hello World", "State-separated words retain text order and spacing");
        Check(Export(scoped.Replace("0 0 400 240 re", "0 150 90 60 re")).Result.TextBoxCount == 2, "Real clipping boundaries still separate text");
        var columns = Export("BT /F1 12 Tf 1 0 0 1 20 200 Tm (Left) Tj 1 0 0 1 250 200 Tm (Right) Tj 1 0 0 1 20 160 Tm (Next line) Tj ET");
        Check(columns.Result.TextBoxCount == 3, "Separate columns and baselines remain independent");
        var layered = Export("BT /F1 12 Tf 20 200 Td (Hello) Tj ET 40 190 20 20 re f BT /F1 12 Tf 50 200 Td (World) Tj ET");
        Check(layered.Result.TextBoxCount == 2 && layered.Result.ShapeCount == 1, "Painted graphics prevent text merging across layers");
        Check(layered.Slide.Descendants(p + "spTree").Single().Elements(p + "sp").Select(shape => shape.Element(p + "txBody") is not null).SequenceEqual(new[] { true, false, true }), "Painted shape stays between text boxes");
        Check(Export("BT /F1 12 Tf 20 200 Td (Hello) Tj 1 0 0 1 20 200 Tm (Shadow) Tj ET").Result.TextBoxCount == 2, "Overlapping text stays independent");
    }

    private static void Validate(byte[] bytes)
    {
        using var presentation = PresentationDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator().Validate(presentation).Take(10).ToArray();
        if (errors.Length > 0)
        {
            throw new Exception(string.Join("\n", errors.Select(error => $"{error.Path?.XPath}: {error.Description}")));
        }
    }

    private static void Check(bool condition, string message)
    {
        if (!condition)
        {
            throw new Exception(message);
        }
    }

    private static void Reject(Action action)
    {
        try
        {
            action();
        }
        catch (ArgumentException)
        {
            return;
        }
        throw new Exception("Invalid PPTX options accepted.");
    }
}
