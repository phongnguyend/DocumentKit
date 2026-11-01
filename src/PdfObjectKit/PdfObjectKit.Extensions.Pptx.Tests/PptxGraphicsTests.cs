using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using PdfObjectKit.Converters;
using PdfObjectKit.Rendering;
using StbImageSharp;

internal static class PptxGraphicsTests
{
    internal static void Run()
    {
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var warnings = new List<string>();
        using var package = new PptxPackage();
        var writer = new PptxSlideWriter(package, new(1, 0, 0, -1, 0, 240), 400, 240, new(), warnings.Add);
        var gradient = new PdfCanvasCommand("fillGradient") { Numbers = [0, 0, 200, 0, 1, 0, 0, 1, 0, 0, 0, 1, 0, 0, 1, 0, 0, 1] };
        writer.Write([
            gradient,
            new("re") { Numbers = [50, 50, 100, 50] }, new("f"),
            new("q"), new("g") { Numbers = [0] }, new("re") { Numbers = [50, 110, 100, 20] }, new("f"), new("Q"),
            new("text") { Numbers = [1, 0, 0, 1, 50, 180, 12, 1, 0], FontName = "Arial", Glyphs = [new("Text", 24, 24)] },
            new("rg") { Numbers = [0, 1, 0] }, new("re") { Numbers = [50, 10, 100, 20] }, new("f")
        ], default);
        var bytes = package.Finish(400, 240);
        using (var document = PresentationDocument.Open(new MemoryStream(bytes), false))
        {
            Check(!new OpenXmlValidator().Validate(document).Any(), "Gradient package validates");
        }
        using var zip = new ZipArchive(new MemoryStream(bytes));
        using var stream = zip.GetEntry("ppt/slides/slide1.xml")!.Open();
        var xml = XDocument.Load(stream);
        var shapes = xml.Descendants(p + "sp").ToArray();
        var colors = shapes[0].Descendants(a + "gs").Select(stop => stop.Element(a + "srgbClr")!.Attribute("val")!.Value).ToArray();
        Check(colors.SequenceEqual(new[] { "BF0040", "4000BF" }), "Gradient stops remapped to object position");
        Check(shapes[1].Descendants(a + "gradFill").Count() == 0 && shapes[1].Descendants(a + "srgbClr").Single().Attribute("val")!.Value == "000000", "Solid fill clears gradient inside saved state");
        Check(shapes[2].Descendants(p + "txBody").Any() && shapes[2].Descendants(a + "gradFill").Any(), "Restored gradient remains editable text");
        Check(shapes[3].Descendants(a + "gradFill").Count() == 0, "Subsequent solid color clears gradient");
        Check(warnings.Count == 0, "Supported gradients do not produce approximation warnings");

        using (var outlinedPackage = new PptxPackage())
        {
            var outlinedWriter = new PptxSlideWriter(outlinedPackage, new(1, 0, 0, -1, 0, 240), 400, 240, new(), _ => { });
            var commands = new List<PdfCanvasCommand> { gradient };
            void Word(double x, double width, string text)
            {
                commands.AddRange([
                    new("alpha") { Numbers = [1, -1] },
                    new("m") { Numbers = [x, 200] },
                    new("c") { Numbers = [x, 209, x + width, 209, x + width, 200] }, new("h"), new("f"),
                    new("alpha") { Numbers = [0, -1] },
                    new("text") { Numbers = [1, 0, 0, 1, x, 200, 12, 1, 0], FontName = "Arial-BoldMT", Glyphs = [new(text, width, width)] }
                ]);
            }
            Word(20, 30, "Thank ");
            Word(52, 20, "you");
            commands.Add(new("text") { Numbers = [1, 0, 0, 1, 200, 50, 12, 1, 0], FontName = "Arial", Glyphs = [new("Hidden unrelated text", 100, 100)] });
            outlinedWriter.Write(commands, default);
            var outlinedBytes = outlinedPackage.Finish(400, 240);
            using var document = PresentationDocument.Open(new MemoryStream(outlinedBytes), false);
            Check(!new OpenXmlValidator().Validate(document).Any(), "Recovered outline text validates");
            using var outlinedZip = new ZipArchive(new MemoryStream(outlinedBytes));
            using var outlinedStream = outlinedZip.GetEntry("ppt/slides/slide1.xml")!.Open();
            var outlinedXml = XDocument.Load(outlinedStream);
            Check(outlinedWriter.TextCount == 1 && outlinedWriter.ShapeCount == 0, "Adjacent annotated glyph outlines become one editable text box");
            Check(string.Concat(outlinedXml.Descendants(a + "t").Select(t => t.Value)) == "Thank you", "Recovery keeps source words and excludes unrelated invisible text");
            Check(outlinedXml.Descendants(a + "gradFill").Count() == 2, "Recovered formatted runs retain gradient");
        }

        var jpeg = Convert.FromBase64String("/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAACAAIDAREAAhEBAxEB/8QAFAABAAAAAAAAAAAAAAAAAAAACv/EABQQAQAAAAAAAAAAAAAAAAAAAAD/xAAUAQEAAAAAAAAAAAAAAAAAAAAH/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAwDAQACEQMRAD8AP+FzA//Z");
        var command = new PdfCanvasCommand("jpeg") { Width = 2, Height = 2, AlphaData = Convert.ToBase64String([0, 128, 255, 64]), Matte = [0, 0, 0] };
        var original = ImageResult.FromMemory(jpeg, ColorComponents.RedGreenBlueAlpha);
        var png = PptxMaskedImage.Decode(command, jpeg);
        var decoded = ImageResult.FromMemory(png, ColorComponents.RedGreenBlueAlpha);
        Check(decoded.Data.Where((_, index) => index % 4 == 3).SequenceEqual(new byte[] { 0, 128, 255, 64 }), "PNG retains transparent, partial and opaque mask samples");
        Check(decoded.Data[0] == 0 && decoded.Data[4] == Math.Clamp(Math.Round(original.Data[4] * 255d / 128), 0, 255), "Black matte removed without retaining hidden RGB");
        Check(decoded.Data[8] == original.Data[8], "Opaque pixels preserved");
        try
        {
            PptxMaskedImage.Decode(new("jpeg") { Width = 2, Height = 2, AlphaData = "AA==" }, jpeg);
            throw new Exception("Invalid mask accepted");
        }
        catch (InvalidDataException)
        {
        }
    }

    private static void Check(bool condition, string message)
    {
        if (!condition)
        {
            throw new Exception(message);
        }
    }
}
