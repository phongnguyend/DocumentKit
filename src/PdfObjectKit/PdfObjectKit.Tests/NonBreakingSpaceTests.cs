using System.Text;
using PdfObjectKit;
using PdfObjectKit.Converters;
using PdfObjectKit.Extractors;
using PdfObjectKit.Fonts;
using PdfObjectKit.Graphs;
using PdfObjectKit.Rendering;
using PdfObjectKit.Readers;
using static PdfObjectKit.Primitives.PdfPrimitives;

internal static class NonBreakingSpaceTests
{
    public static void Run()
    {
        foreach (var name in new[] { "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique" })
        {
            var document = PdfReader.Read(PdfFixture.Create(null, new FixturePage()));
            var graph = new PdfGraph(document);
            var page = graph.Pages()[0];
            var source = graph.Resolve(graph.Get(page.Resources, "Font")!.Entries!["F1"]);
            source.Entries!["BaseFont"] = Name(name);
            var font = new PdfNativeFont(graph, source);
            var glyph = font.Decode([160]).Single();
            if (glyph.Text != "\u00a0" || glyph.Width != font.Decode([32]).Single().Width)
            {
                throw new Exception("WinAnsi non-breaking space must preserve Unicode and use the space advance.");
            }
            page.Streams[0].Value.Base64 = Convert.ToBase64String(Encoding.ASCII.GetBytes("BT /F1 12 Tf 20 100 Td <41A042> Tj ET"));
            var before = PdfConverter.ToJson(document);
            var lines = PdfParagraphExtractor.ExtractSelectionLines(document, 1);
            if (lines.Count != 1 || lines[0].Text != "A\u00a0B" || lines[0].Right <= lines[0].Left
                || PdfTextExtractor.Extract(document)[0].Warning is not null
                || PdfTextSearcher.Search(document, "B").Matches.Count != 1
                || !PdfMarkdownConverter.Convert(document).Markdown.Contains('B')
                || !PdfCanvasExporter.Export(document, 1).Commands.Any(command => command.Kind == "text"))
            {
                throw new Exception("Non-breaking spaces must not suppress selection, Markdown, search or rendering.");
            }
            if (before != PdfConverter.ToJson(document))
            {
                throw new Exception("Selection must not alter the PDF.");
            }
        }
        Console.WriteLine("Standard Helvetica non-breaking space checks passed.");
    }
}
