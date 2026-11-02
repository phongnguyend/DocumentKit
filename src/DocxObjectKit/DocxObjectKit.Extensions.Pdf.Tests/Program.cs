using DocxObjectKit.Readers;
using DocxObjectKit.Extensions.Pdf;
using DocxObjectKit.Application;
using DocxObjectKit.Editors;
using static DocxObjectKit.Models.DocxNamespaces;
using System.Xml.Linq;

var count = 0;
void Check(bool value, string message)
{
    if (!value)
    {
        throw new Exception(message);
    }
    count++;
}
EmfDiagramTests.Run(Check);
var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures/sample.docx"));
var source = DocxReader.Read(bytes);
LayoutFeatureTests.Run(source, Check);
PageFieldTests.Run(source, Check);
LineShapeTests.Run(source, Check);
var layout = DocumentApplication.Inspect(bytes);
Check(layout.Pages.Count == 2, "Explicit page break not honored.");
Check(layout.Pages[0].Runs.Any(item => item.Text == "Red" && item.Bold && item.Color == "CC0033"), "Text styling missing.");
Check(layout.Pages[0].Rules.Count == 4, "Table borders missing.");
Check(layout.Pages[0].Images.Count == 2 && layout.Media.Count == 1, "Shared image resources not extracted.");
Check(layout.Pages[1].TextObjects.Any(item => item.Text == "Second page content."), "Page two text missing.");
Check(layout.Pages.SelectMany(page => page.Runs).All(run => run.X >= 0 && run.Y >= 0 && run.Baseline > run.Y), "Invalid text layout.");
var transformed = DocumentApplication.Edit(bytes, new("resize-image", ObjectId: "i0", X: 150, Y: 220, Width: 180, Height: 60));
var moved = DocumentApplication.Inspect(transformed).Pages[0].Images[0];
Check(moved.X == 150 && moved.Y == 220 && moved.Width == 180 && moved.Height == 60, "Moved image layout does not follow XML.");
var text = DocumentApplication.Edit(bytes, new("transform-text", ObjectId: "p0", X: 100, Y: 120, Width: 250, Height: 40));
Check(Math.Abs(DocumentApplication.Inspect(text).Pages[0].TextObjects.First().Bounds.X - 100) < .1, "Text frame positioning missing.");
var pdf = DocxPdfConverter.Convert(source);
Check(System.Text.Encoding.ASCII.GetString(pdf.Bytes, 0, 5) == "%PDF-" && pdf.PageCount == 2, "PDF output missing.");
Check(PdfObjectKit.Extractors.PdfTextSearcher.Search(PdfObjectKit.Readers.PdfReader.Read(pdf.Bytes), "Document toolkit sample").Matches.Count > 0, "Exported text is not searchable.");
Check(DocxPdfConverter.Convert(source, new([2])).PageCount == 1, "Page selection ignored.");
var xml = source.ReadXml("word/document.xml");
var body = xml.Root!.Element(W + "body")!;
body.Elements().Where(item => item.Name != W + "sectPr").Remove();
for (var i = 0; i < 150; i++)
{
    body.AddFirst(new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "Flowing paragraph " + i))));
}
using var flowEngine = new DocxLayout(source.WithXml("word/document.xml", xml));
var flow = flowEngine.Build();
Check(flow.Pages.Count > 2 && flow.Pages.Sum(page => page.TextObjects.Count) == 150, "Automatic pagination lost paragraphs.");
var ocr = DocumentApplication.Edit(bytes, new("apply-image-ocr", ObjectId: "i0", OcrLines: [new("Recognized words", 0, 0, .9, .4, .3)]));
Check(DocxReader.Read(ocr).ReadXml("word/document.xml").Descendants(W + "t").Any(item => item.Value == "Recognized words"), "OCR text not applied.");
// Synthetic regressions: these examples contain no customer document content.
XElement Paragraph(string value) => new(W + "p", new XElement(W + "r", new XElement(W + "t", value)));
DocxObjectKit.Models.DocxObjectDocument Synthetic(params XElement[] blocks)
{
    var package = source.ReadXml("word/document.xml");
    var container = package.Root!.Element(W + "body")!;
    container.ReplaceNodes(blocks, new XElement(W + "sectPr", new XElement(W + "pgSz", new XAttribute(W + "w", 12240), new XAttribute(W + "h", 15840)), new XElement(W + "pgMar", new XAttribute(W + "top", 1440), new XAttribute(W + "bottom", 1440), new XAttribute(W + "left", 1440), new XAttribute(W + "right", 1440))));
    return source.WithXml("word/document.xml", package);
}
var table = new XElement(W + "tbl", new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 6000)), new XElement(W + "gridCol", new XAttribute(W + "w", 3000))));
for (var row = 0; row < 50; row++)
{
    table.Add(new XElement(W + "tr", new XElement(W + "tc", Paragraph("Synthetic row " + row)), new XElement(W + "tc", Paragraph("Cell " + row))));
}
using var tableEngine = new DocxLayout(Synthetic(table));
var tableLayout = tableEngine.Build();
Check(tableLayout.Pages.Count is > 1 and < 6, "Table pagination must advance once per row, not per cell.");
Check(tableLayout.Pages.Sum(item => item.TextObjects.Count) == 100, "Table pagination lost cell text.");
Check(tableLayout.Pages.All(item => item.Rules.Chunk(2).All(pair => pair.Length == 2 && Math.Abs(pair[0].Y - pair[1].Y) < .01 && pair[0].Width > pair[1].Width)), "Table grid widths or row alignment were lost.");
var spanning = new XElement(W + "tbl", new XElement(W + "tblGrid", Enumerable.Range(0, 3).Select(_ => new XElement(W + "gridCol", new XAttribute(W + "w", 2000)))), new XElement(W + "tr", new XElement(W + "tc", new XElement(W + "tcPr", new XElement(W + "gridSpan", new XAttribute(W + "val", 2))), Paragraph("Spanning cell")), new XElement(W + "tc", Paragraph("Final cell"))));
using var spanEngine = new DocxLayout(Synthetic(spanning));
var spanLayout = spanEngine.Build();
Check(Math.Abs(spanLayout.Pages[0].Rules[0].Width - 2 * spanLayout.Pages[0].Rules[1].Width) < .01, "Column span was not applied.");
using var controlEngine = new DocxLayout(Synthetic(new XElement(W + "sdt", new XElement(W + "sdtContent", Paragraph("Synthetic contents entry")))));
Check(controlEngine.Build().Pages[0].Runs.Any(run => run.Text == "Synthetic"), "Content-control paragraphs were skipped.");
var numbered = Paragraph("List entry");
numbered.AddFirst(new XElement(W + "pPr", new XElement(W + "numPr", new XElement(W + "numId", new XAttribute(W + "val", 1)))));
var numberPackage = Synthetic(numbered, new XElement(numbered)).WithXml("word/numbering.xml", new XDocument(new XElement(W + "numbering", new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", 0), new XElement(W + "lvl", new XAttribute(W + "ilvl", 0), new XElement(W + "start", new XAttribute(W + "val", 1)), new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")), new XElement(W + "lvlText", new XAttribute(W + "val", "%1.")))), new XElement(W + "num", new XAttribute(W + "numId", 1), new XElement(W + "abstractNumId", new XAttribute(W + "val", 0))))));
using var numberEngine = new DocxLayout(numberPackage);
var numberedLayout = numberEngine.Build();
Check(numberedLayout.Pages[0].Runs.Any(run => run.Text == "1.") && numberedLayout.Pages[0].Runs.Any(run => run.Text == "2."), "Numbered lists must increment their counters.");
var sectionParagraph = Paragraph("First section");
sectionParagraph.AddFirst(new XElement(W + "pPr", new XElement(W + "sectPr", new XElement(W + "pgMar", new XAttribute(W + "left", 720)))));
using var sectionEngine = new DocxLayout(Synthetic(sectionParagraph, Paragraph("Second section")));
var sectionLayout = sectionEngine.Build();
Check(sectionLayout.Pages.Count == 2 && sectionLayout.Pages[0].Runs[0].X == 36 && sectionLayout.Pages[1].Runs[0].X == 72, "Section breaks and margins were not applied.");
XNamespace wordShape = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
var anchor = new XElement(WP + "anchor", new XElement(WP + "positionV", new XAttribute("relativeFrom", "paragraph"), new XElement(WP + "posOffset", 254000)), new XElement(WP + "extent", new XAttribute("cx", 2540000), new XAttribute("cy", 762000)), new XElement(WP + "wrapSquare"), new XElement(wordShape + "wsp", new XElement(wordShape + "txbx", new XElement(W + "txbxContent", Paragraph("Floating explanation")))));
var headingWithBox = new XElement(W + "p", new XElement(W + "r", new XElement(W + "drawing", anchor)), new XElement(W + "r", new XElement(W + "t", "Heading beside anchor")));
using var boxEngine = new DocxLayout(Synthetic(headingWithBox, Paragraph("Following paragraph")));
var boxes = boxEngine.Build().Pages[0];
Check(boxes.Runs.First(run => run.Text == "Heading").Y < boxes.Runs.First(run => run.Text == "Floating").Y, "Floating drawing must not split or displace its own heading text.");
Check(boxes.Runs.First(run => run.Text == "Following").Y >= boxes.Runs.First(run => run.Text == "Floating").Y + 40, "Wrapping text box does not reserve flow space.");
Check(boxes.TextObjects.First(item => item.Id == "p0").Text == "Heading beside anchor", "Parent paragraph selection includes nested textbox content.");
var headerPackage = Synthetic(Paragraph("Body content"));
var headerXml = headerPackage.ReadXml("word/document.xml");
headerXml.Descendants(W + "sectPr").Single().Add(new XElement(W + "headerReference", new XAttribute(W + "type", "default"), new XAttribute(R + "id", "syntheticHeader")));
var headerRelations = headerPackage.ReadXml("word/_rels/document.xml.rels");
headerRelations.Root!.Add(new XElement(REL + "Relationship", new XAttribute("Id", "syntheticHeader"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"), new XAttribute("Target", "synthetic-header.xml")));
headerPackage = headerPackage.WithXml("word/document.xml", headerXml).WithXml("word/_rels/document.xml.rels", headerRelations).WithXml("word/synthetic-header.xml", new XDocument(new XElement(W + "hdr", Paragraph("Synthetic header"))));
using var headerEngine = new DocxLayout(headerPackage);
var headerPage = headerEngine.Build().Pages[0];
Check(headerPage.Runs.Any(run => run.Text == "Synthetic") && headerPage.TextObjects.Count == 1, "Header text should render without creating editable body selections.");
var imagePart = DocxEditor.ImagePart(source, DocxEditor.Pictures(source.ReadXml("word/document.xml")).First());
foreach (var clip in new[] { false, true })
{
    var emfPackage = source.WithPart(imagePart, EmfFixtures.Create(clip, curve: true));
    using var emfEngine = new DocxLayout(emfPackage);
    var emfLayout = emfEngine.Build();
    Check(emfLayout.Pages[0].Images.Count == 2 && emfLayout.Media[imagePart].StartsWith("data:image/png;base64,"), "EMF images should be decoded for both preview and PDF.");
    var uri = emfLayout.Media[imagePart];
    using var raster = SkiaSharp.SKBitmap.Decode(Convert.FromBase64String(uri[(uri.IndexOf(',') + 1)..]));
    var middle = raster.GetPixel(raster.Width / 3, raster.Height / 2);
    Check(middle.Red == 0xdd && middle.Green == 0x22 && middle.Blue == 0x33 && middle.Alpha == 255, "EMF brush color or logical coordinate mapping is incorrect.");
    var rightPixel = raster.GetPixel(raster.Width * 3 / 4, raster.Height / 2);
    Check(clip ? rightPixel.Alpha == 0 : rightPixel.Alpha == 255, "EMF clipping was not applied.");
    var emfPdf = PdfObjectKit.Readers.PdfReader.Read(DocxPdfConverter.Convert(emfPackage).Bytes);
    Check(PdfObjectKit.Extractors.PdfImageExtractor.Extract(emfPdf).Sum(page => page.Images.Count) == 2, "PDF export must embed both decoded EMF image placements.");
}
var invalidEmf = EmfFixtures.Create();
System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(invalidEmf.AsSpan(92), int.MaxValue);
using var invalidEmfEngine = new DocxLayout(source.WithPart(imagePart, invalidEmf));
Check(invalidEmfEngine.Build().Warnings.Any(item => item.Contains("EMF record size")), "Malformed EMF must be rejected safely.");
using var unsupportedEmfEngine = new DocxLayout(source.WithPart(imagePart, EmfFixtures.Create(unsupported: true)));
var unsupportedEmf = unsupportedEmfEngine.Build();
Check(OperatingSystem.IsWindows()
    ? unsupportedEmf.Pages.Sum(item => item.Images.Count) == 2
    : unsupportedEmf.Pages.Sum(item => item.Images.Count) == 0 && unsupportedEmf.Warnings.Any(item => item.Contains("Unsupported EMF")), "Windows fallback must render rectangle records; other platforms must report unsupported EMF.");
var invalidPoints = EmfFixtures.Create();
for (var recordOffset = 0; recordOffset < invalidPoints.Length;)
{
    var recordType = System.Buffers.Binary.BinaryPrimitives.ReadInt32LittleEndian(invalidPoints.AsSpan(recordOffset));
    if (recordType == 89)
    {
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(invalidPoints.AsSpan(recordOffset + 24), int.MaxValue);
        break;
    }
    recordOffset += System.Buffers.Binary.BinaryPrimitives.ReadInt32LittleEndian(invalidPoints.AsSpan(recordOffset + 4));
}
using var invalidPointsEngine = new DocxLayout(source.WithPart(imagePart, invalidPoints));
Check(invalidPointsEngine.Build().Warnings.Any(item => item.Contains("EMF point array")), "Oversized EMF point arrays must not allocate or read outside the record.");
DocxObjectKit.Models.DocxObjectDocument ColumnDocument(XElement[] blocks, bool unequal = false)
{
    var package = Synthetic(blocks);
    var content = package.ReadXml("word/document.xml");
    var section = content.Descendants(W + "sectPr").Last();
    section.Element(W + "pgSz")!.SetAttributeValue(W + "h", 3600);
    foreach (var margin in new[] { "top", "bottom", "left", "right" })
    {
        section.Element(W + "pgMar")!.SetAttributeValue(W + margin, 360);
    }
    var columns = new XElement(W + "cols", new XAttribute(W + "num", 2), new XAttribute(W + "space", 360));
    if (unequal)
    {
        columns.SetAttributeValue(W + "equalWidth", 0);
        columns.Add(new XElement(W + "col", new XAttribute(W + "w", 1800), new XAttribute(W + "space", 360)), new XElement(W + "col", new XAttribute(W + "w", 9360)));
    }
    section.Add(columns);
    return package.WithXml("word/document.xml", content);
}
using var columnsEngine = new DocxLayout(ColumnDocument(Enumerable.Range(0, 40).Select(index => Paragraph("Column paragraph " + index)).ToArray()));
var columnsLayout = columnsEngine.Build();
Check(columnsLayout.Pages.Count is > 1 and < 5 && columnsLayout.Pages.Sum(item => item.TextObjects.Count) == 40, "Column pagination must retain every paragraph.");
Check(columnsLayout.Pages[0].Runs.Any(run => run.X >= 315) && columnsLayout.Pages[0].Runs.Any(run => run.X == 18), "Body text must flow through both columns before starting another page.");
var breaks = new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "First"), new XElement(W + "br", new XAttribute(W + "type", "column")), new XElement(W + "t", "Second"), new XElement(W + "br", new XAttribute(W + "type", "page")), new XElement(W + "t", "Third")));
using var breaksEngine = new DocxLayout(ColumnDocument([breaks], true));
var broken = breaksEngine.Build();
Check(broken.Pages.Count == 2 && broken.Pages[0].Runs.Single(run => run.Text == "Second").X == 126 && broken.Pages[1].Runs.Single().X == 18, "Explicit column/page breaks did not reset flow correctly.");
using var columnTableEngine = new DocxLayout(ColumnDocument([new XElement(table)]));
var columnTable = columnTableEngine.Build();
Check(columnTable.Pages[0].Rules.Any(rule => rule.X >= 315) && columnTable.Pages.SelectMany(item => item.Rules).All(rule => rule.X + rule.Width <= 594.01), "Table rows must remain inside their destination columns.");
var anchoredXml = source.ReadXml("word/document.xml");
var anchored = DocxEditor.Pictures(anchoredXml).First();
anchored.Name = WP + "anchor";
anchored.Add(new XElement(WP + "positionH", new XAttribute("relativeFrom", "margin"), new XElement(WP + "posOffset", 19 * 12700)), new XElement(WP + "positionV", new XAttribute("relativeFrom", "margin"), new XElement(WP + "posOffset", 11 * 12700)));
using var marginEngine = new DocxLayout(source.WithXml("word/document.xml", anchoredXml));
var marginImage = marginEngine.Build().Pages[0].Images.First();
Check(marginImage.X == 91 && marginImage.Y == 83, "Margin-relative images must use page margins rather than the current paragraph position.");
anchored.Element(WP + "positionH")!.ReplaceNodes(new XElement(WP + "align", "center"));
using var alignedEngine = new DocxLayout(source.WithXml("word/document.xml", anchoredXml));
var alignedLayout = alignedEngine.Build();
var alignedImage = alignedLayout.Pages[0].Images.First();
Check(Math.Abs(alignedImage.X - (alignedLayout.Width - alignedImage.Width) / 2) < .01, "Centered drawing alignment was ignored.");
using var anchoredFlowEngine = new DocxLayout(ColumnDocument(Enumerable.Range(0, 5).Select(index => Paragraph("Lead " + index)).Append(new XElement(headingWithBox)).ToArray()));
var anchoredFlow = anchoredFlowEngine.Build();
var anchorHeading = anchoredFlow.Pages.SelectMany(item => item.Runs).First(run => run.Text == "Heading");
var anchorExplanation = anchoredFlow.Pages.SelectMany(item => item.Runs).First(run => run.Text == "Floating");
Check(anchorHeading.X >= 315 && anchorExplanation.X >= 315 && anchorExplanation.Y > anchorHeading.Y, "A floating text box must move its heading with it into the next column.");
var footerAnchor = new XElement(anchor);
footerAnchor.Element(WP + "positionV")!.SetAttributeValue("relativeFrom", "page");
footerAnchor.Element(WP + "positionV")!.Element(WP + "posOffset")!.Value = (200 * 12700).ToString();
var absoluteHeader = headerPackage.WithXml("word/synthetic-header.xml", new XDocument(new XElement(W + "hdr", new XElement(W + "p", new XElement(W + "r", new XElement(W + "drawing", footerAnchor))))));
using var absoluteHeaderEngine = new DocxLayout(absoluteHeader);
Check(Math.Abs(absoluteHeaderEngine.Build().Pages[0].Runs.First(run => run.Text == "Floating").Y - 203) < .01, "Page-relative header text boxes must not receive the header flow offset twice.");
var footerFrame = Paragraph("42");
footerFrame.AddFirst(new XElement(W + "pPr", new XElement(W + "framePr", new XAttribute(W + "hAnchor", "margin"), new XAttribute(W + "vAnchor", "text"), new XAttribute(W + "xAlign", "right"), new XAttribute(W + "y", 1))));
var footerPackage = headerPackage;
var footerDocument = footerPackage.ReadXml("word/document.xml");
footerDocument.Descendants(W + "headerReference").Single().Name = W + "footerReference";
footerPackage = footerPackage.WithXml("word/document.xml", footerDocument).WithXml("word/synthetic-header.xml", new XDocument(new XElement(W + "ftr", footerFrame, Paragraph("example.test"))));
using var footerEngine = new DocxLayout(footerPackage);
var footerLayout = footerEngine.Build();
var pageNumberRun = footerLayout.Pages[0].Runs.Single(run => run.Text == "42");
Check(Math.Abs(pageNumberRun.X + pageNumberRun.Width - (footerLayout.Width - 72)) < .01, "Automatic-width footer frame must align its text against the right margin.");
Check(pageNumberRun.Y > footerLayout.Height - 100 && pageNumberRun.Y < footerLayout.Height, "Text-anchored footer frame must stay in footer flow.");
var paddedAnchor = new XElement(anchor);
paddedAnchor.Add(new XElement(wordShape + "bodyPr", new XAttribute("lIns", 127000), new XAttribute("rIns", 127000), new XAttribute("tIns", 127000), new XAttribute("bIns", 127000), new XAttribute("anchor", "b")));
using var paddingEngine = new DocxLayout(Synthetic(new XElement(W + "p", new XElement(W + "r", new XElement(W + "drawing", paddedAnchor)))));
var paddedLayout = paddingEngine.Build();
var paddedRun = paddedLayout.Pages[0].Runs.First(run => run.Text == "Floating");
Check(Math.Abs(paddedRun.X - 82) < .01 && paddedRun.Y > 120, "Text-box insets and bottom alignment must position the text inside its bounds.");
Check(Math.Abs(paddedLayout.Pages[0].TextObjects.First(item => item.Text == "Floating explanation").Bounds.Y - paddedRun.Y) < .01, "Text-box selection bounds must follow vertical alignment.");
var parityPackage = source;
var parityDocument = parityPackage.ReadXml("word/document.xml");
var paritySection = parityDocument.Descendants(W + "sectPr").Last();
paritySection.Add(new XElement(W + "titlePg", new XAttribute(W + "val", 0)), new XElement(W + "footerReference", new XAttribute(W + "type", "default"), new XAttribute(R + "id", "oddFooter")), new XElement(W + "footerReference", new XAttribute(W + "type", "even"), new XAttribute(R + "id", "evenFooter")));
var parityRelations = parityPackage.ReadXml("word/_rels/document.xml.rels");
foreach (var parity in new[] { "odd", "even" })
{
    parityRelations.Root!.Add(new XElement(REL + "Relationship", new XAttribute("Id", parity + "Footer"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"), new XAttribute("Target", parity + "-footer.xml")));
    parityPackage = parityPackage.WithXml("word/" + parity + "-footer.xml", new XDocument(new XElement(W + "ftr", Paragraph(parity + " footer"), new XElement(W + "p", new XElement(W + "fldSimple", new XAttribute(W + "instr", "PAGE"), new XElement(W + "r", new XElement(W + "t", "99")))), new XElement(W + "p", new XElement(W + "fldSimple", new XAttribute(W + "instr", "NUMPAGES"))))));
}
parityPackage = parityPackage.WithXml("word/document.xml", parityDocument).WithXml("word/_rels/document.xml.rels", parityRelations).WithXml("word/settings.xml", new XDocument(new XElement(W + "settings", new XElement(W + "evenAndOddHeaders"))));
using var parityEngine = new DocxLayout(parityPackage);
var parityLayout = parityEngine.Build();
Check(parityLayout.Pages[0].Runs.Any(run => run.Text == "odd") && parityLayout.Pages[1].Runs.Any(run => run.Text == "even"), "Distinct odd/even footers and disabled first-page override must be respected.");
Check(parityLayout.Pages.All(item => !item.Runs.Any(run => run.Text == "99")) && parityLayout.Pages[0].Runs.Any(run => run.Text == "1") && parityLayout.Pages[1].Runs.Count(run => run.Text == "2") >= 2, "Simple PAGE and NUMPAGES fields must use current pagination, including empty cached results.");
var styledTable = new XElement(W + "tbl", new XElement(W + "tblPr", new XElement(W + "tblStyle", new XAttribute(W + "val", "SyntheticTable")), new XElement(W + "tblLook", new XAttribute(W + "firstRow", 1))), new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 4000))), new XElement(W + "tr", new XElement(W + "tc", Paragraph("Header"))), new XElement(W + "tr", new XElement(W + "tc", Paragraph("Body"))));
var styledPackage = Synthetic(styledTable);
var tableStyles = styledPackage.ReadXml("word/styles.xml");
tableStyles.Root!.Add(new XElement(W + "style", new XAttribute(W + "type", "table"), new XAttribute(W + "styleId", "SyntheticTable"), new XElement(W + "rPr", new XElement(W + "sz", new XAttribute(W + "val", 20))), new XElement(W + "tblStylePr", new XAttribute(W + "type", "firstRow"), new XElement(W + "rPr", new XElement(W + "color", new XAttribute(W + "val", "FFFFFF"))), new XElement(W + "tcPr", new XElement(W + "shd", new XAttribute(W + "fill", "345678"))))));
using var tableStyleEngine = new DocxLayout(styledPackage.WithXml("word/styles.xml", tableStyles));
var tableStylePage = tableStyleEngine.Build().Pages[0];
Check(tableStylePage.Rules[0].Fill == "345678" && tableStylePage.Rules[1].Fill is null, "Conditional first-row shading must not leak into body rows.");
Check(tableStylePage.Runs.First(run => run.Text == "Header").Color == "FFFFFF" && tableStylePage.Runs.All(run => run.FontSize == 10), "Table-style font formatting must apply to cell paragraphs.");
var tabsParagraph = new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "tabs", new XElement(W + "tab", new XAttribute(W + "val", "left"), new XAttribute(W + "pos", 1000)), new XElement(W + "tab", new XAttribute(W + "val", "right"), new XAttribute(W + "pos", 6000), new XAttribute(W + "leader", "dot")))), new XElement(W + "r", new XElement(W + "t", "A"), new XElement(W + "tab"), new XElement(W + "t", "Entry"), new XElement(W + "tab"), new XElement(W + "t", "12")));
using var tabsEngine = new DocxLayout(Synthetic(tabsParagraph));
var tabRuns = tabsEngine.Build().Pages[0].Runs;
Check(Math.Abs(tabRuns.Single(run => run.Text == "Entry").X - 122) < .01, "The first tab stop was lost while merging paragraph properties.");
var tabNumber = tabRuns.Single(run => run.Text == "12");
Check(Math.Abs(tabNumber.X + tabNumber.Width - 372) < .01 && tabRuns.Any(run => run.Text.StartsWith("...")), "TOC page numbers need right-aligned tabs and dot leaders.");
XElement Spaced(string text, int before, int after, bool contextual = false)
{
    var paragraph = Paragraph(text);
    paragraph.AddFirst(new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "before", before), new XAttribute(W + "after", after)), contextual ? new XElement(W + "contextualSpacing") : null));
    return paragraph;
}
using var spacingControlEngine = new DocxLayout(Synthetic(Spaced("One", 0, 0), Spaced("Two", 0, 0)));
var spacingControl = spacingControlEngine.Build().Pages[0].Runs;
var normalAdvance = spacingControl[1].Y - spacingControl[0].Y;
using var spacingEngine = new DocxLayout(Synthetic(Spaced("One", 0, 200), Spaced("Two", 120, 0)));
var spacedRuns = spacingEngine.Build().Pages[0].Runs;
Check(Math.Abs(spacedRuns[1].Y - spacedRuns[0].Y - normalAdvance - 10) < .01, "Adjacent paragraph spacing must use the larger gap rather than sum both gaps.");
using var contextualEngine = new DocxLayout(Synthetic(Spaced("One", 0, 200, true), Spaced("Two", 120, 0, true)));
var contextualRuns = contextualEngine.Build().Pages[0].Runs;
Check(Math.Abs(contextualRuns[1].Y - contextualRuns[0].Y - normalAdvance) < .01, "Contextual spacing must suppress gaps between matching paragraph styles.");
var sectionBreakParagraph = Paragraph("Before");
sectionBreakParagraph.AddFirst(new XElement(W + "pPr", new XElement(W + "sectPr")));
sectionBreakParagraph.Add(new XElement(W + "r", new XElement(W + "br", new XAttribute(W + "type", "page"))));
using var combinedBreakEngine = new DocxLayout(Synthetic(sectionBreakParagraph, Paragraph("After")));
Check(combinedBreakEngine.Build().Pages.Count == 2, "A section break after a page break must not create an extra empty page.");
var repeatingTable = new XElement(table);
repeatingTable.Elements(W + "tr").First().AddFirst(new XElement(W + "trPr", new XElement(W + "tblHeader")));
using var repeatHeaderEngine = new DocxLayout(Synthetic(repeatingTable));
var repeatHeaderLayout = repeatHeaderEngine.Build();
Check(repeatHeaderLayout.Pages.Count > 1 && repeatHeaderLayout.Pages.All(item => item.Runs.Any(run => run.Id == "p0")), "Repeated table header rows must appear on continuation pages.");
var shadedParagraph = Paragraph("Shaded text");
shadedParagraph.Element(W + "r")!.AddFirst(new XElement(W + "rPr", new XElement(W + "shd", new XAttribute(W + "fill", "CCDDEE"))));
using var shadingEngine = new DocxLayout(Synthetic(shadedParagraph));
Check(shadingEngine.Build().Pages[0].Runs.All(run => run.Background == "CCDDEE"), "Run shading must survive layout for PDF and browser painting.");
var numberedBlocks = new List<XElement>();
for (var number = 0; number < 7; number++)
{
    var paragraph = Paragraph("Physical page " + number);
    if (number is 3 or 5)
    {
        var section = new XElement(W + "sectPr");
        if (number == 3)
        {
            section.Add(new XElement(W + "footerReference", new XAttribute(W + "type", "default"), new XAttribute(R + "id", "syntheticHeader")));
        }
        else
        {
            section.Add(new XElement(W + "pgNumType", new XAttribute(W + "start", 1)));
        }
        paragraph.AddFirst(new XElement(W + "pPr", section));
    }
    else if (number != 6)
    {
        paragraph.Add(new XElement(W + "r", new XElement(W + "br", new XAttribute(W + "type", "page"))));
    }
    numberedBlocks.Add(paragraph);
}
var restartedPackage = Synthetic(numberedBlocks.ToArray()).WithXml("word/_rels/document.xml.rels", headerRelations).WithXml("word/synthetic-header.xml", new XDocument(new XElement(W + "ftr",
    new XElement(W + "p", new XElement(W + "fldSimple", new XAttribute(W + "instr", "PAGE"))),
    new XElement(W + "p", new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))), new XElement(W + "r", new XElement(W + "instrText", " PAGE ")), new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))), new XElement(W + "r", new XElement(W + "t", "99")), new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end")))),
    new XElement(W + "p", new XElement(W + "fldSimple", new XAttribute(W + "instr", "NUMPAGES"))))));
using var restartedEngine = new DocxLayout(restartedPackage);
var restartedLayout = restartedEngine.Build();
Check(restartedLayout.Pages.Count == 7 && restartedLayout.Pages.Select(item => item.Number).SequenceEqual(Enumerable.Range(1, 7)), "Section numbering must not change physical navigation indices.");
var displayedNumbers = new[] { "1", "2", "3", "4", "1", "2", "3" };
for (var index = 0; index < displayedNumbers.Length; index++)
{
    var footerRuns = restartedLayout.Pages[index].Runs.Where(run => run.Id.Contains("-footer-")).ToArray();
    Check(footerRuns.Count(run => run.Text == displayedNumbers[index]) == 2 && footerRuns.Any(run => run.Text == "7"), "Section PAGE fields must restart and continue while NUMPAGES retains physical total.");
}
// Independent edge formatting must not turn an open frame into a cell grid.
XElement TestBorder(string side, string value, int size = 24, string color = "2367AB") =>
    new(W + side, new XAttribute(W + "val", value), new XAttribute(W + "sz", size), new XAttribute(W + "color", color));
var borderTable = new XElement(W + "tbl",
    new XElement(W + "tblPr", new XElement(W + "tblBorders", TestBorder("top", "single"), TestBorder("bottom", "single"), TestBorder("left", "single"), TestBorder("right", "single"), TestBorder("insideV", "single", 8, "008800"))),
    new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 2400)), new XElement(W + "gridCol", new XAttribute(W + "w", 2400))),
    new XElement(W + "tr",
        new XElement(W + "tc", new XElement(W + "tcPr", new XElement(W + "tcBorders", TestBorder("right", "nil"), TestBorder("bottom", "none"))), Paragraph("First example")),
        new XElement(W + "tc", Paragraph("Second example"))));
using var borderEngine = new DocxLayout(Synthetic(borderTable));
var borderRules = borderEngine.Build().Pages[0].Rules;
Check(borderRules[0].StrokeWidth == 0 && borderRules[0].Edges is [not null, null, null, not null], "Explicit hidden cell borders must suppress only their own edges.");
Check(borderRules[0].Edges![0] == new DocxBorder(3, "2367AB"), "Border size must be converted from eighth-points and retain its color.");
Check(borderRules[1].Edges![3] == new DocxBorder(1, "008800") && borderRules[1].Edges![1] == new DocxBorder(3, "2367AB"), "Interior borders must not inherit exterior border formatting.");
var markShading = new XElement(W + "p",
    new XElement(W + "pPr", new XElement(W + "rPr",
        new XElement(W + "shd", new XAttribute(W + "val", "clear"), new XAttribute(W + "fill", "CCCCCC")))),
    new XElement(W + "r", new XElement(W + "t", "Label")),
    new XElement(W + "r", new XElement(W + "rPr",
        new XElement(W + "shd", new XAttribute(W + "val", "clear"), new XAttribute(W + "fill", "FFFF00"))),
        new XElement(W + "t", "Marked")));
using var markShadingEngine = new DocxLayout(Synthetic(markShading));
var markRuns = markShadingEngine.Build().Pages[0].Runs;
Check(markRuns.Single(run => run.Text == "Label").Background is null, "Paragraph-mark shading must not appear behind text.");
Check(markRuns.Single(run => run.Text == "Marked").Background == "FFFF00", "Explicit text-run shading must be preserved.");
var partialIndentHeading = Paragraph("Synthetic heading");
partialIndentHeading.AddFirst(new XElement(W + "pPr",
    new XElement(W + "numPr", new XElement(W + "numId", new XAttribute(W + "val", 90)), new XElement(W + "ilvl", new XAttribute(W + "val", 2))),
    new XElement(W + "ind", new XAttribute(W + "hanging", 180))));
var headingNumbering = new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", 90));
for (var levelIndex = 0; levelIndex < 3; levelIndex++)
{
    headingNumbering.Add(new XElement(W + "lvl", new XAttribute(W + "ilvl", levelIndex),
        new XElement(W + "start", new XAttribute(W + "val", levelIndex == 1 ? 4 : 1)),
        new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")),
        new XElement(W + "lvlText", new XAttribute(W + "val", "%1.%2.%3")),
        new XElement(W + "pPr", new XElement(W + "ind", new XAttribute(W + "left", 720), new XAttribute(W + "hanging", 360)))));
}
var headingPackage = Synthetic(partialIndentHeading).WithXml("word/numbering.xml", new XDocument(new XElement(W + "numbering", headingNumbering,
    new XElement(W + "num", new XAttribute(W + "numId", 90), new XElement(W + "abstractNumId", new XAttribute(W + "val", 90))))));
using var headingEngine = new DocxLayout(headingPackage);
var headingRuns = headingEngine.Build().Pages[0].Runs;
Check(headingRuns.Any(run => run.Text == "1.4.1"), "Ancestor numbering placeholders must honor their level start values.");
Check(Math.Abs(headingRuns.First().X - 99) < .01, "Partial indentation must retain the numbering left indent and explicit paragraph hanging indent.");
var croppedXml = source.ReadXml("word/document.xml");
var croppedPicture = croppedXml.Descendants(A + "blip").First().Parent!;
croppedPicture.Add(new XElement(A + "srcRect", new XAttribute("l", 25000), new XAttribute("r", 12500), new XAttribute("t", -10000)));
var pictureTransform = croppedXml.Descendants(A + "xfrm").First();
pictureTransform.SetAttributeValue("rot", 5400000);
pictureTransform.SetAttributeValue("flipH", "1");
pictureTransform.SetAttributeValue("flipV", "true");
var croppedPackage = source.WithXml("word/document.xml", croppedXml);
using var cropEngine = new DocxLayout(croppedPackage);
var cropLayout = cropEngine.Build();
var croppedImage = cropLayout.Pages[0].Images[0];
Check(croppedImage.CropLeft == .25 && croppedImage.CropRight == .125 && croppedImage.CropTop == -.1, "Image crop fractions and negative padding must be retained.");
Check(croppedImage.Rotation == 90 && croppedImage.FlipH && croppedImage.FlipV, "Drawing transforms must retain rotation and flips.");
Check(!cropLayout.Warnings.Any(note => note.Contains("cropping")), "Supported crop should not produce an unsupported warning.");
File.WriteAllBytes(Path.Combine(AppContext.BaseDirectory, "cropped.pdf"), DocxPdfConverter.Convert(croppedPackage).Bytes);

foreach (var (numberFormat, startValue, expected) in new[] { ("upperRoman", 9, "IX."), ("lowerRoman", 4, "iv."), ("upperLetter", 27, "AA."), ("decimalZero", 0, "00.") })
{
    var formatXml = numberPackage.ReadXml("word/numbering.xml");
    var formatLevel = formatXml.Descendants(W + "lvl").First();
    formatLevel.Element(W + "numFmt")!.SetAttributeValue(W + "val", numberFormat);
    formatLevel.Element(W + "start")!.SetAttributeValue(W + "val", startValue);
    using var formatEngine = new DocxLayout(numberPackage.WithXml("word/numbering.xml", formatXml));
    var formatRuns = formatEngine.Build().Pages.SelectMany(item => item.Runs).ToArray();
    Check(formatRuns.Any(run => run.Text == expected), "Numbering format/start failed: " + numberFormat);
    if (startValue == 0)
    {
        Check(formatRuns.Any(run => run.Text == "01."), "Zero-based numbering must increment instead of restarting at zero.");
    }
}
var restartXml = headingPackage.ReadXml("word/numbering.xml");
restartXml.Descendants(W + "lvl").Last().Add(new XElement(W + "lvlRestart", new XAttribute(W + "val", 0)));
var upperHeading = new XElement(partialIndentHeading);
upperHeading.Descendants(W + "ilvl").Single().SetAttributeValue(W + "val", 0);
var restartPackage = Synthetic(new XElement(partialIndentHeading), upperHeading, new XElement(partialIndentHeading)).WithXml("word/numbering.xml", restartXml);
using var restartEngine = new DocxLayout(restartPackage);
Check(restartEngine.Build().Pages.SelectMany(item => item.Runs).Any(run => run.Text == "1.4.2"), "Never-restart list levels must survive parent headings.");
XElement FixedLine(string text, int points, params XElement[] flags)
{
    var paragraph = Paragraph(text);
    paragraph.AddFirst(new XElement(W + "pPr", flags, new XElement(W + "spacing", new XAttribute(W + "before", 0), new XAttribute(W + "after", 0), new XAttribute(W + "line", points * 20), new XAttribute(W + "lineRule", "exact"))));
    return paragraph;
}
var pageFiller = FixedLine("Synthetic preceding content", 620);
var keepHeading = FixedLine("Kept heading", 20, new XElement(W + "keepNext"));
var followingParagraph = FixedLine("Following text", 20);
using var keepNextEngine = new DocxLayout(Synthetic(new XElement(pageFiller), keepHeading, followingParagraph));
var keptPages = keepNextEngine.Build().Pages;
Check(keptPages.Count == 2 && keptPages[1].TextObjects.Count == 2 && keptPages[0].TextObjects.Count == 1, "Keep-next must move a heading with its following text without duplicating objects.");
var wholeParagraph = FixedLine("First line", 20, new XElement(W + "keepLines"));
wholeParagraph.Add(new XElement(W + "r", new XElement(W + "br"), new XElement(W + "t", "Second line")));
using var keepLinesEngine = new DocxLayout(Synthetic(new XElement(pageFiller), wholeParagraph));
Check(keepLinesEngine.Build().Pages[1].Runs.Count(run => run.Text is "First" or "Second") == 2, "Keep-lines must move both lines to the next page.");
var disabledKeep = new XElement(wholeParagraph);
disabledKeep.Descendants(W + "keepLines").Single().SetAttributeValue(W + "val", 0);
disabledKeep.Element(W + "pPr")!.Add(new XElement(W + "widowControl", new XAttribute(W + "val", 0)));
using var splitEngine = new DocxLayout(Synthetic(new XElement(pageFiller), disabledKeep));
var splitPages = splitEngine.Build().Pages;
Check(splitPages[0].Runs.Any(run => run.Text == "First") && splitPages[1].Runs.Any(run => run.Text == "Second"), "Disabled keep-lines must allow normal splitting.");
var numberedHeading = FixedLine("Numbered heading", 20, new XElement(W + "keepNext"), new XElement(W + "numPr", new XElement(W + "numId", new XAttribute(W + "val", 1))));
var numberedFollowing = FixedLine("Numbered body", 20, new XElement(W + "numPr", new XElement(W + "numId", new XAttribute(W + "val", 1))));
var keptNumberPackage = Synthetic(new XElement(pageFiller), numberedHeading, numberedFollowing).WithXml("word/numbering.xml", numberPackage.ReadXml("word/numbering.xml"));
using var keptNumberEngine = new DocxLayout(keptNumberPackage);
var keptNumbers = keptNumberEngine.Build().Pages.SelectMany(item => item.Runs).Where(run => run.Text is "1." or "2." or "3.").Select(run => run.Text).ToArray();
Check(keptNumbers.SequenceEqual(new[] { "1.", "2." }), "Pagination measurement must not advance numbering counters.");
var longKept = FixedLine("Long first line", 20, new XElement(W + "keepLines"));
for (var lineIndex = 0; lineIndex < 40; lineIndex++)
{
    longKept.Add(new XElement(W + "r", new XElement(W + "br"), new XElement(W + "t", "Additional line")));
}
using var longKeepEngine = new DocxLayout(Synthetic(new XElement(pageFiller), longKept));
var longKeepPages = longKeepEngine.Build().Pages;
Check(longKeepPages.Count == 3 && longKeepPages.Sum(item => item.Runs.Count(run => run.Text == "Additional")) == 40, "Oversized keep-lines paragraphs must split without losing text or creating blank pages.");
var bulletXml = numberPackage.ReadXml("word/numbering.xml");
var bulletLevel = bulletXml.Descendants(W + "lvl").First();
bulletLevel.Element(W + "numFmt")!.SetAttributeValue(W + "val", "bullet");
bulletLevel.Element(W + "lvlText")!.SetAttributeValue(W + "val", "\uf0b7");
bulletLevel.Add(new XElement(W + "rPr", new XElement(W + "rFonts", new XAttribute(W + "ascii", "Wingdings"))));
using var bulletEngine = new DocxLayout(numberPackage.WithXml("word/numbering.xml", bulletXml));
var bulletLayout = bulletEngine.Build();
var bulletRun = bulletLayout.Pages[0].Runs.First();
Check(bulletRun.Text == "\u2022" && !bulletLayout.Warnings.Any(note => note.Contains("Some characters are unavailable")), "Normalized symbol-font bullets must use a font containing the replacement glyph.");
var largeMarkerXml = numberPackage.ReadXml("word/numbering.xml");
largeMarkerXml.Descendants(W + "lvl").First().Add(new XElement(W + "rPr", new XElement(W + "sz", new XAttribute(W + "val", 48))));
using var largeMarkerEngine = new DocxLayout(numberPackage.WithXml("word/numbering.xml", largeMarkerXml));
var largeMarkerRuns = largeMarkerEngine.Build().Pages.SelectMany(item => item.Runs).Where(run => run.Text == "List").ToArray();
var normalMarkerRuns = numberedLayout.Pages.SelectMany(item => item.Runs).Where(run => run.Text == "List").ToArray();
Check(largeMarkerRuns.Length == 2 && Math.Abs((largeMarkerRuns[1].Baseline - largeMarkerRuns[0].Baseline) - (normalMarkerRuns[1].Baseline - normalMarkerRuns[0].Baseline)) < .01, "Numbering marker font metrics must not inflate paragraph line spacing.");
XNamespace legacyVml = "urn:schemas-microsoft-com:vml";
var legacyRelation = (string)source.ReadXml("word/document.xml").Descendants(A + "blip").First().Attribute(R + "embed")!;
var legacyParagraph = new XElement(W + "p", new XElement(W + "r", new XElement(W + "object",
    new XElement(legacyVml + "shape", new XAttribute("style", "width:120pt;height:80pt"),
        new XElement(legacyVml + "imagedata", new XAttribute(R + "id", legacyRelation))))));
using var legacyEngine = new DocxLayout(Synthetic(legacyParagraph));
var legacyImage = legacyEngine.Build().Pages[0].Images.Single();
Check(legacyImage.Width == 120 && legacyImage.Height == 80 && legacyImage.ReadOnly, "Embedded-object preview must retain dimensions and be read-only.");
File.WriteAllBytes(Path.Combine(AppContext.BaseDirectory, "sample.pdf"), pdf.Bytes);
Console.WriteLine($"Passed {count} DOCX native layout/PDF checks.");
