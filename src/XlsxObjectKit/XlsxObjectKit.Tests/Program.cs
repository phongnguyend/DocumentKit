using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XlsxObjectKit.Application;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using XlsxObjectKit.Readers;
using XlsxObjectKit.Writers;
using static XlsxObjectKit.Models.XlsxNamespaces;

if (args is ["--write-formatting-fixture", var fixturePath])
{
    File.WriteAllBytes(fixturePath, XlsxWriter.Write(FormattingFixture.Create()));
    return;
}
if (args is ["--write-calculation-fixture", var calculationPath])
{
    var fixture = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("A1", "1", "number"), new("A2", "2", "number"), new("B1", "SUM(A1:A2)", "formula")]);
    var xml = fixture.ReadXml("xl/worksheets/sheet1.xml");
    xml.Descendants(S + "c").Single(cell => (string?)cell.Attribute("r") == "B1").Add(new XElement(S + "v", "999"));
    File.WriteAllBytes(calculationPath, XlsxWriter.Write(fixture.WithXml("xl/worksheets/sheet1.xml", xml)));
    return;
}

var checks = 0;
void Check(bool condition, string message)
{
    if (!condition)
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
    catch (Exception exception) when (exception is ArgumentException or InvalidDataException or NotSupportedException)
    {
        checks++;
        return;
    }
    throw new Exception("Invalid operation was accepted.");
}
void Validate(XlsxObjectDocument document)
{
    using var package = SpreadsheetDocument.Open(new MemoryStream(XlsxWriter.Write(document)), false);
    var errors = new OpenXmlValidator().Validate(package).ToArray();
    Check(errors.Length == 0, string.Join("\n", errors.Select(error => error.Description)));
}
var source = XlsxEditor.Create().WithPart("customXml/preserved.xml", "<custom>Keep me</custom>"u8.ToArray());
Validate(source);
var edited = XlsxEditor.SetCells(source, "1", [new("A1", "Hello <&> café"), new("B2", "42.5", "number"), new("C3", "true", "boolean")]);
Check(XlsxWorkbookExtractor.Extract(source).Sheets[0].Cells.Count == 0, "Edits preserve source");
Check(edited.GetPart("customXml/preserved.xml").SequenceEqual(source.GetPart("customXml/preserved.xml")), "Unknown parts preserved");
var roundTrip = XlsxWorkbookExtractor.Extract(XlsxReader.Read(XlsxWriter.Write(edited)));
Check(roundTrip.Sheets[0].Cells[0].Value == "Hello <&> café" && roundTrip.Sheets[0].Cells[1].Value == "42.5", "Text/number round trip");
Check(roundTrip.Sheets[0].Cells[2].Type == "b" && roundTrip.Sheets[0].Cells[2].Value == "1", "Boolean round trip");
edited = XlsxEditor.SetCells(edited, "1", [new("A1", "")], new(true, true, "#ff0000", "#ccffee", "center"));
Validate(edited);
var styles = XlsxWorkbookExtractor.Extract(edited);
var style = styles.Styles[styles.Sheets[0].Cells[0].Style];
Check(style.Bold && style.Italic && style.Background == "#ccffee" && style.Alignment == "center", "Style rendered");
edited = XlsxEditor.Dimensions(edited, "1", "B2", 40, 25);
Validate(edited);
Check(XlsxWorkbookExtractor.Extract(edited).Sheets[0].ColumnWidths[2] == 25, "Column width saved");
var added = XlsxEditor.Worksheet(edited, "add-sheet", "1", "Budget");
Validate(added);
Check(XlsxWorkbookExtractor.Extract(added).Sheets.Count == 2, "Add sheet");
var reordered = XlsxEditor.Worksheet(added, "move-sheet", "2", position: 0);
Check(XlsxWorkbookExtractor.Extract(reordered).Sheets[0].Name == "Budget", "Reorder sheet");
var viewedBook = added.ReadXml("xl/workbook.xml");
viewedBook.Root!.Element(S + "sheets")!.AddBeforeSelf(new XElement(S + "bookViews", new XElement(S + "workbookView", new XAttribute("activeTab", "1"))));
viewedBook.Root.Element(S + "sheets")!.AddAfterSelf(new XElement(S + "definedNames", new XElement(S + "definedName", new XAttribute("name", "LocalName"), new XAttribute("localSheetId", "1"), "Budget!$A$1")));
var viewed = XlsxEditor.Worksheet(added.WithXml("xl/workbook.xml", viewedBook), "move-sheet", "2", position: 0);
Check(viewed.ReadXml("xl/workbook.xml").Descendants(S + "workbookView").Single().Attribute("activeTab")?.Value == "0", "Active worksheet follows reorder");
Check(viewed.ReadXml("xl/workbook.xml").Descendants(S + "definedName").Single().Attribute("localSheetId")?.Value == "0", "Local named range follows reorder");
Validate(viewed);
var renamed = XlsxEditor.Worksheet(reordered, "rename-sheet", "2", "Annual budget");
Check(XlsxWorkbookExtractor.Extract(renamed).Sheets[0].Name == "Annual budget", "Rename sheet");
var deleted = XlsxEditor.Worksheet(renamed, "delete-sheet", "2");
Validate(deleted);
Check(XlsxWorkbookExtractor.Extract(deleted).Sheets.Count == 1, "Delete sheet");
Reject(() => XlsxEditor.Worksheet(deleted, "delete-sheet", "1"));
Reject(() => XlsxEditor.Worksheet(added, "rename-sheet", "2", "sheet1"));
Reject(() => XlsxEditor.SetCells(edited, "1", [new("XFE1", "bad")]));
Reject(() => XlsxEditor.SetCells(edited, "1", [new("A1", "NaN", "number")]));
Reject(() => XlsxEditor.Dimensions(edited, "1", "A1", double.NaN, 10));
var formulas = XlsxEditor.SetCells(edited, "1", [new("D4", "=B2*2", "formula")]);
Validate(formulas);
Check(XlsxWorkbookExtractor.Extract(formulas).Sheets[0].Cells.Single(cell => cell.Address == "D4").Formula == "B2*2", "Formula preserved");
Check(formulas.ReadXml("xl/workbook.xml").Descendants(S + "calcPr").Single().Attribute("fullCalcOnLoad")?.Value == "1", "Recalculation requested");
Reject(() => XlsxEditor.Worksheet(formulas, "rename-sheet", "1", "New name"));
var sheet = edited.ReadXml("xl/worksheets/sheet1.xml");
sheet.Root!.Add(new XElement(S + "mergeCells", new XElement(S + "mergeCell", new XAttribute("ref", "D6:E6"))));
var merged = edited.WithXml("xl/worksheets/sheet1.xml", sheet);
Reject(() => XlsxEditor.SetCells(merged, "1", [new("E6", "hidden")]));
sheet.Root.Add(new XElement(S + "sheetProtection", new XAttribute("sheet", "1")));
Reject(() => XlsxEditor.SetCells(edited.WithXml("xl/worksheets/sheet1.xml", sheet), "1", [new("A1", "bad")]));
Reject(() => XlsxReader.Read([1, 2, 3]));
void RejectWith(byte[] input, string message)
{
    try
    {
        XlsxReader.Read(input);
    }
    catch (Exception exception) when (exception is InvalidDataException or NotSupportedException)
    {
        Check(exception.Message.Contains(message, StringComparison.Ordinal), "Actionable upload diagnostic: " + exception.Message);
        return;
    }
    throw new Exception("Invalid workbook was accepted.");
}
RejectWith([], "empty");
RejectWith([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1], "compound container");
RejectWith("<html>Not an XLSX</html>"u8.ToArray(), "not an XLSX ZIP package");
RejectWith([0x50, 0x4B, 0x03, 0x04, 0, 0], "incomplete or damaged");
var bytes = WorkbookApplication.Edit(XlsxWriter.Write(source), new("cells", "1", [new("Z100", "API edit")]));
Check(WorkbookApplication.Inspect(bytes).Sheets[0].Cells.Single().Address == "Z100", "Application edits round trip");
var cachedSheet = formulas.ReadXml("xl/worksheets/sheet1.xml");
cachedSheet.Descendants(S + "c").Single(item => (string?)item.Attribute("r") == "D4").Add(new XElement(S + "v", "85"));
var recalculated = XlsxEditor.SetCells(formulas.WithXml("xl/worksheets/sheet1.xml", cachedSheet), "1", [new("B2", "50", "number")]);
Check(XlsxWorkbookExtractor.Extract(recalculated).Sheets[0].Cells.Single(item => item.Address == "D4").Value == "", "Dependent formula cache invalidated");
var existingStyles = edited.ReadXml("xl/styles.xml");
existingStyles.Descendants(S + "cellXfs").Single().Elements().Last().SetAttributeValue("numFmtId", "2");
var styledAgain = XlsxEditor.SetCells(edited.WithXml("xl/styles.xml", existingStyles), "1", [new("A1", "")], new());
Check(styledAgain.ReadXml("xl/styles.xml").Descendants(S + "cellXfs").Single().Elements().Last().Attribute("numFmtId")?.Value == "2", "Existing numeric formatting preserved");
Validate(styledAgain);
var formatted = XlsxWorkbookExtractor.Extract(FormattingFixture.Create());
Check(formatted.Date1904, "1904 date system extracted");
Check(formatted.Styles[1].FontName == "Arial" && formatted.Styles[1].FontSize == 20 && formatted.Styles[1].Strike && formatted.Styles[1].Underline == "double", "Font formatting extracted");
Check(formatted.Styles[1].Color == "#008000" && formatted.Styles[1].Background == "#40ff40", "Theme color and luminance tint");
Check(formatted.Styles[1].Borders["left"].Color == "#0000ff" && formatted.Styles[1].Borders["bottom"].Style == "double", "Indexed and RGB borders");
Check(formatted.Styles[1].NumberFormat == "#,##0.00" && formatted.Styles[2].NumberFormatId == 10, "Custom and built-in number formats");
Check(formatted.Styles[4].Bold && formatted.Styles[4].WrapText && formatted.Styles[4].VerticalAlignment == "center", "Base style and direct alignment inheritance");
Check(formatted.Sheets[0].Cells.Single(cell => cell.Address == "A4").Style == 4 && formatted.Sheets[0].Cells.Single(cell => cell.Address == "B5").Style == 4, "Row and column style inheritance");
Check(formatted.Sheets[0].Cells.Single(cell => cell.Address == "B4").Style == 0, "Explicit default style overrides inherited style");
Check(formatted.Sheets[0].DefaultRowHeight == 22 && formatted.Sheets[0].DefaultColumnWidth == 18, "Default worksheet dimensions extracted");
var inheritedEdit = XlsxWorkbookExtractor.Extract(XlsxEditor.SetCells(FormattingFixture.Create(), "1", [new("A4", "")], new(true)));
var inheritedStyle = inheritedEdit.Styles[inheritedEdit.Sheets[0].Cells.Single(cell => cell.Address == "A4").Style];
Check(inheritedStyle.FontSize == 20 && inheritedStyle.WrapText && inheritedStyle.VerticalAlignment == "center", "Formatting edits retain inherited font size and alignment");
checks += CalculationTests.Run();
Console.WriteLine($"Passed {checks} XLSX package, editing, formatting, calculation, worksheet and validation checks.");
