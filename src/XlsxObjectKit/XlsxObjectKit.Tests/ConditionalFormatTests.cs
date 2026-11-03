using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XlsxObjectKit.Application;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Readers;
using XlsxObjectKit.Writers;
using System.Xml.Linq;
using static XlsxObjectKit.Models.XlsxNamespaces;

internal static class ConditionalFormatTests
{
    internal static XlsxObjectKit.Models.XlsxObjectDocument ReferenceFixture()
    {
        var document = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("A1", "12", "number"), new("A2", "5", "number"), new("D1", "100", "number")]);
        document = XlsxConditionalFormatEditor.Edit(document, "1", "A1:A2", "notBetween", -5, 5, "#ffffff", "#ffff00");
        var sheet = document.ReadXml("xl/worksheets/sheet1.xml");
        var formulas = sheet.Descendants(S + "cfRule").Single().Elements(S + "formula").ToArray();
        formulas[0].Value = "-5%*$D$1";
        formulas[1].Value = "5%*$D$1";
        var styles = document.ReadXml("xl/styles.xml");
        var dxf = styles.Descendants(S + "dxf").Single();
        dxf.Element(S + "font")!.Element(S + "color")!.ReplaceAttributes(new XAttribute("theme", 0));
        var fill = dxf.Element(S + "fill")!.Element(S + "patternFill")!;
        fill.RemoveAttributes();
        fill.ReplaceNodes(new XElement(S + "bgColor", new XAttribute("rgb", "FFFF0000")));
        return document.WithXml("xl/worksheets/sheet1.xml", sheet).WithXml("xl/styles.xml", styles);
    }

    internal static int Run()
    {
        var checks = 0;
        void Check(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(message);
            }
            checks++;
        }
        var source = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("A1", "12", "number"), new("A2", "5", "number")]);
        var reference = ReferenceFixture();
        var referenceRule = XlsxWorkbookExtractor.Extract(reference).Sheets[0].ConditionalRules.Single();
        Check(referenceRule.Supported && referenceRule.Color == "#ffffff" && referenceRule.Background == "#ff0000", "Differential background and theme font colors resolved");
        Check(referenceRule.MatchedCells!.SequenceEqual(["A1"]), "Percentage thresholds with absolute references match only outside interval");
        var changed = XlsxEditor.SetCells(reference, "1", [new("D1", "300", "number")]);
        Check(XlsxWorkbookExtractor.Extract(changed).Sheets[0].ConditionalRules.Single().MatchedCells!.Count == 0, "Referenced threshold updates on inspection");
        var differentialStyles = reference.ReadXml("xl/styles.xml");
        differentialStyles.Descendants(S + "patternFill").Last().SetAttributeValue("patternType", "solid");
        Check(XlsxWorkbookExtractor.Extract(reference.WithXml("xl/styles.xml", differentialStyles)).Sheets[0].ConditionalRules.Single().Background == "#ff0000", "Solid differential bgColor supported");
        var bytes = WorkbookApplication.Edit(XlsxWriter.Write(source), new("add-conditional-rule", "1", Range: "A2:A1", Operator: "greaterThan", Value: 10));
        var document = XlsxReader.Read(bytes);
        var sheet = XlsxWorkbookExtractor.Extract(document).Sheets[0];
        var rule = sheet.ConditionalRules.Single();
        Check(rule is { Range: "A1:A2", Value: 10, Color: "#9c0006", Background: "#ffc7ce", Supported: true, StopIfTrue: true }, "Native rule survives round trip");
        Check(sheet.Cells.Select(c => c.Value).SequenceEqual(["12", "5"]) && sheet.Cells.All(c => c.Style == 0), "Conditional rules do not alter cell data or base styles");
        document = XlsxConditionalFormatEditor.Edit(document, "1", "A1:A10", "between", -2.5, 3.5, "#000000", "#ffff00");
        Check(XlsxWorkbookExtractor.Extract(document).Sheets[0].ConditionalRules.Last() is { Priority: 2, SecondValue: 3.5 }, "Interval bounds and priorities preserved");
        using (var package = SpreadsheetDocument.Open(new MemoryStream(XlsxWriter.Write(document)), false))
        {
            var errors = new OpenXmlValidator().Validate(package).ToArray();
            Check(errors.Length == 0, "Conditional formatting schema: " + string.Join("; ", errors.Select(e => e.Description)));
        }
        document = XlsxStructureEditor.Insert(document, "1", true, 1);
        Check(XlsxWorkbookExtractor.Extract(document).Sheets[0].ConditionalRules.First().Range == "A2:A3", "Insertion shifts conditional ranges");
        bytes = WorkbookApplication.Edit(XlsxWriter.Write(document), new("remove-conditional-rule", "1", RuleId: 0));
        Check(WorkbookApplication.Inspect(bytes).Sheets[0].ConditionalRules.Single().Operator == "between", "Remove selected rule preserves other rules");
        Check(XlsxWorkbookExtractor.Extract(source).Sheets[0].ConditionalRules.Count == 0, "Input remains unchanged");
        try
        {
            XlsxConditionalFormatEditor.Edit(document, "1", "A1", "between", 5, 2, "#ffffff", "#000000");
            throw new Exception("Invalid interval accepted");
        }
        catch (ArgumentException)
        {
            checks++;
        }
        return checks;
    }
}
