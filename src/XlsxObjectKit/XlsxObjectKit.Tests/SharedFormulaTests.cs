using System.Xml.Linq;
using XlsxObjectKit.Application;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Writers;
using static XlsxObjectKit.Models.XlsxNamespaces;

internal static class SharedFormulaTests
{
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
        var doc = XlsxEditor.SetCells(XlsxEditor.Create(), "1", [new("F1", "47484", "number"), new("J1", "flag"), new("J2", "hold"), new("J3", "ready")]);
        var xml = doc.ReadXml("xl/worksheets/sheet1.xml");
        var data = xml.Root!.Element(S + "sheetData")!;
        for (var row = 1; row <= 8; row++)
        {
            var element = data.Elements(S + "row").FirstOrDefault(r => (int?)r.Attribute("r") == row);
            if (element is null)
            {
                element = new XElement(S + "row", new XAttribute("r", row));
                data.Add(element);
            }
            if (row > 1)
            {
                var formula = new XElement(S + "f", new XAttribute("t", "shared"), new XAttribute("si", 1));
                if (row == 2)
                {
                    formula.SetAttributeValue("ref", "F2:F8");
                    formula.Value = "F1+1";
                }
                element.Add(new XElement(S + "c", new XAttribute("r", "F" + row), formula));
            }
            var text = new XElement(S + "f", new XAttribute("t", "shared"), new XAttribute("si", 0));
            if (row == 1)
            {
                text.SetAttributeValue("ref", "G1:G8");
                text.Value = "TEXT(F1,\"dddd\")";
            }
            element.Add(new XElement(S + "c", new XAttribute("r", "G" + row), text));
        }
        doc = doc.WithXml("xl/worksheets/sheet1.xml", xml);
        var inspected = XlsxWorkbookExtractor.Extract(doc).Sheets[0];
        Check(inspected.Cells.Single(c => c.Address == "F8").Formula == "F7+1", "Shared follower has translated formula");
        Check(inspected.Cells.Single(c => c.Address == "G8").Formula == "TEXT(F8,\"dddd\")", "Shared formula translates nested references");
        var result = WorkbookApplication.EditWithDiagnostics(XlsxWriter.Write(doc), new("recalculate", "1"));
        Check(result.Warnings.Count == 0, "Shared dependency chains calculate without warnings");
        var calculated = WorkbookApplication.Inspect(result.Bytes).Sheets[0];
        Check(calculated.Cells.Single(c => c.Address == "F8").Value == "47491", "Shared chained additions calculate");
        var day = DateTime.FromOADate(47491).ToString("dddd", System.Globalization.CultureInfo.InvariantCulture);
        Check(calculated.Cells.Single(c => c.Address == "G8").Value == day, "TEXT uses calculated predecessor");
        var edited = WorkbookApplication.EditWithDiagnostics(XlsxWriter.Write(doc), new("cells", "1", Cells: [new("F4", "48000", "number")]));
        var changed = WorkbookApplication.Inspect(edited.Bytes).Sheets[0];
        Check(changed.Cells.Single(c => c.Address == "F8").Value == "48004", "Editing a shared follower updates downstream dependencies");
        Check(changed.Cells.Single(c => c.Address == "F3").Value == "47486", "Editing a follower preserves preceding shared formulas");
        doc = XlsxConditionalFormatEditor.Edit(doc, "1", "J1:J3", "equal", 1, 0, "#000000", "#ffffff");
        xml = doc.ReadXml("xl/worksheets/sheet1.xml");
        var rule = xml.Descendants(S + "cfRule").Single();
        rule.SetAttributeValue("type", "expression");
        rule.Attribute("operator")!.Remove();
        rule.Element(S + "formula")!.Value = "OR(J1=\"flag\",J1=\"hold\")";
        var styles = doc.ReadXml("xl/styles.xml");
        styles.Descendants(S + "dxf").Single().ReplaceNodes(new XElement(S + "font", new XElement(S + "b"), new XElement(S + "i", new XAttribute("val", 0))));
        doc = doc.WithXml("xl/worksheets/sheet1.xml", xml).WithXml("xl/styles.xml", styles);
        var conditional = XlsxWorkbookExtractor.Extract(doc).Sheets[0].ConditionalRules.Single();
        Check(conditional.Supported && conditional.Bold == true && conditional.Italic == false, "Expression differential bold and explicit italic false are supported");
        Check(conditional.MatchedCells!.SequenceEqual(["J1", "J2"]), "Relative expression comparisons match only appropriate text cells");
        return checks;
    }
}
