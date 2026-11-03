using XlsxObjectKit.Application;
using XlsxObjectKit.Calculation;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Readers;
using XlsxObjectKit.Writers;

internal static class CalculationTests
{
    public static int Run()
    {
        var count = 0;
        void Check(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Calculation: " + message);
            }
            count++;
        }
        var original = XlsxEditor.Create().WithPart("customXml/keep.xml", "<keep/>"u8.ToArray());
        original = XlsxEditor.Worksheet(original, "add-sheet", "1", "Other sheet");
        original = XlsxEditor.SetCells(original, "2", [new("B1", "7", "number")]);
        original = XlsxEditor.SetCells(original, "1", [new("A1", "10", "number"), new("A2", "20", "number"), new("A3", "ignored"), new("A4", "true", "boolean")]);
        var examples = new (string Formula, string Value, string Type)[]
        {
            ("1+2*3", "7", "n"),
            ("(1+2)*3", "9", "n"),
            ("-2^2", "4", "n"),
            ("2^3^2", "64", "n"),
            ("50%*200", "100", "n"),
            ("$A$1+A$2", "30", "n"),
            ("'Other sheet'!$B$1*2", "14", "n"),
            ("SUM(A1:A5)", "30", "n"),
            ("AVERAGE(A1:A5)", "15", "n"),
            ("MIN(A1:A5)", "10", "n"),
            ("MAX(A1:A5)", "20", "n"),
            ("COUNT(A1:A5)", "2", "n"),
            ("COUNTA(A1:A5)", "4", "n"),
            ("IF(A1>5,\"yes\",1/0)", "yes", "str"),
            ("IFERROR(1/0,99)", "99", "n"),
            ("IF(TRUE,\"\",0)", "", "str"),
            ("1/0", "#DIV/0!", "e"),
            ("SQRT(-1)", "#NUM!", "e"),
            ("\"not numeric\"+1", "#VALUE!", "e"),
            ("Missing!A1", "#REF!", "e"),
            ("AND(A1>0,A2>0)", "1", "b"),
            ("OR(FALSE,NOT(FALSE))", "1", "b"),
            ("ROUND(2.555,2)", "2.56", "n"),
            ("MOD(-3,2)", "1", "n"),
            ("ABS(-3)+INT(-1.2)", "1", "n"),
            ("LEFT(\"abcdef\",2)&RIGHT(\"abcdef\",2)", "abef", "str"),
            ("MID(\"abcdef\",2,3)", "bcd", "str"),
            ("LEN(TRIM(\"  a   b  \"))", "3", "n"),
            ("UPPER(CONCAT(\"a\",LOWER(\"B\")))", "AB", "str"),
            ("\"abc\"=\"ABC\"", "1", "b"),
            ("1=\"1\"", "0", "b"),
            ("A99", "0", "n"),
            ("SUM(\"2\",TRUE,3)", "6", "n"),
            ("IF(FALSE,UNSUPPORTED(1),5)", "5", "n")
        };
        var edits = examples.Select((item, index) => new XlsxCellEdit("C" + (index + 1), item.Formula, "formula")).ToArray();
        var document = XlsxEditor.SetCells(original, "1", edits);
        var calculated = XlsxCalculator.Calculate(document);
        var sheet = XlsxWorkbookExtractor.Extract(calculated.Document).Sheets[0];
        for (var i = 0; i < examples.Length; i++)
        {
            var cell = sheet.Cells.Single(cell => cell.Address == "C" + (i + 1));
            Check(cell.HasCachedValue && cell.Value == examples[i].Value && cell.Type == examples[i].Type,
                examples[i].Formula + " expected " + examples[i].Value + " received " + cell.Value + " (" + cell.Type + ")");
        }
        Check(calculated.Unresolved == 0, string.Join("; ", calculated.Warnings));
        using (var package = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(new MemoryStream(XlsxWriter.Write(calculated.Document)), false))
        {
            var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(package).ToArray();
            Check(errors.Length == 0, "Calculated workbook schema: " + string.Join("; ", errors.Select(error => error.Description)));
        }
        Check(calculated.Document.GetPart("customXml/keep.xml").SequenceEqual(document.GetPart("customXml/keep.xml")), "Preserve unrelated package parts");
        Check(XlsxWorkbookExtractor.Extract(document).Sheets[0].Cells.Where(cell => cell.Formula is not null).All(cell => !cell.HasCachedValue), "Source remains immutable");
        var dependency = XlsxEditor.SetCells(original, "1", [new("B1", "=B2*2", "formula"), new("B2", "=A1+1", "formula")]);
        var updated = WorkbookApplication.Edit(XlsxWriter.Write(dependency), new("cells", "1", [new("A1", "100", "number")]));
        Check(WorkbookApplication.Inspect(updated).Sheets[0].Cells.Single(cell => cell.Address == "B1").Value == "202", "Recalculate dependencies after editing");
        Check(XlsxWorkbookExtractor.Extract(XlsxReader.Read(updated)).Sheets[0].Cells.Single(cell => cell.Address == "B1").Formula == "B2*2", "Formula preserved in saved package");
        var problematic = XlsxEditor.SetCells(original, "1", [new("B1", "=B2", "formula"), new("B2", "=B1", "formula"), new("B3", "=VLOOKUP(1,A1:A2,1)", "formula"), new("B4", "=IFERROR(B3,0)", "formula"), new("B5", "=SUM(A1:XFD1048576)", "formula"), new("B6", "=1+", "formula")]);
        var unresolved = XlsxCalculator.Calculate(problematic);
        Check(unresolved.Unresolved == 6 && unresolved.Warnings.Any(message => message.Contains("Circular")), "Cycles and unsupported formulas remain unresolved");
        Check(XlsxWorkbookExtractor.Extract(unresolved.Document).Sheets[0].Cells.Where(cell => cell.Formula is not null).All(cell => !cell.HasCachedValue), "No invented cached results; IFERROR does not mask unsupported calculations");
        return count;
    }
}
