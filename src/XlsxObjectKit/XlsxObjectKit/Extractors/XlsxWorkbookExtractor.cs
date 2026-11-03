using System.Globalization;
using System.Xml.Linq;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Extractors;

public static class XlsxWorkbookExtractor
{
    public static XlsxWorkbook Extract(XlsxObjectDocument document)
    {
        var workbook = document.ReadXml("xl/workbook.xml");
        var relations = document.ReadXml("xl/_rels/workbook.xml.rels").Root!.Elements().ToDictionary(item => (string)item.Attribute("Id")!);
        string? Part(string suffix)
        {
            var relation = relations.Values.FirstOrDefault(item => ((string?)item.Attribute("Type"))?.EndsWith("/" + suffix) == true);
            return relation is null ? null : Resolve("xl/workbook.xml", (string)relation.Attribute("Target")!);
        }
        var strings = Part("sharedStrings") is { } stringPart ? document.ReadXml(stringPart).Root!.Elements(S + "si")
            .Select(item => string.Concat(item.Descendants(S + "t").Where(text => !text.Ancestors(S + "rPh").Any()).Select(text => text.Value))).ToArray() : [];
        var styles = XlsxStyleReader.Read(document, Part("styles"), Part("theme"));
        var sheets = new List<XlsxSheet>();
        var total = 0;
        foreach (var sheet in workbook.Descendants(S + "sheet"))
        {
            if (!relations.TryGetValue((string)sheet.Attribute(R + "id")!, out var relation) || (string?)relation.Attribute("TargetMode") == "External")
            {
                throw new InvalidDataException("Missing or external worksheet relationship.");
            }
            var part = Resolve("xl/workbook.xml", (string)relation.Attribute("Target")!);
            var xml = document.ReadXml(part);
            if (xml.Root?.Name != S + "worksheet")
            {
                throw new NotSupportedException("Only worksheets are supported; this workbook contains a chart sheet or another sheet type.");
            }
            var cells = new List<XlsxCell>();
            foreach (var cell in xml.Descendants(S + "sheetData").Elements(S + "row").Elements(S + "c"))
            {
                if (++total > 200000)
                {
                    throw new NotSupportedException("The playground supports up to 200,000 stored cells per workbook.");
                }
                var address = ((string?)cell.Attribute("r") ?? throw new InvalidDataException("Cell address missing.")).ToUpperInvariant();
                CellAddress.Parse(address);
                var type = (string?)cell.Attribute("t") ?? "n";
                var value = cell.Element(S + "v")?.Value ?? "";
                if (type == "s")
                {
                    if (!int.TryParse(value, out var index) || index < 0 || index >= strings.Length)
                    {
                        throw new InvalidDataException("Invalid shared-string index.");
                    }
                    value = strings[index];
                }
                else if (type == "inlineStr")
                {
                    value = string.Concat(cell.Descendants(S + "t").Where(text => !text.Ancestors(S + "rPh").Any()).Select(text => text.Value));
                }
                cells.Add(new(address, value, type, cell.Element(S + "f")?.Value, (int?)cell.Attribute("s") ?? -1)
                {
                    HasCachedValue = cell.Element(S + "v") is not null
                });
            }
            var heights = xml.Descendants(S + "row").Where(row => row.Attribute("ht") is not null)
                .ToDictionary(row => (int)row.Attribute("r")!, row => Number(row, "ht", 20));
            var widths = new Dictionary<int, double>();
            var columnStyles = new Dictionary<int, int>();
            foreach (var col in xml.Descendants(S + "cols").Elements(S + "col"))
            {
                for (var i = Math.Max(1, (int?)col.Attribute("min") ?? 1); i <= Math.Min(16384, (int?)col.Attribute("max") ?? 1); i++)
                {
                    widths[i] = Number(col, "width", 8.43);
                    if (col.Attribute("style") is not null)
                    {
                        columnStyles[i] = (int)col.Attribute("style")!;
                    }
                }
            }
            sheets.Add(new((string)sheet.Attribute("sheetId")!, (string)sheet.Attribute("name")!, part,
                (string?)sheet.Attribute("state") is "hidden" or "veryHidden", cells,
                xml.Descendants(S + "mergeCell").Select(item => (string)item.Attribute("ref")!).ToArray(), heights, widths)
            {
                RowStyles = xml.Descendants(S + "row").Where(row => row.Attribute("s") is not null && (string?)row.Attribute("customFormat") is "1" or "true")
                    .ToDictionary(row => (int)row.Attribute("r")!, row => (int)row.Attribute("s")!),
                ColumnStyles = columnStyles,
                DefaultRowHeight = Number(xml.Root.Element(S + "sheetFormatPr") ?? new XElement("empty"), "defaultRowHeight", 15),
                DefaultColumnWidth = Number(xml.Root.Element(S + "sheetFormatPr") ?? new XElement("empty"), "defaultColWidth", 8.43)
            });
        }
        for (var i = 0; i < sheets.Count; i++)
        {
            var sheet = sheets[i];
            sheets[i] = sheet with { Cells = sheet.Cells.Select(cell =>
            {
                var location = CellAddress.Parse(cell.Address);
                return cell.Style >= 0 ? cell : cell with { Style = sheet.RowStyles.GetValueOrDefault(location.Row, sheet.ColumnStyles.GetValueOrDefault(location.Column, 0)) };
            }).ToArray() };
        }
        if (sheets.Count == 0 || sheets.Count > 256 || sheets.All(sheet => sheet.Hidden))
        {
            throw new NotSupportedException("Choose a workbook with at least one visible sheet and 1–256 worksheets.");
        }
        var pending = sheets.Sum(sheet => sheet.Cells.Count(cell => cell.Formula is not null && !cell.HasCachedValue));
        return new(sheets, styles, [$"Native calculation supports common formulas. {pending} formula(s) have no cached result. Recalculate to refresh uploaded results; unsupported formulas remain unresolved.",
            "Charts, drawings, conditional formatting, filters and advanced Excel features are preserved in the file but are not rendered. Pattern/gradient fills, rich-text runs and exact Excel font metrics may differ."])
        {
            Date1904 = (string?)workbook.Root!.Element(S + "workbookPr")?.Attribute("date1904") is "1" or "true"
        };
    }

    public static string? StylesPart(XlsxObjectDocument document) => document.ReadXml("xl/_rels/workbook.xml.rels").Root!.Elements()
        .Where(item => ((string?)item.Attribute("Type"))?.EndsWith("/styles") == true)
        .Select(item => Resolve("xl/workbook.xml", (string)item.Attribute("Target")!)).FirstOrDefault();

    private static double Number(XElement item, string name, double fallback) => double.TryParse((string?)item.Attribute(name), NumberStyles.Float,
        CultureInfo.InvariantCulture, out var value) && double.IsFinite(value) ? value : fallback;
}
