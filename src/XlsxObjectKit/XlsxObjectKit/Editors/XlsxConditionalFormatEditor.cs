using System.Globalization;
using XlsxObjectKit.Calculation;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Editors;

public sealed record XlsxConditionalRule(int Id, string Range, string Operator, double? Value, double? SecondValue,
    string? Color, string? Background, bool StopIfTrue, int Priority, bool Supported)
{
    public bool? Bold { get; init; }

    public bool? Italic { get; init; }

    public IReadOnlyList<string>? MatchedCells { get; init; }
}

public static class XlsxConditionalFormatEditor
{
    private static readonly string[] Operators = ["greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual", "equal", "notEqual", "between", "notBetween"];

    public static IReadOnlyList<XlsxConditionalRule> Read(XlsxObjectDocument document, string sheetPart, IReadOnlyList<XlsxCell>? cells = null)
    {
        var stylePart = XlsxWorkbookExtractor.StylesPart(document);
        var dxfs = stylePart is null ? [] : document.ReadXml(stylePart).Root!.Element(S + "dxfs")?.Elements().ToArray() ?? [];
        var result = new List<XlsxConditionalRule>();
        var lookup = (cells ?? []).ToDictionary(c => c.Address, StringComparer.OrdinalIgnoreCase);
        double? Number(string? value) => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) && double.IsFinite(n) ? n : null;
        var relationships = document.ReadXml("xl/_rels/workbook.xml.rels").Root!.Elements();
        var theme = relationships.FirstOrDefault(r => ((string?)r.Attribute("Type"))?.EndsWith("/theme") == true);
        var colors = XlsxStyleReader.ColorReader(document, stylePart, theme is null ? null : Resolve("xl/workbook.xml", (string)theme.Attribute("Target")!));
        string? Color(XElement? element) => element is null ? null : colors(element, "") is { Length: > 0 } resolved ? resolved : null;
        foreach (var group in document.ReadXml(sheetPart).Root!.Elements(S + "conditionalFormatting"))
        {
            foreach (var rule in group.Elements(S + "cfRule"))
            {
                var dxf = dxfs.ElementAtOrDefault((int?)rule.Attribute("dxfId") ?? -1);
                var formulas = rule.Elements(S + "formula").Select(f => f.Value).ToArray();
                var op = (string?)rule.Attribute("operator") ?? (string?)rule.Attribute("type") ?? "unknown";
                var color = Color(dxf?.Element(S + "font")?.Element(S + "color"));
                var fill = dxf?.Element(S + "fill")?.Element(S + "patternFill");
                var background = fill is not null && (string?)fill.Attribute("patternType") is null or "solid" ? Color(fill.Element(S + "fgColor")) ?? Color(fill.Element(S + "bgColor")) : null;
                var first = Number(formulas.ElementAtOrDefault(0));
                var second = Number(formulas.ElementAtOrDefault(1));
                var range = (string?)group.Attribute("sqref") ?? "";
                var anchor = range.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.Split(':')[0] ?? "A1";
                var condition = (string?)rule.Attribute("type") == "expression" ? formulas.FirstOrDefault() : null;
                if ((string?)rule.Attribute("type") == "cellIs" && Operators.Contains(op) && formulas.Length == (op is "between" or "notBetween" ? 2 : 1))
                {
                    var a = "(" + formulas[0] + ")";
                    condition = op switch
                    {
                        "between" => $"AND({anchor}>={a},{anchor}<=({formulas[1]}))",
                        "notBetween" => $"OR({anchor}<{a},{anchor}>({formulas[1]}))",
                        _ => anchor + (op switch { "equal" => "=", "notEqual" => "<>", "lessThan" => "<", "lessThanOrEqual" => "<=", "greaterThan" => ">", _ => ">=" }) + a
                    };
                }
                var expression = condition is not null && ConditionalExpressions.Supports(condition);
                var supported = expression
                    && dxf is not null && dxf.Elements().All(e => e.Name == S + "font" || e.Name == S + "fill")
                    && (dxf.Element(S + "font") is not { } font || font.Elements().All(e => e.Name == S + "color" || e.Name == S + "b" || e.Name == S + "i") && (font.Element(S + "color") is null || color is not null))
                    && (dxf.Element(S + "fill") is null || background is not null);
                var matched = new List<string>();
                if (supported && expression && cells is not null)
                {
                    var areas = range.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(area => area.Replace("$", "").Split(':').Select(CellAddress.Parse).ToArray()).ToArray();
                    var origin = areas[0][0];
                    foreach (var cell in cells)
                    {
                        if ((string?)rule.Attribute("type") == "cellIs" && !formulas[0].TrimStart().StartsWith('"')
                            && (cell.Type != "n" || cell.Value.Length == 0))
                        {
                            continue;
                        }
                        var at = CellAddress.Parse(cell.Address);
                        if (areas.Any(area => at.Row >= area[0].Row && at.Row <= area[^1].Row && at.Column >= area[0].Column && at.Column <= area[^1].Column)
                            && ConditionalExpressions.Matches(condition!, at.Row - origin.Row, at.Column - origin.Column, lookup))
                        {
                            matched.Add(cell.Address);
                        }
                    }
                }
                result.Add(new(result.Count, (string?)group.Attribute("sqref") ?? "", op, first, second, color, background,
                    (bool?)rule.Attribute("stopIfTrue") ?? false, (int?)rule.Attribute("priority") ?? int.MaxValue, supported)
                {
                    Bold = dxf?.Element(S + "font")?.Element(S + "b") is { } bold ? (bool?)bold.Attribute("val") ?? true : null,
                    Italic = dxf?.Element(S + "font")?.Element(S + "i") is { } italic ? (bool?)italic.Attribute("val") ?? true : null,
                    MatchedCells = expression && supported && cells is not null ? matched : null
                });
            }
        }
        return result;
    }

    public static XlsxObjectDocument Edit(XlsxObjectDocument document, string sheetId, string? range, string? op,
        double value, double secondValue, string color, string background, int removeId = -1)
    {
        var sheet = XlsxWorkbookExtractor.Extract(document).Sheets.Single(s => s.Id == sheetId);
        XlsxEditor.Guard(document, sheet.Part);
        var xml = document.ReadXml(sheet.Part);
        if (removeId >= 0)
        {
            var rule = xml.Root!.Elements(S + "conditionalFormatting").Elements(S + "cfRule").ElementAtOrDefault(removeId)
                ?? throw new ArgumentException("Conditional formatting rule no longer exists.");
            var parent = rule.Parent!;
            rule.Remove();
            if (!parent.HasElements)
            {
                parent.Remove();
            }
            return document.WithXml(sheet.Part, xml);
        }
        if (!Operators.Contains(op) || !double.IsFinite(value) || !double.IsFinite(secondValue)
            || !Regex.IsMatch(color, "^#[0-9a-fA-F]{6}$") || !Regex.IsMatch(background, "^#[0-9a-fA-F]{6}$")
            || op is "between" or "notBetween" && secondValue < value)
        {
            throw new ArgumentException("Choose a supported comparison, finite numeric values, valid colors, and an ordered interval.");
        }
        var ends = (range ?? "").Split(':');
        if (ends.Length > 2)
        {
            throw new ArgumentException("Choose an A1:B2 range.");
        }
        var a = CellAddress.Parse(ends[0]);
        var b = CellAddress.Parse(ends[^1]);
        range = CellAddress.Format(Math.Min(a.Row, b.Row), Math.Min(a.Column, b.Column)) + ":" + CellAddress.Format(Math.Max(a.Row, b.Row), Math.Max(a.Column, b.Column));
        var part = XlsxWorkbookExtractor.StylesPart(document);
        XDocument styles;
        if (part is null)
        {
            part = "xl/styles.xml";
            styles = XDocument.Parse($"<styleSheet xmlns=\"{S}\"><fonts count=\"1\"><font/></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf/></cellStyleXfs><cellXfs count=\"1\"><xf fontId=\"0\" fillId=\"0\" borderId=\"0\" numFmtId=\"0\" xfId=\"0\"/></cellXfs></styleSheet>");
            var rels = document.ReadXml("xl/_rels/workbook.xml.rels");
            rels.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", "styles" + Guid.NewGuid().ToString("N")), new XAttribute("Type", R.NamespaceName + "/styles"), new XAttribute("Target", "styles.xml")));
            var types = document.ReadXml("[Content_Types].xml");
            types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")));
            document = document.WithXml("xl/_rels/workbook.xml.rels", rels).WithXml("[Content_Types].xml", types);
        }
        else
        {
            styles = document.ReadXml(part);
        }
        var dxfs = styles.Root!.Element(S + "dxfs");
        if (dxfs is null)
        {
            dxfs = new XElement(S + "dxfs");
            var following = styles.Root.Elements().FirstOrDefault(e => e.Name.LocalName is "tableStyles" or "colors" or "extLst");
            if (following is null)
            {
                styles.Root.Add(dxfs);
            }
            else
            {
                following.AddBeforeSelf(dxfs);
            }
        }
        var id = dxfs.Elements().Count();
        dxfs.Add(new XElement(S + "dxf", new XElement(S + "font", new XElement(S + "color", new XAttribute("rgb", "FF" + color[1..]))),
            new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "solid"), new XElement(S + "fgColor", new XAttribute("rgb", "FF" + background[1..]))))));
        dxfs.SetAttributeValue("count", id + 1);
        var priority = xml.Descendants(S + "cfRule").Select(r => (int?)r.Attribute("priority") ?? 0).DefaultIfEmpty(0).Max() + 1;
        var added = new XElement(S + "cfRule", new XAttribute("type", "cellIs"), new XAttribute("operator", op!), new XAttribute("dxfId", id),
            new XAttribute("priority", priority), new XAttribute("stopIfTrue", "1"), new XElement(S + "formula", value.ToString("R", CultureInfo.InvariantCulture)));
        if (op is "between" or "notBetween")
        {
            added.Add(new XElement(S + "formula", secondValue.ToString("R", CultureInfo.InvariantCulture)));
        }
        var group = new XElement(S + "conditionalFormatting", new XAttribute("sqref", range), added);
        var next = xml.Root!.Elements().FirstOrDefault(e => new[] { "dataConsolidate", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst" }.Contains(e.Name.LocalName));
        if (next is null)
        {
            xml.Root.Add(group);
        }
        else
        {
            next.AddBeforeSelf(group);
        }
        return document.WithXml(part, styles).WithXml(sheet.Part, xml);
    }
}
