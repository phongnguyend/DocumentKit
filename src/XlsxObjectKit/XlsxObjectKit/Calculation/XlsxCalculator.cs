using System.Globalization;
using System.Xml.Linq;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Calculation;

public sealed record XlsxCalculationResult(XlsxObjectDocument Document, int Calculated, int Unresolved, IReadOnlyList<string> Warnings);

public static class XlsxCalculator
{
    public static XlsxCalculationResult Calculate(XlsxObjectDocument document) => new FormulaEngine(document).Run();
}

internal sealed partial class FormulaEngine
{
    private readonly XlsxObjectDocument source;
    private readonly XlsxWorkbook workbook;
    private readonly Dictionary<string, Dictionary<string, XlsxCell>> cells;
    private readonly Dictionary<(string Sheet, string Cell), FormulaValue> values = [];
    private readonly HashSet<(string Sheet, string Cell)> visiting = [];
    private readonly HashSet<(string Sheet, string Cell)> special = [];
    private int steps;
    private int depth;

    public FormulaEngine(XlsxObjectDocument document)
    {
        source = document;
        workbook = XlsxWorkbookExtractor.Extract(document);
        cells = workbook.Sheets.ToDictionary(sheet => sheet.Id, sheet => sheet.Cells.ToDictionary(cell => cell.Address, StringComparer.OrdinalIgnoreCase));
    }

    public XlsxCalculationResult Run()
    {
        if (source.PartNames.Any(part => part.StartsWith("_xmlsignatures/", StringComparison.Ordinal)))
        {
            throw new NotSupportedException("Signed workbooks cannot be recalculated without invalidating their signatures.");
        }
        var sheets = workbook.Sheets.ToDictionary(sheet => sheet.Id, sheet => source.ReadXml(sheet.Part));
        foreach (var sheet in sheets)
        {
            foreach (var cell in sheet.Value.Descendants(S + "c").Where(cell => cell.Element(S + "f") is { } f && (string?)f.Attribute("t") is not (null or "normal")))
            {
                special.Add((sheet.Key, (string)cell.Attribute("r")!));
            }
        }
        var count = workbook.Sheets.Sum(sheet => sheet.Cells.Count(cell => cell.Formula is not null));
        var warnings = new List<string>();
        var calculated = 0;
        var unresolved = 0;
        var result = source;
        foreach (var sheet in workbook.Sheets)
        {
            var xml = sheets[sheet.Id];
            var changed = false;
            foreach (var element in xml.Descendants(S + "c").Where(cell => cell.Element(S + "f") is not null))
            {
                changed = true;
                var address = (string)element.Attribute("r")!;
                var value = count > 10000 ? FormulaValue.Pending("Workbook exceeds the 10,000-formula calculation limit.")
                    : special.Count > 0 ? FormulaValue.Pending("Workbooks containing shared/array/data-table formulas are not supported for calculation yet.") : Cell(sheet, address);
                element.Elements(S + "v").Remove();
                element.Elements(S + "is").Remove();
                element.Attribute("t")?.Remove();
                if (value.Kind == "pending")
                {
                    unresolved++;
                    if (warnings.Count < 100)
                    {
                        warnings.Add($"{sheet.Name}!{address}: {value.Text}");
                    }
                    continue;
                }
                calculated++;
                var type = value.Kind == "blank" ? "n" : value.Kind;
                if (type != "n")
                {
                    element.SetAttributeValue("t", type);
                }
                element.Element(S + "f")!.AddAfterSelf(new XElement(S + "v", value.Kind == "b" ? value.Number == 0 ? "0" : "1" : value.Kind == "blank" ? "0" : value.Text));
            }
            if (changed)
            {
                result = result.WithXml(sheet.Part, xml);
            }
        }
        return new(result, calculated, unresolved, warnings);
    }

    private FormulaValue Cell(XlsxSheet sheet, string address)
    {
        if (++steps > 200000)
        {
            return FormulaValue.Pending("Calculation work limit reached.");
        }
        var key = (sheet.Id, address);
        if (values.TryGetValue(key, out var cached))
        {
            return cached;
        }
        if (!cells[sheet.Id].TryGetValue(address, out var cell))
        {
            return new("blank", "");
        }
        if (cell.Formula is null)
        {
            return cell.Type switch
            {
                "b" => FormulaValue.Boolean(cell.Value == "1"),
                "e" => FormulaValue.Error(cell.Value),
                "n" when cell.Value == "" => new("blank", ""),
                "n" => double.TryParse(cell.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? FormulaValue.Numeric(number) : FormulaValue.Error("#VALUE!"),
                "d" => FormulaValue.Pending("ISO date cells are not supported by calculation yet."),
                _ => new("str", cell.Value)
            };
        }
        if (visiting.Contains(key))
        {
            return FormulaValue.Pending("Circular reference detected.");
        }
        if (visiting.Count >= 128 || special.Contains(key))
        {
            return FormulaValue.Pending("Dependency depth limit reached, or shared/array/data-table formula is unsupported.");
        }
        visiting.Add(key);
        FormulaValue value;
        try
        {
            value = Evaluate(new FormulaParser(cell.Formula).Parse(), sheet);
        }
        catch (Exception exception) when (exception is NotSupportedException or ArgumentException or System.Text.RegularExpressions.RegexMatchTimeoutException)
        {
            value = FormulaValue.Pending(exception.Message);
        }
        finally
        {
            visiting.Remove(key);
        }
        values[key] = value;
        return value;
    }

    private FormulaValue Reference(FormulaNode node, XlsxSheet sheet)
    {
        try
        {
            CellAddress.Parse(node.Text);
        }
        catch (ArgumentException)
        {
            return FormulaValue.Pending("Named ranges, whole-row/column references and this reference syntax are not supported yet.");
        }
        var target = node.Sheet is null ? sheet : workbook.Sheets.FirstOrDefault(item => item.Name.Equals(node.Sheet, StringComparison.OrdinalIgnoreCase));
        return target is null ? FormulaValue.Error("#REF!") : Cell(target, node.Text);
    }

    private FormulaValue Evaluate(FormulaNode node, XlsxSheet sheet)
    {
        if (++steps > 200000 || depth >= 128)
        {
            return FormulaValue.Pending("Calculation work or expression depth limit reached.");
        }
        depth++;
        try
        {
            return node.Kind switch
            {
                "number" => FormulaValue.Numeric(double.Parse(node.Text, CultureInfo.InvariantCulture)),
                "text" => new("str", node.Text),
                "blank" => new("blank", ""),
                "boolean" => FormulaValue.Boolean(node.Text == "TRUE"),
                "error" => FormulaValue.Error(node.Text),
                "ref" => Reference(node, sheet),
                "unary" => Unary(node, sheet),
                "binary" => Binary(node, sheet),
                "call" => Function(node, sheet),
                _ => FormulaValue.Pending("Ranges must be used in a supported aggregate function.")
            };
        }
        finally
        {
            depth--;
        }
    }

    private FormulaValue Unary(FormulaNode node, XlsxSheet sheet)
    {
        var value = Evaluate(node.Children[0], sheet).AsNumber();
        return value.Failed ? value : FormulaValue.Numeric(node.Text == "-" ? -value.Number : node.Text == "%" ? value.Number / 100 : value.Number);
    }

    private FormulaValue Binary(FormulaNode node, XlsxSheet sheet)
    {
        var a = Evaluate(node.Children[0], sheet);
        var b = Evaluate(node.Children[1], sheet);
        if (a.Failed || b.Failed)
        {
            return a.Kind == "pending" ? a : b.Kind == "pending" ? b : a.Failed ? a : b;
        }
        if (node.Text == "&")
        {
            return Text(a.Text + b.Text);
        }
        if (node.Text is "=" or "<>" or "<" or ">" or "<=" or ">=")
        {
            if (a.Kind == "blank")
            {
                a = b.Kind == "str" ? new("str", "") : FormulaValue.Numeric(0);
            }
            if (b.Kind == "blank")
            {
                b = a.Kind == "str" ? new("str", "") : FormulaValue.Numeric(0);
            }
            int Rank(FormulaValue value) => value.Kind == "str" ? 1 : value.Kind == "b" ? 2 : 0;
            var comparison = Rank(a) != Rank(b) ? Rank(a).CompareTo(Rank(b)) : a.Kind == "str" ? StringComparer.OrdinalIgnoreCase.Compare(a.Text, b.Text) : a.Number.CompareTo(b.Number);
            return FormulaValue.Boolean(node.Text switch { "=" => comparison == 0, "<>" => comparison != 0, "<" => comparison < 0, ">" => comparison > 0, "<=" => comparison <= 0, _ => comparison >= 0 });
        }
        a = a.AsNumber();
        b = b.AsNumber();
        if (a.Failed || b.Failed)
        {
            return a.Failed ? a : b;
        }
        return node.Text switch
        {
            "+" => FormulaValue.Numeric(a.Number + b.Number),
            "-" => FormulaValue.Numeric(a.Number - b.Number),
            "*" => FormulaValue.Numeric(a.Number * b.Number),
            "/" => b.Number == 0 ? FormulaValue.Error("#DIV/0!") : FormulaValue.Numeric(a.Number / b.Number),
            "^" => FormulaValue.Numeric(Math.Pow(a.Number, b.Number)),
            _ => FormulaValue.Pending("Unsupported operator.")
        };
    }

    private static FormulaValue Text(string value) => value.Length > 32767 ? FormulaValue.Error("#VALUE!") : new("str", value);

    private IEnumerable<(FormulaValue Value, bool Referenced)> Arguments(FormulaNode node, XlsxSheet sheet)
    {
        foreach (var argument in node.Children)
        {
            if (argument.Kind != "range")
            {
                yield return (Evaluate(argument, sheet), argument.Kind == "ref");
                continue;
            }
            var start = CellAddress.Parse(argument.Children[0].Text);
            var end = CellAddress.Parse(argument.Children[1].Text);
            var minRow = Math.Min(start.Row, end.Row);
            var maxRow = Math.Max(start.Row, end.Row);
            var minColumn = Math.Min(start.Column, end.Column);
            var maxColumn = Math.Max(start.Column, end.Column);
            if ((long)(maxRow - minRow + 1) * (maxColumn - minColumn + 1) > 100000)
            {
                yield return (FormulaValue.Pending("Range exceeds 100,000 cells."), true);
                yield break;
            }
            for (var row = minRow; row <= maxRow; row++)
            {
                for (var column = minColumn; column <= maxColumn; column++)
                {
                    var value = Reference(argument.Children[0] with { Text = CellAddress.Format(row, column) }, sheet);
                    yield return (value, true);
                    if (steps > 200000)
                    {
                        yield break;
                    }
                }
            }
        }
    }
}
