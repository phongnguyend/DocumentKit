using System.Text.RegularExpressions;
using XlsxObjectKit.Models;

namespace XlsxObjectKit.Calculation;

internal sealed partial class FormulaEngine
{
    private FormulaValue Function(FormulaNode node, XlsxSheet sheet)
    {
        var name = node.Text.StartsWith("_XLFN.", StringComparison.Ordinal) ? node.Text[6..] : node.Text;
        var count = node.Children.Count;
        FormulaValue Value(int index) => Evaluate(node.Children[index], sheet);
        if (name == "TEXT" && count == 2)
        {
            var number = Value(0).AsNumber();
            var format = Value(1);
            if (number.Failed || format.Failed)
            {
                return number.Failed ? number : format;
            }
            if (format.Text is "dddd" or "ddd")
            {
                if (number.Number < 0 || number.Number > 2958465)
                {
                    return FormulaValue.Error("#VALUE!");
                }
                var days = (int)Math.Floor(number.Number);
                var day = workbook.Date1904 ? (days + 5) % 7 : (days + (days < 60 ? 0 : 6)) % 7;
                var culture = System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat;
                return new("str", format.Text == "dddd" ? culture.DayNames[day] : culture.AbbreviatedDayNames[day]);
            }
            return FormulaValue.Pending("TEXT currently supports ddd and dddd date formats.");
        }
        if (name is "TRUE" or "FALSE")
        {
            return count == 0 ? FormulaValue.Boolean(name == "TRUE") : FormulaValue.Error("#VALUE!");
        }
        if (name == "IF")
        {
            if (count is < 2 or > 3)
            {
                return FormulaValue.Error("#VALUE!");
            }
            var condition = Value(0).AsBoolean();
            return condition.Failed ? condition : condition.Number != 0 ? Value(1) : count == 3 ? Value(2) : FormulaValue.Boolean(false);
        }
        if (name == "IFERROR")
        {
            if (count != 2)
            {
                return FormulaValue.Error("#VALUE!");
            }
            var first = Value(0);
            return first.Kind == "e" ? Value(1) : first;
        }
        if (name is "SUM" or "AVERAGE" or "MIN" or "MAX" or "COUNT" or "COUNTA" or "AND" or "OR" or "CONCAT" or "CONCATENATE")
        {
            if (count == 0)
            {
                return FormulaValue.Error("#VALUE!");
            }
            var numbers = new List<double>();
            var text = new System.Text.StringBuilder();
            var nonempty = 0;
            foreach (var (value, referenced) in Arguments(node, sheet))
            {
                if (value.Kind == "pending")
                {
                    return value;
                }
                if (value.Kind != "blank")
                {
                    nonempty++;
                }
                if (name == "COUNTA")
                {
                    continue;
                }
                if (value.Kind == "e")
                {
                    if (name == "COUNT")
                    {
                        continue;
                    }
                    return value;
                }
                if (name is "CONCAT" or "CONCATENATE")
                {
                    text.Append(value.Text);
                    if (text.Length > 32767)
                    {
                        return FormulaValue.Error("#VALUE!");
                    }
                    continue;
                }
                if (name is "AND" or "OR")
                {
                    if (referenced && value.Kind is "str" or "blank")
                    {
                        continue;
                    }
                    var logical = value.AsBoolean();
                    if (logical.Failed)
                    {
                        return logical;
                    }
                    numbers.Add(logical.Number);
                    continue;
                }
                if (referenced && value.Kind != "n")
                {
                    continue;
                }
                var numeric = value.AsNumber();
                if (numeric.Failed)
                {
                    if (name == "COUNT")
                    {
                        continue;
                    }
                    return numeric;
                }
                numbers.Add(numeric.Number);
            }
            return name switch
            {
                "SUM" => FormulaValue.Numeric(numbers.Sum()),
                "COUNT" => FormulaValue.Numeric(numbers.Count),
                "COUNTA" => FormulaValue.Numeric(nonempty),
                "MIN" => FormulaValue.Numeric(numbers.Count == 0 ? 0 : numbers.Min()),
                "MAX" => FormulaValue.Numeric(numbers.Count == 0 ? 0 : numbers.Max()),
                "AVERAGE" => numbers.Count == 0 ? FormulaValue.Error("#DIV/0!") : FormulaValue.Numeric(numbers.Sum() / numbers.Count),
                "AND" => numbers.Count == 0 ? FormulaValue.Error("#VALUE!") : FormulaValue.Boolean(numbers.All(value => value != 0)),
                "OR" => numbers.Count == 0 ? FormulaValue.Error("#VALUE!") : FormulaValue.Boolean(numbers.Any(value => value != 0)),
                _ => Text(text.ToString())
            };
        }
        if (name is "NOT" or "ABS" or "INT" or "SQRT" or "LEN" or "LOWER" or "UPPER" or "TRIM")
        {
            if (count != 1)
            {
                return FormulaValue.Error("#VALUE!");
            }
            var value = Value(0);
            if (value.Failed)
            {
                return value;
            }
            if (name is "LEN" or "LOWER" or "UPPER" or "TRIM")
            {
                return name switch
                {
                    "LEN" => FormulaValue.Numeric(value.Text.Length),
                    "LOWER" => Text(value.Text.ToLowerInvariant()),
                    "UPPER" => Text(value.Text.ToUpperInvariant()),
                    _ => Text(Regex.Replace(value.Text.Trim(' '), " +", " "))
                };
            }
            value = name == "NOT" ? value.AsBoolean() : value.AsNumber();
            if (value.Failed)
            {
                return value;
            }
            return name switch
            {
                "NOT" => FormulaValue.Boolean(value.Number == 0),
                "ABS" => FormulaValue.Numeric(Math.Abs(value.Number)),
                "INT" => FormulaValue.Numeric(Math.Floor(value.Number)),
                _ => FormulaValue.Numeric(Math.Sqrt(value.Number))
            };
        }
        if (name is "ROUND" or "POWER" or "MOD")
        {
            if (count != 2)
            {
                return FormulaValue.Error("#VALUE!");
            }
            var a = Value(0).AsNumber();
            var b = Value(1).AsNumber();
            if (a.Failed || b.Failed)
            {
                return a.Kind == "pending" ? a : b.Kind == "pending" ? b : a.Failed ? a : b;
            }
            if (name == "POWER")
            {
                return FormulaValue.Numeric(Math.Pow(a.Number, b.Number));
            }
            if (name == "MOD")
            {
                return b.Number == 0 ? FormulaValue.Error("#DIV/0!") : FormulaValue.Numeric(a.Number - b.Number * Math.Floor(a.Number / b.Number));
            }
            if (Math.Abs(b.Number) > 308)
            {
                return FormulaValue.Error("#NUM!");
            }
            var factor = Math.Pow(10, Math.Truncate(b.Number));
            return FormulaValue.Numeric(Math.Round(a.Number * factor, MidpointRounding.AwayFromZero) / factor);
        }
        if (name is "LEFT" or "RIGHT" or "MID")
        {
            if (name == "MID" ? count != 3 : count is < 1 or > 2)
            {
                return FormulaValue.Error("#VALUE!");
            }
            var text = Value(0);
            var length = count == 1 ? FormulaValue.Numeric(1) : Value(name == "MID" ? 2 : 1).AsNumber();
            var start = name == "MID" ? Value(1).AsNumber() : FormulaValue.Numeric(1);
            if (text.Failed || length.Failed || start.Failed)
            {
                return text.Kind == "pending" ? text : length.Kind == "pending" ? length : start.Kind == "pending" ? start : text.Failed ? text : length.Failed ? length : start;
            }
            if (length.Number < 0 || start.Number < 1)
            {
                return FormulaValue.Error("#VALUE!");
            }
            var offset = name == "MID" ? (int)Math.Min(text.Text.Length, start.Number - 1) : 0;
            var take = (int)Math.Min(text.Text.Length - offset, length.Number);
            return Text(name == "RIGHT" ? text.Text[^take..] : text.Text.Substring(offset, take));
        }
        return FormulaValue.Pending($"Function {name} is not supported by the native engine yet.");
    }
}
