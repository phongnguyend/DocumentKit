using System.Globalization;
using XlsxObjectKit.Models;

namespace XlsxObjectKit.Calculation;

internal static class ConditionalExpressions
{
    internal static bool Supports(string formula)
    {
        bool Supported(FormulaNode n) => (n.Kind is "text" or "number" or "boolean" || n.Kind == "ref" && n.Sheet is null
            || n.Kind == "unary" && n.Text is "+" or "-" or "%"
            || n.Kind == "binary" && n.Text is "=" or "<>" or "<" or ">" or "<=" or ">=" or "+" or "-" or "*" or "/"
            || n.Kind == "call" && (n.Text is "AND" or "OR" || n.Text == "NOT" && n.Children.Count == 1)) && n.Children.All(Supported);
        try
        {
            return Supported(new FormulaParser(formula).Parse());
        }
        catch (Exception e) when (e is NotSupportedException or ArgumentException)
        {
            return false;
        }
    }

    internal static bool Matches(string formula, int rowOffset, int columnOffset, IReadOnlyDictionary<string, XlsxCell> cells)
    {
        object Eval(FormulaNode n)
        {
            switch (n.Kind)
            {
                case "text": return n.Text;
                case "number": return double.Parse(n.Text, CultureInfo.InvariantCulture);
                case "boolean": return n.Text == "TRUE";
                case "ref":
                    if (!cells.TryGetValue(n.Text, out var cell))
                    {
                        return 0d;
                    }
                    if (cell.Type == "e" || cell.Formula is not null && !cell.HasCachedValue)
                    {
                        throw new NotSupportedException();
                    }
                    return cell.Type == "n" && double.TryParse(cell.Value, CultureInfo.InvariantCulture, out var number) ? number : cell.Type == "b" ? cell.Value == "1" : cell.Value;
                case "call":
                    return n.Text == "OR" ? n.Children.Any(c => Truth(Eval(c))) : n.Text == "AND" ? n.Children.All(c => Truth(Eval(c))) : !Truth(Eval(n.Children[0]));
                case "unary":
                    var operand = Numeric(Eval(n.Children[0]));
                    return n.Text == "-" ? -operand : n.Text == "%" ? operand / 100 : operand;
                case "binary":
                    var a = Eval(n.Children[0]);
                    var b = Eval(n.Children[1]);
                    if (n.Text is "+" or "-" or "*" or "/")
                    {
                        var left = Numeric(a);
                        var right = Numeric(b);
                        var result = n.Text switch { "+" => left + right, "-" => left - right, "*" => left * right, _ => left / right };
                        if (!double.IsFinite(result))
                        {
                            throw new NotSupportedException();
                        }
                        return result;
                    }
                    var comparison = a is double x && b is double y ? x.CompareTo(y)
                        : a is string s && b is string t ? StringComparer.OrdinalIgnoreCase.Compare(s, t)
                        : a is bool ab && b is bool bb ? ab.CompareTo(bb) : Rank(a).CompareTo(Rank(b));
                    return n.Text switch { "=" => comparison == 0, "<>" => comparison != 0, "<" => comparison < 0, ">" => comparison > 0, "<=" => comparison <= 0, _ => comparison >= 0 };
                default: throw new NotSupportedException();
            }
        }
        static int Rank(object value) => value is double ? 0 : value is string ? 1 : 2;
        static double Numeric(object value) => value is double number ? number : throw new NotSupportedException();
        static bool Truth(object value) => value is bool b ? b : value is double n && n != 0;
        try
        {
            return Truth(Eval(new FormulaParser(SharedFormulas.Translate(formula, rowOffset, columnOffset)).Parse()));
        }
        catch (Exception e) when (e is NotSupportedException or ArgumentException or FormatException)
        {
            return false;
        }
    }
}
