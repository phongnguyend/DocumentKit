using System.Globalization;
using System.Text.RegularExpressions;

namespace XlsxObjectKit.Calculation;

internal sealed record FormulaNode(string Kind, string Text, IReadOnlyList<FormulaNode> Children, string? Sheet = null);

internal sealed class FormulaParser
{
    private static readonly Regex Tokens = new("\\G\\s*(?<token>\"(?:[^\"]|\"\")*\"|'(?:[^']|'')*'|(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)(?:[eE][+-]?[0-9]+)?|#[A-Z0-9/]+[!?]?|[A-Za-z_$][A-Za-z0-9_.$]*|<=|>=|<>|[+*/^%&=<>(),!:\\-])", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
    private readonly List<string> tokens = [];
    private int position;
    private int depth;

    public FormulaParser(string formula)
    {
        if (formula.Length > 8192)
        {
            throw new NotSupportedException("Formula exceeds 8,192 characters.");
        }
        for (var offset = 0; offset < formula.Length;)
        {
            var match = Tokens.Match(formula, offset);
            if (!match.Success)
            {
                if (string.IsNullOrWhiteSpace(formula[offset..]))
                {
                    break;
                }
                throw new NotSupportedException("Unsupported formula syntax.");
            }
            tokens.Add(match.Groups["token"].Value);
            offset += match.Length;
        }
    }

    private string Current => position < tokens.Count ? tokens[position] : "";

    private string Take() => position < tokens.Count ? tokens[position++] : throw new NotSupportedException("Incomplete formula.");

    private void Require(string token)
    {
        if (Take() != token)
        {
            throw new NotSupportedException("Invalid formula syntax.");
        }
    }

    public FormulaNode Parse()
    {
        var result = Expression(0);
        if (Current != "")
        {
            throw new NotSupportedException("Unsupported formula syntax or reference.");
        }
        return result;
    }

    private FormulaNode Expression(int minimum)
    {
        if (++depth > 64)
        {
            throw new NotSupportedException("Formula nesting limit reached.");
        }
        FormulaNode left;
        var token = Take();
        if (token is "+" or "-")
        {
            left = new("unary", token, [Expression(6)]);
        }
        else if (token == "(")
        {
            left = Expression(0);
            Require(")");
        }
        else if (token.StartsWith('"'))
        {
            left = new("text", token[1..^1].Replace("\"\"", "\""), []);
        }
        else if (token.StartsWith('#'))
        {
            left = new("error", token, []);
        }
        else if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
        {
            left = new("number", token, []);
        }
        else if (Current == "(")
        {
            Take();
            var arguments = new List<FormulaNode>();
            if (Current != ")")
            {
                while (true)
                {
                    arguments.Add(Current is "," or ")" ? new("blank", "", []) : Expression(0));
                    if (Current != ",")
                    {
                        break;
                    }
                    Take();
                }
            }
            Require(")");
            left = new("call", token.ToUpperInvariant(), arguments);
        }
        else if (token.ToUpperInvariant() is "TRUE" or "FALSE" && Current != "!")
        {
            left = new("boolean", token.ToUpperInvariant(), []);
        }
        else
        {
            string? sheet = null;
            if (Current == "!")
            {
                Take();
                sheet = token.StartsWith('\'') ? token[1..^1].Replace("''", "'") : token;
                token = Take();
            }
            left = new("ref", token.Replace("$", "").ToUpperInvariant(), [], sheet);
            if (Current == ":")
            {
                Take();
                left = new("range", "", [left, new("ref", Take().Replace("$", "").ToUpperInvariant(), [], sheet)]);
            }
        }
        while (true)
        {
            if (Current == "%")
            {
                Take();
                left = new("unary", "%", [left]);
                continue;
            }
            var precedence = Current switch { "=" or "<>" or "<" or ">" or "<=" or ">=" => 1, "&" => 2, "+" or "-" => 3, "*" or "/" => 4, "^" => 5, _ => -1 };
            if (precedence < minimum)
            {
                break;
            }
            var operation = Take();
            left = new("binary", operation, [left, Expression(precedence + 1)]);
        }
        depth--;
        return left;
    }
}
