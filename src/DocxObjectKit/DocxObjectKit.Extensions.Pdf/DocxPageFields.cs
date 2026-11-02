using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private sealed class PageField
    {
        public StringBuilder Instruction { get; } = new();

        public List<XElement> Texts { get; } = new();

        public bool Result { get; set; }
    }

    private static void UpdatePageFields(XDocument xml, int pageNumber, int totalPages, string sectionFormat)
    {
        string? Evaluate(string instruction)
        {
            var command = Regex.Match(instruction, @"^\s*(PAGE|NUMPAGES)(?:\s|$)", RegexOptions.IgnoreCase);
            if (!command.Success)
            {
                return null;
            }
            var total = command.Groups[1].Value.Equals("NUMPAGES", StringComparison.OrdinalIgnoreCase);
            var format = total ? "decimal" : sectionFormat;
            var explicitFormat = Regex.Match(instruction, @"\\\*\s+(ROMAN|roman|ALPHABETIC|alphabetic|Arabic)(?:\s|$)");
            format = explicitFormat.Groups[1].Value switch
            {
                "ROMAN" => "upperRoman",
                "roman" => "lowerRoman",
                "ALPHABETIC" => "upperLetter",
                "alphabetic" => "lowerLetter",
                "Arabic" => "decimal",
                _ => format
            };
            return FormatNumber(total ? totalPages : pageNumber, format);
        }
        void Replace(IEnumerable<XElement> values, string? value)
        {
            if (value is null)
            {
                return;
            }
            var first = true;
            foreach (var text in values)
            {
                text.Value = first ? value : "";
                first = false;
            }
        }
        foreach (var field in xml.Descendants(W + "fldSimple"))
        {
            var value = Evaluate(Value(field, "instr") ?? "");
            if (value is null)
            {
                continue;
            }
            var texts = field.Descendants(W + "t").ToArray();
            if (texts.Length == 0)
            {
                field.Add(new XElement(W + "r", new XElement(W + "t", value)));
            }
            else
            {
                Replace(texts, value);
            }
        }
        var fields = new Stack<PageField>();
        foreach (var run in xml.Descendants(W + "r").Where(run => !run.Ancestors(W + "fldSimple").Any()))
        {
            foreach (var node in run.Elements())
            {
                if (node.Name == W + "fldChar")
                {
                    switch (Value(node, "fldCharType"))
                    {
                        case "begin":
                            fields.Push(new());
                            break;
                        case "separate" when fields.Count > 0:
                            fields.Peek().Result = true;
                            break;
                        case "end" when fields.Count > 0:
                            var field = fields.Pop();
                            if (field.Result)
                            {
                                Replace(field.Texts, Evaluate(field.Instruction.ToString()));
                            }
                            break;
                    }
                }
                else if (fields.Count > 0 && node.Name == W + "instrText" && !fields.Peek().Result)
                {
                    fields.Peek().Instruction.Append(node.Value);
                }
                else if (fields.Count > 0 && node.Name == W + "t" && fields.Peek().Result)
                {
                    fields.Peek().Texts.Add(node);
                }
            }
        }
    }
}
