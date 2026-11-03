using System.Text.RegularExpressions;
using System.Xml.Linq;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Calculation;

internal static class SharedFormulas
{
    internal static string Translate(string formula, int rows, int columns) => Regex.Replace(formula,
        "\"(?:[^\"]|\"\")*\"|'(?:[^']|'')*'!|(?<![\\w.])(?<col>\\$?[A-Za-z]{1,3})(?<row>\\$?[1-9][0-9]*)(?![\\w.(!])", match =>
        {
            if (!match.Groups["col"].Success)
            {
                return match.Value;
            }
            var c = match.Groups["col"].Value;
            var r = match.Groups["row"].Value;
            var location = CellAddress.Parse(c.Replace("$", "") + r.Replace("$", ""));
            var row = location.Row + (r.StartsWith('$') ? 0 : rows);
            var column = location.Column + (c.StartsWith('$') ? 0 : columns);
            if (row is < 1 or > 1048576 || column is < 1 or > 16384)
            {
                return "#REF!";
            }
            return (c.StartsWith('$') ? "$" : "") + CellAddress.Format(1, column)[..^1] + (r.StartsWith('$') ? "$" : "") + row;
        }, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    internal static void ExpandSheet(XDocument xml)
    {
        var formulas = xml.Descendants(S + "f").Where(f => (string?)f.Attribute("t") == "shared").ToArray();
        var masters = formulas.Where(f => f.Value.Length > 0).ToDictionary(f => (string)f.Attribute("si")!, f => ((string)f.Parent!.Attribute("r")!, f.Value));
        foreach (var formula in formulas)
        {
            if (!masters.TryGetValue((string?)formula.Attribute("si") ?? "", out var master))
            {
                throw new InvalidDataException("Shared formula master is missing.");
            }
            var start = CellAddress.Parse(master.Item1);
            var target = CellAddress.Parse((string)formula.Parent!.Attribute("r")!);
            formula.Value = Translate(master.Value, target.Row - start.Row, target.Column - start.Column);
            formula.Attribute("t")?.Remove();
            formula.Attribute("si")?.Remove();
            formula.Attribute("ref")?.Remove();
        }
    }

    internal static XlsxObjectDocument Expand(XlsxObjectDocument document)
    {
        foreach (var part in document.PartNames.Where(p => p.StartsWith("xl/worksheets/", StringComparison.Ordinal) && p.EndsWith(".xml", StringComparison.Ordinal)))
        {
            var xml = document.ReadXml(part);
            if (xml.Descendants(S + "f").Any(f => (string?)f.Attribute("t") == "shared"))
            {
                ExpandSheet(xml);
                document = document.WithXml(part, xml);
            }
        }
        return document;
    }
}
