using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Editors;

public static class XlsxStructureEditor
{
    // String literals are matched first so addresses inside them remain literal text.
    private static readonly Regex References = new("""
        "(?:[^"]|"")*"|(?<![\w.])(?:(?<sheet>'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!)?(?<ref>\$?[A-Za-z]{1,3}\$?[1-9][0-9]*|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9][0-9]*:\$?[1-9][0-9]*)(?::(?<end>\$?[A-Za-z]{1,3}\$?[1-9][0-9]*))?(?![\w.(])
        """, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    public static XlsxObjectDocument Insert(XlsxObjectDocument document, string sheetId, bool rows, int position, int count = 1)
        => Change(document, sheetId, rows, position, count, false);

    public static XlsxObjectDocument Delete(XlsxObjectDocument document, string sheetId, bool rows, int position, int count = 1)
        => Change(document, sheetId, rows, position, count, true);

    private static XlsxObjectDocument Change(XlsxObjectDocument document, string sheetId, bool rows, int position, int count, bool deleting)
    {
        var limit = rows ? 1048576 : 16384;
        if (position < 1 || position > limit || count < 1 || count > 10000 || position + count - 1 > limit)
        {
            throw new ArgumentException("Choose a valid position and 1-10,000 rows or columns.");
        }
        var sheets = XlsxWorkbookExtractor.Extract(document).Sheets;
        var target = sheets.Single(item => item.Id == sheetId);
        XlsxEditor.Guard(document, target.Part);
        if (document.PartNames.Any(part => part.StartsWith("xl/tables/", StringComparison.Ordinal)
            || part.StartsWith("xl/pivot", StringComparison.Ordinal) || part.StartsWith("xl/charts/", StringComparison.Ordinal)
            || part.StartsWith("xl/externalLinks/", StringComparison.Ordinal)))
        {
            throw new NotSupportedException("Structural editing in workbooks containing Excel tables, pivot tables, charts or external links is not supported yet.");
        }
        bool Removed(int value) => deleting && value >= position && value < position + count;
        int Shift(int value, int maximum)
        {
            var result = value < position ? value : deleting ? Math.Max(position, value - count) : value + count;
            if (result > maximum)
            {
                throw new ArgumentException("Insertion would move existing content or references beyond the worksheet limit.");
            }
            return result;
        }
        string Coordinate(string value)
        {
            var absolute = value.StartsWith('$');
            var text = value.Replace("$", "");
            if (int.TryParse(text, out var row))
            {
                return (absolute ? "$" : "") + (rows ? Shift(row, 1048576) : row).ToString(CultureInfo.InvariantCulture);
            }
            if (text.All(char.IsLetter))
            {
                var column = CellAddress.Parse(text + "1").Column;
                return (absolute ? "$" : "") + CellAddress.Format(1, rows ? column : Shift(column, 16384))[..^1];
            }
            var cell = CellAddress.Parse(text);
            var formatted = CellAddress.Format(rows ? Shift(cell.Row, 1048576) : cell.Row, rows ? cell.Column : Shift(cell.Column, 16384));
            var split = formatted.TakeWhile(char.IsLetter).Count();
            return (absolute ? "$" : "") + formatted[..split] + (value.IndexOf('$', absolute ? 1 : 0) >= 0 ? "$" : "") + formatted[split..];
        }
        int Axis(string value)
        {
            var text = value.Replace("$", "");
            if (int.TryParse(text, out var row))
            {
                return rows ? row : 0;
            }
            if (text.All(char.IsLetter))
            {
                return rows ? 0 : CellAddress.Parse(text + "1").Column;
            }
            var cell = CellAddress.Parse(text);
            return rows ? cell.Row : cell.Column;
        }
        string WithAxis(string value, int index)
        {
            if (rows)
            {
                return Regex.Replace(value, "[0-9]+$", index.ToString(CultureInfo.InvariantCulture));
            }
            return Regex.Replace(value, "[A-Za-z]+", CellAddress.Format(1, index)[..^1]);
        }
        string Area(string value)
        {
            var ends = value.Split(':');
            if (!deleting)
            {
                return string.Join(':', ends.Select(Coordinate));
            }
            if (ends.Length == 1)
            {
                return Removed(Axis(value)) ? "#REF!" : Coordinate(value);
            }
            var a = Axis(ends[0]);
            var b = Axis(ends[1]);
            var low = Math.Min(a, b);
            var high = Math.Max(a, b);
            if (Removed(low) && Removed(high))
            {
                return "#REF!";
            }
            // Trim deleted endpoints before shifting the surviving interval.
            int Surviving(int index, bool lower) => Removed(index) ? (lower ? position + count : position - 1) : index;
            return Coordinate(WithAxisIfAffected(ends[0], Surviving(a, a <= b))) + ":" + Coordinate(WithAxisIfAffected(ends[1], Surviving(b, b < a)));
        }
        string WithAxisIfAffected(string value, int index) => Axis(value) == 0 ? value : WithAxis(value, index);
        string Range(string value) => string.Join(' ', value.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(Area));
        string Formula(string formula, string? context)
        {
            var nonStrings = Regex.Replace(formula, "\"(?:[^\"]|\"\")*\"", "");
            if (nonStrings.Contains('[') || nonStrings.Contains(']') || Regex.IsMatch(nonStrings, @"(?:[\w']+):(?:[\w']+)!|'(?:[^']|'')*:(?:[^']|'')*'!"))
            {
                throw new NotSupportedException("Structural editing with structured, external or 3-D formula references is not supported yet.");
            }
            return References.Replace(formula, match =>
            {
                if (match.Value.StartsWith('"'))
                {
                    return match.Value;
                }
                var prefix = match.Groups["sheet"].Value;
                var name = prefix.Length == 0 ? context : (prefix.StartsWith('\'') ? prefix[1..^1].Replace("''", "'") : prefix);
                if (!string.Equals(name, target.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return match.Value;
                }
                var reference = match.Groups["ref"].Value;
                // Names resembling cells outside Excel's address space are defined names.
                try
                {
                    foreach (var end in reference.Split(':'))
                    {
                        var plain = end.Replace("$", "");
                        if (plain.Any(char.IsLetter))
                        {
                            CellAddress.Parse(plain.All(char.IsLetter) ? plain + "1" : plain);
                        }
                    }
                }
                catch (ArgumentException)
                {
                    return match.Value;
                }
                var updated = Range(reference + (match.Groups["end"].Success ? ":" + match.Groups["end"].Value : ""));
                return updated == "#REF!" ? updated : (prefix.Length == 0 ? "" : prefix + "!") + updated;
            });
        }
        foreach (var sheet in sheets)
        {
            XlsxEditor.Guard(document, sheet.Part);
            var xml = document.ReadXml(sheet.Part);
            if (xml.Descendants(S + "f").Any(f => f.Attribute("t") is not null && (string?)f.Attribute("t") != "normal"))
            {
                throw new NotSupportedException("Structural editing with shared, array or data-table formulas is not supported yet.");
            }
            if (xml.Descendants().Any(e => e.Name.LocalName == "extLst"))
            {
                throw new NotSupportedException("Structural editing with worksheet extension features is not supported yet.");
            }
            foreach (var formula in xml.Descendants().Where(e => e.Name == S + "f" || e.Name == S + "formula" || e.Name == S + "formula1" || e.Name == S + "formula2"))
            {
                formula.Value = Formula(formula.Value, sheet.Name);
            }
            foreach (var link in xml.Descendants(S + "hyperlink").Attributes("location"))
            {
                link.Value = Formula(link.Value, sheet.Name);
            }
            if (sheet.Id == sheetId)
            {
                if (xml.Descendants().Any(e => e.Name == S + "legacyDrawing" || e.Name == S + "oleObjects" || e.Name == S + "controls"))
                {
                    throw new NotSupportedException("Structural editing with legacy drawings, comments or embedded controls is not supported yet.");
                }
                if (deleting)
                {
                    if (rows)
                    {
                        xml.Descendants(S + "row").Where(r => Removed((int)r.Attribute("r")!)).Remove();
                    }
                    else
                    {
                        xml.Descendants(S + "c").Where(c => Removed(CellAddress.Parse((string)c.Attribute("r")!).Column)).Remove();
                    }
                }
                foreach (var cell in xml.Descendants(S + "c"))
                {
                    cell.SetAttributeValue("r", Coordinate((string)cell.Attribute("r")!));
                }
                foreach (var row in xml.Descendants(S + "row"))
                {
                    if (rows)
                    {
                        row.SetAttributeValue("r", Shift((int)row.Attribute("r")!, 1048576));
                    }
                    row.Attribute("spans")?.Remove();
                }
                if (!rows)
                {
                    foreach (var filter in xml.Descendants(S + "autoFilter"))
                    {
                        var firstColumn = CellAddress.Parse(((string)filter.Attribute("ref")!).Split(':')[0].Replace("$", "")).Column;
                        var newFirstColumn = Shift(firstColumn, 16384);
                        foreach (var filterColumn in filter.Elements(S + "filterColumn").ToArray())
                        {
                            var column = firstColumn + (int)filterColumn.Attribute("colId")!;
                            if (Removed(column))
                            {
                                filterColumn.Remove();
                            }
                            else
                            {
                                filterColumn.SetAttributeValue("colId", Shift(column, 16384) - newFirstColumn);
                            }
                        }
                    }
                    foreach (var col in xml.Descendants(S + "col").ToArray())
                    {
                        var min = (int)col.Attribute("min")!;
                        var max = (int)col.Attribute("max")!;
                        if (deleting)
                        {
                            if (Removed(min) && Removed(max))
                            {
                                col.Remove();
                                continue;
                            }
                            min = Shift(Removed(min) ? position + count : min, 16384);
                            max = Shift(Removed(max) ? position - 1 : max, 16384);
                        }
                        else
                        {
                            min += min >= position ? count : 0;
                            max += max >= position ? count : 0;
                        }
                        if (min > 16384)
                        {
                            col.Remove();
                        }
                        else
                        {
                            col.SetAttributeValue("min", min);
                            col.SetAttributeValue("max", Math.Min(16384, max));
                        }
                    }
                }
                foreach (var attribute in xml.Descendants().Attributes().Where(a => a.Name.LocalName is "ref" or "sqref" or "activeCell" or "topLeftCell").ToArray())
                {
                    if (attribute.Parent?.Document is null)
                    {
                        continue;
                    }
                    var updated = string.Join(' ', Range(attribute.Value).Split(' ').Where(area => area != "#REF!"));
                    if (updated.Length > 0)
                    {
                        attribute.Value = updated;
                    }
                    else if (attribute.Name.LocalName is "activeCell" or "topLeftCell" || attribute.Parent.Name == S + "selection")
                    {
                        attribute.Value = Coordinate(attribute.Value.Split(' ')[0].Split(':')[0]);
                    }
                    else
                    {
                        attribute.Parent.Remove();
                    }
                }
                foreach (var pane in xml.Descendants(S + "pane").Where(e => ((string?)e.Attribute("state"))?.StartsWith("frozen", StringComparison.Ordinal) == true))
                {
                    var split = pane.Attribute(rows ? "ySplit" : "xSplit");
                    if (split is not null && double.TryParse(split.Value, CultureInfo.InvariantCulture, out var value) && value >= position)
                    {
                        split.Value = (deleting ? value - Math.Min(count, value - position + 1) : value + count).ToString(CultureInfo.InvariantCulture);
                    }
                }
                foreach (var brk in xml.Descendants(S + (rows ? "rowBreaks" : "colBreaks")).Elements(S + "brk").ToArray())
                {
                    if (Removed((int)brk.Attribute("id")!))
                    {
                        brk.Remove();
                        continue;
                    }
                    brk.SetAttributeValue("id", Shift((int)brk.Attribute("id")!, limit));
                }
                foreach (var drawing in xml.Descendants(S + "drawing"))
                {
                    var part = XlsxImageExtractor.Target(document, sheet.Part, (string?)drawing.Attribute(R + "id"))!;
                    if (sheets.Where(s => s.Id != sheetId).Any(s => document.ReadXml(s.Part).Descendants(S + "drawing").Any(d => XlsxImageExtractor.Target(document, s.Part, (string?)d.Attribute(R + "id")) == part)))
                    {
                        throw new NotSupportedException("Structural editing with shared drawing parts is not supported yet.");
                    }
                    var drawingXml = document.ReadXml(part);
                    foreach (var marker in drawingXml.Descendants().Where(e => e.Name == XlsxImageExtractor.D + "from" || e.Name == XlsxImageExtractor.D + "to"))
                    {
                        var coordinate = marker.Element(XlsxImageExtractor.D + (rows ? "row" : "col"))!;
                        coordinate.Value = (Shift((int)coordinate + 1, limit) - 1).ToString(CultureInfo.InvariantCulture);
                    }
                    if (deleting)
                    {
                        foreach (var anchor in drawingXml.Root!.Elements(XlsxImageExtractor.D + "twoCellAnchor").ToArray())
                        {
                            var from = anchor.Element(XlsxImageExtractor.D + "from")!;
                            var to = anchor.Element(XlsxImageExtractor.D + "to")!;
                            var axis = rows ? "row" : "col";
                            if ((int)from.Element(XlsxImageExtractor.D + axis)! == (int)to.Element(XlsxImageExtractor.D + axis)!
                                && (long)from.Element(XlsxImageExtractor.D + (axis + "Off"))! >= (long)to.Element(XlsxImageExtractor.D + (axis + "Off"))!)
                            {
                                anchor.Remove();
                            }
                        }
                    }
                    document = document.WithXml(part, drawingXml);
                }
            }
            if (deleting && sheet.Id == sheetId)
            {
                foreach (var merge in xml.Descendants(S + "mergeCell").Where(e => ((string)e.Attribute("ref")!).Split(':').Distinct().Count() == 1).ToArray())
                {
                    merge.Remove();
                }
                foreach (var container in xml.Root!.Elements().Where(e => e.Name == S + "mergeCells" || e.Name == S + "dataValidations" || e.Name == S + "rowBreaks" || e.Name == S + "colBreaks").ToArray())
                {
                    var remaining = container.Elements().Count();
                    if (remaining == 0)
                    {
                        container.Remove();
                    }
                    else
                    {
                        container.SetAttributeValue("count", remaining);
                        if (container.Attribute("manualBreakCount") is not null)
                        {
                            container.SetAttributeValue("manualBreakCount", container.Elements().Count(e => (bool?)e.Attribute("man") == true));
                        }
                    }
                }
                xml.Root.Elements().Where(e => (e.Name == S + "cols" || e.Name == S + "hyperlinks") && !e.HasElements).Remove();
            }
            document = document.WithXml(sheet.Part, xml);
        }
        var workbook = document.ReadXml("xl/workbook.xml");
        foreach (var name in workbook.Descendants(S + "definedName"))
        {
            var local = (int?)name.Attribute("localSheetId");
            name.Value = Formula(name.Value, local.HasValue ? sheets[local.Value].Name : null);
        }
        workbook.Root!.Element(S + "calcPr")?.Remove();
        // Drop stale calculation chains; Excel will reconstruct dependencies when opened.
        var relationships = document.ReadXml("xl/_rels/workbook.xml.rels");
        var types = document.ReadXml("[Content_Types].xml");
        foreach (var relation in relationships.Root!.Elements().Where(e => ((string?)e.Attribute("Type"))?.EndsWith("/calcChain", StringComparison.Ordinal) == true).ToArray())
        {
            var part = Resolve("xl/workbook.xml", (string)relation.Attribute("Target")!);
            document = new XlsxObjectDocument(document.PartNames.Where(p => p != part).ToDictionary(p => p, document.GetPart, StringComparer.Ordinal));
            types.Root!.Elements().Where(e => (string?)e.Attribute("PartName") == "/" + part).Remove();
            relation.Remove();
        }
        return document.WithXml("xl/workbook.xml", workbook).WithXml("xl/_rels/workbook.xml.rels", relationships).WithXml("[Content_Types].xml", types);
    }
}
