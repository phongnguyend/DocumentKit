using System.Xml.Linq;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Editors;

public static class XlsxVisibilityEditor
{
    public static XlsxObjectDocument SetHidden(XlsxObjectDocument document, string sheetId, bool rows, int position, int count, bool hidden)
    {
        var limit = rows ? 1048576 : 16384;
        if (position < 1 || count < 1 || count > limit || position > limit - count + 1 || hidden && count > 10000)
        {
            throw new ArgumentException("Choose a valid range; hide up to 10,000 rows or columns at once.");
        }
        var sheet = XlsxWorkbookExtractor.Extract(document).Sheets.Single(s => s.Id == sheetId);
        XlsxEditor.Guard(document, sheet.Part);
        var xml = document.ReadXml(sheet.Part);
        var end = position + count - 1;
        if (rows)
        {
            var data = xml.Root!.Element(S + "sheetData")!;
            var existing = data.Elements(S + "row").ToDictionary(r => (int)r.Attribute("r")!);
            foreach (var row in existing.Where(r => r.Key >= position && r.Key <= end).Select(r => r.Value))
            {
                row.SetAttributeValue("hidden", hidden ? "1" : null);
            }
            if (hidden)
            {
                for (var index = position; index <= end; index++)
                {
                    if (!existing.ContainsKey(index))
                    {
                        data.Add(new XElement(S + "row", new XAttribute("r", index), new XAttribute("hidden", "1")));
                    }
                }
                var sorted = data.Elements().OrderBy(r => (int)r.Attribute("r")!).ToArray();
                data.ReplaceNodes(sorted);
            }
        }
        else
        {
            var cols = xml.Root!.Element(S + "cols");
            if (cols is null)
            {
                if (!hidden)
                {
                    return document;
                }
                cols = new XElement(S + "cols");
                xml.Root.Element(S + "sheetData")!.AddBeforeSelf(cols);
            }
            var covered = new HashSet<int>();
            foreach (var col in cols.Elements().ToArray())
            {
                var first = (int)col.Attribute("min")!;
                var last = (int)col.Attribute("max")!;
                if (first > end || last < position)
                {
                    continue;
                }
                void Add(int a, int b, bool change)
                {
                    if (a > b)
                    {
                        return;
                    }
                    var copy = new XElement(col);
                    copy.SetAttributeValue("min", a);
                    copy.SetAttributeValue("max", b);
                    if (change)
                    {
                        copy.SetAttributeValue("hidden", hidden ? "1" : null);
                        for (var i = a; i <= b; i++)
                        {
                            covered.Add(i);
                        }
                    }
                    cols.Add(copy);
                }
                Add(first, position - 1, false);
                Add(Math.Max(first, position), Math.Min(last, end), true);
                Add(end + 1, last, false);
                col.Remove();
            }
            if (hidden)
            {
                for (var i = position; i <= end; i++)
                {
                    if (!covered.Contains(i))
                    {
                        cols.Add(new XElement(S + "col", new XAttribute("min", i), new XAttribute("max", i), new XAttribute("hidden", "1")));
                    }
                }
            }
            cols.ReplaceNodes(cols.Elements().OrderBy(c => (int)c.Attribute("min")!).ToArray());
        }
        return document.WithXml(sheet.Part, xml);
    }
}
