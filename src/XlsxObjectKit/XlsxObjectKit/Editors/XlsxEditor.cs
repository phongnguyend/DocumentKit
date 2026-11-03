using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using XlsxObjectKit.Extractors;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

namespace XlsxObjectKit.Editors;

public sealed record XlsxCellEdit(string Address, string Value, string Type = "text");

public sealed record XlsxFormat(bool Bold = false, bool Italic = false, string Color = "#24334a", string Background = "#ffffff", string Alignment = "left");

public static class XlsxEditor
{
    public static XlsxObjectDocument Create()
    {
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        void Add(string name, string xml)
        {
            parts.Add(name, Encoding.UTF8.GetBytes(xml));
        }
        Add("[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>
            """);
        Add("_rels/.rels", $"<Relationships xmlns=\"{Rel}\"><Relationship Id=\"book\" Type=\"{R}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
        Add("xl/workbook.xml", $"<workbook xmlns=\"{S}\" xmlns:r=\"{R}\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"sheet1\"/></sheets></workbook>");
        Add("xl/_rels/workbook.xml.rels", $"<Relationships xmlns=\"{Rel}\"><Relationship Id=\"sheet1\" Type=\"{R}/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
        Add("xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{S}\"><sheetData/></worksheet>");
        return new(parts);
    }

    private static void Guard(XlsxObjectDocument document, string? part = null)
    {
        if (document.PartNames.Any(name => name.StartsWith("_xmlsignatures/", StringComparison.Ordinal))
            || document.ReadXml("xl/workbook.xml").Descendants(S + "workbookProtection").Any()
            || part is not null && document.ReadXml(part).Descendants(S + "sheetProtection").Any())
        {
            throw new NotSupportedException("Signed or protected workbooks/worksheets cannot be edited.");
        }
    }

    public static XlsxObjectDocument SetCells(XlsxObjectDocument document, string sheetId, IReadOnlyList<XlsxCellEdit> edits, XlsxFormat? format = null)
    {
        if (edits.Count is < 1 or > 10000)
        {
            throw new ArgumentException("Edit 1–10,000 cells at a time.");
        }
        var sheet = XlsxWorkbookExtractor.Extract(document).Sheets.Single(item => item.Id == sheetId);
        Guard(document, sheet.Part);
        var xml = document.ReadXml(sheet.Part);
        var data = xml.Root!.Element(S + "sheetData")!;
        var styles = new Dictionary<int, int>();
        foreach (var edit in edits)
        {
            ArgumentNullException.ThrowIfNull(edit);
            ArgumentNullException.ThrowIfNull(edit.Address);
            ArgumentNullException.ThrowIfNull(edit.Value);
            var address = edit.Address.ToUpperInvariant();
            var (rowNumber, column) = CellAddress.Parse(address);
            if (edit.Value.Length > 32767)
            {
                throw new ArgumentException("Cell text exceeds 32,767 characters.");
            }
            foreach (var merge in sheet.Merges)
            {
                var ends = merge.Split(':');
                var first = CellAddress.Parse(ends[0]);
                var last = CellAddress.Parse(ends[^1]);
                if (rowNumber >= first.Row && rowNumber <= last.Row && column >= first.Column && column <= last.Column && address != ends[0])
                {
                    throw new NotSupportedException("Edit the top-left cell of a merged range.");
                }
            }
            var row = data.Elements(S + "row").FirstOrDefault(item => (int?)item.Attribute("r") == rowNumber);
            if (row is null)
            {
                row = new XElement(S + "row", new XAttribute("r", rowNumber));
                var next = data.Elements(S + "row").FirstOrDefault(item => (int?)item.Attribute("r") > rowNumber);
                if (next is null)
                {
                    data.Add(row);
                }
                else
                {
                    next.AddBeforeSelf(row);
                }
            }
            var cell = row.Elements(S + "c").FirstOrDefault(item => (string?)item.Attribute("r") == address);
            if (cell?.Element(S + "f")?.Attribute("t") is not null || xml.Descendants(S + "f").Any(f => f.Attribute("ref") is not null))
            {
                throw new NotSupportedException("Shared, array and spill formula ranges require Excel for editing.");
            }
            if (cell is null)
            {
                cell = new XElement(S + "c", new XAttribute("r", address));
                var next = row.Elements(S + "c").FirstOrDefault(item => CellAddress.Parse((string)item.Attribute("r")!).Column > column);
                if (next is null)
                {
                    row.Add(cell);
                }
                else
                {
                    next.AddBeforeSelf(cell);
                }
            }
            if (format is not null)
            {
                var originalStyle = (int?)cell.Attribute("s") ?? sheet.RowStyles.GetValueOrDefault(rowNumber, sheet.ColumnStyles.GetValueOrDefault(column, 0));
                if (!styles.TryGetValue(originalStyle, out var style))
                {
                    style = AddStyle(ref document, format, originalStyle);
                    styles.Add(originalStyle, style);
                }
                cell.SetAttributeValue("s", style);
                continue;
            }
            cell.Elements().Where(item => item.Name == S + "f" || item.Name == S + "v" || item.Name == S + "is").Remove();
            cell.Attribute("t")?.Remove();
            switch (edit.Type)
            {
                case "clear":
                    break;
                case "text":
                    cell.SetAttributeValue("t", "inlineStr");
                    cell.Add(new XElement(S + "is", new XElement(S + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), edit.Value)));
                    break;
                case "number":
                    if (!double.TryParse(edit.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !double.IsFinite(number))
                    {
                        throw new ArgumentException("Enter a finite number using a decimal point.");
                    }
                    cell.Add(new XElement(S + "v", number.ToString("R", CultureInfo.InvariantCulture)));
                    break;
                case "boolean":
                    if (!bool.TryParse(edit.Value, out var boolean))
                    {
                        throw new ArgumentException("Enter true or false.");
                    }
                    cell.SetAttributeValue("t", "b");
                    cell.Add(new XElement(S + "v", boolean ? "1" : "0"));
                    break;
                case "formula":
                    var formula = edit.Value.TrimStart('=');
                    if (string.IsNullOrWhiteSpace(formula) || formula.Length > 8192)
                    {
                        throw new ArgumentException("Enter a formula of 1–8192 characters.");
                    }
                    cell.Add(new XElement(S + "f", formula));
                    break;
                default:
                    throw new ArgumentException("Unknown cell type.");
            }
        }
        xml.Root.Element(S + "dimension")?.Remove();
        document = document.WithXml(sheet.Part, xml);
        return Recalculate(document);
    }

    private static XlsxObjectDocument Recalculate(XlsxObjectDocument document)
    {
        foreach (var sheet in XlsxWorkbookExtractor.Extract(document).Sheets)
        {
            var xml = document.ReadXml(sheet.Part);
            foreach (var cell in xml.Descendants(S + "c").Where(item => item.Element(S + "f") is not null))
            {
                cell.Element(S + "v")?.Remove();
            }
            document = document.WithXml(sheet.Part, xml);
        }
        var book = document.ReadXml("xl/workbook.xml");
        var calc = book.Root!.Element(S + "calcPr");
        if (calc is null)
        {
            calc = new XElement(S + "calcPr");
            (book.Root.Elements().LastOrDefault(item => item.Name.LocalName is "definedNames" or "externalReferences" or "functionGroups") ?? book.Root.Element(S + "sheets")!).AddAfterSelf(calc);
        }
        calc.SetAttributeValue("fullCalcOnLoad", "1");
        calc.SetAttributeValue("forceFullCalc", "1");
        calc.SetAttributeValue("calcMode", "auto");
        // Remove stale calculation-chain references; keep unrelated package parts untouched.
        var rels = document.ReadXml("xl/_rels/workbook.xml.rels");
        rels.Root!.Elements().Where(item => ((string?)item.Attribute("Type"))?.EndsWith("/calcChain") == true).Remove();
        return document.WithXml("xl/workbook.xml", book).WithXml("xl/_rels/workbook.xml.rels", rels);
    }

    private static int AddStyle(ref XlsxObjectDocument document, XlsxFormat format, int originalStyle)
    {
        if (!Regex.IsMatch(format.Color, "^#[0-9a-fA-F]{6}$") || !Regex.IsMatch(format.Background, "^#[0-9a-fA-F]{6}$") || format.Alignment is not ("general" or "left" or "center" or "right" or "justify" or "distributed" or "fill" or "centerContinuous"))
        {
            throw new ArgumentException("Invalid cell formatting.");
        }
        var part = XlsxWorkbookExtractor.StylesPart(document);
        XDocument styles;
        if (part is null)
        {
            part = "xl/styles.xml";
            styles = XDocument.Parse($"<styleSheet xmlns=\"{S}\"><fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf/></cellStyleXfs><cellXfs count=\"1\"><xf fontId=\"0\" fillId=\"0\" borderId=\"0\" numFmtId=\"0\" xfId=\"0\"/></cellXfs></styleSheet>");
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
        var fonts = styles.Root!.Element(S + "fonts")!;
        var fills = styles.Root.Element(S + "fills")!;
        var xfs = styles.Root.Element(S + "cellXfs")!;
        var fontId = fonts.Elements().Count();
        var fillId = fills.Elements().Count();
        var styleId = xfs.Elements().Count();
        if (styleId > 60000)
        {
            throw new NotSupportedException("Workbook style limit reached.");
        }
        var previous = xfs.Elements().ElementAtOrDefault(originalStyle) ?? throw new InvalidDataException("Invalid style index.");
        var parentStyle = styles.Root.Element(S + "cellStyleXfs")?.Elements().ElementAtOrDefault((int?)previous.Attribute("xfId") ?? 0);
        var previousFontId = (string?)previous.Attribute("applyFont") is "0" or "false"
            ? (int?)parentStyle?.Attribute("fontId") ?? 0
            : (int?)previous.Attribute("fontId") ?? (int?)parentStyle?.Attribute("fontId") ?? 0;
        var previousFont = fonts.Elements().ElementAtOrDefault(previousFontId) ?? throw new InvalidDataException("Invalid font index.");
        var font = new XElement(previousFont);
        font.Elements().Where(item => item.Name == S + "b" || item.Name == S + "i" || item.Name == S + "color").Remove();
        if (format.Bold)
        {
            font.AddFirst(new XElement(S + "b"));
        }
        if (format.Italic)
        {
            if (font.Element(S + "b") is { } bold)
            {
                bold.AddAfterSelf(new XElement(S + "i"));
            }
            else
            {
                font.AddFirst(new XElement(S + "i"));
            }
        }
        var color = new XElement(S + "color", new XAttribute("rgb", "FF" + format.Color[1..]));
        if (font.Element(S + "name") is { } fontName)
        {
            fontName.AddBeforeSelf(color);
        }
        else
        {
            font.Add(color);
        }
        fonts.Add(font);
        fills.Add(new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "solid"), new XElement(S + "fgColor", new XAttribute("rgb", "FF" + format.Background[1..])))));
        var xf = new XElement(previous);
        xf.SetAttributeValue("fontId", fontId);
        xf.SetAttributeValue("fillId", fillId);
        xf.SetAttributeValue("applyFont", 1);
        xf.SetAttributeValue("applyFill", 1);
        xf.SetAttributeValue("applyAlignment", 1);
        var alignment = new XElement(parentStyle?.Element(S + "alignment") ?? new XElement(S + "alignment"));
        if ((string?)previous.Attribute("applyAlignment") is not ("0" or "false") && previous.Element(S + "alignment") is { } directAlignment)
        {
            foreach (var attribute in directAlignment.Attributes())
            {
                alignment.SetAttributeValue(attribute.Name, attribute.Value);
            }
        }
        xf.Element(S + "alignment")?.Remove();
        xf.AddFirst(alignment);
        alignment.SetAttributeValue("horizontal", format.Alignment);
        xfs.Add(xf);
        fonts.SetAttributeValue("count", fontId + 1);
        fills.SetAttributeValue("count", fillId + 1);
        xfs.SetAttributeValue("count", styleId + 1);
        document = document.WithXml(part, styles);
        return styleId;
    }

    public static XlsxObjectDocument Dimensions(XlsxObjectDocument document, string sheetId, string address, double rowHeight, double columnWidth)
    {
        if (!double.IsFinite(rowHeight) || rowHeight < 1 || rowHeight > 409 || !double.IsFinite(columnWidth) || columnWidth < 1 || columnWidth > 255)
        {
            throw new ArgumentException("Row height must be 1–409 points and column width 1–255 characters.");
        }
        var sheet = XlsxWorkbookExtractor.Extract(document).Sheets.Single(item => item.Id == sheetId);
        Guard(document, sheet.Part);
        var (rowNumber, column) = CellAddress.Parse(address);
        var xml = document.ReadXml(sheet.Part);
        var data = xml.Root!.Element(S + "sheetData")!;
        var row = data.Elements(S + "row").FirstOrDefault(item => (int?)item.Attribute("r") == rowNumber);
        if (row is null)
        {
            row = new XElement(S + "row", new XAttribute("r", rowNumber));
            var next = data.Elements().FirstOrDefault(item => (int?)item.Attribute("r") > rowNumber);
            if (next is null)
            {
                data.Add(row);
            }
            else
            {
                next.AddBeforeSelf(row);
            }
        }
        row.SetAttributeValue("ht", rowHeight.ToString(CultureInfo.InvariantCulture));
        row.SetAttributeValue("customHeight", 1);
        var cols = xml.Root.Element(S + "cols");
        if (cols is null)
        {
            cols = new XElement(S + "cols");
            data.AddBeforeSelf(cols);
        }
        var previous = cols.Elements(S + "col").FirstOrDefault(item => (int?)item.Attribute("min") <= column && (int?)item.Attribute("max") >= column);
        var replacement = previous is null ? new XElement(S + "col") : new XElement(previous);
        if (previous is not null)
        {
            var min = (int)previous.Attribute("min")!;
            var max = (int)previous.Attribute("max")!;
            if (min < column)
            {
                var left = new XElement(previous);
                left.SetAttributeValue("max", column - 1);
                previous.AddBeforeSelf(left);
            }
            if (max > column)
            {
                var right = new XElement(previous);
                right.SetAttributeValue("min", column + 1);
                previous.AddAfterSelf(right);
            }
            previous.Remove();
        }
        replacement.SetAttributeValue("min", column);
        replacement.SetAttributeValue("max", column);
        replacement.SetAttributeValue("width", columnWidth.ToString(CultureInfo.InvariantCulture));
        replacement.SetAttributeValue("customWidth", 1);
        cols.Add(replacement);
        cols.ReplaceNodes(cols.Elements().OrderBy(item => (int)item.Attribute("min")!).ToArray());
        return document.WithXml(sheet.Part, xml);
    }

    public static XlsxObjectDocument Worksheet(XlsxObjectDocument document, string action, string id, string? name = null, int position = 0)
    {
        Guard(document);
        var info = XlsxWorkbookExtractor.Extract(document);
        var book = document.ReadXml("xl/workbook.xml");
        var sheets = book.Root!.Element(S + "sheets")!;
        var oldIds = sheets.Elements().Select(item => (string)item.Attribute("sheetId")!).ToArray();
        var sheet = sheets.Elements().FirstOrDefault(item => (string?)item.Attribute("sheetId") == id);
        if (action is "rename-sheet" or "delete-sheet" && (info.Sheets.Any(item => item.Cells.Any(cell => cell.Formula is not null)) || book.Descendants(S + "definedName").Any()))
        {
            throw new NotSupportedException("Rename/delete with formulas or named ranges is not supported yet; use Excel to update references safely.");
        }
        if (action is "add-sheet" or "rename-sheet")
        {
            if (string.IsNullOrWhiteSpace(name) || name.Length > 31 || name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0 || name.StartsWith('\'') || name.EndsWith('\'')
                || info.Sheets.Any(item => item.Id != (action == "rename-sheet" ? id : "") && item.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException("Use a unique worksheet name of 1–31 characters without []:*?/\\ or surrounding apostrophes.");
            }
        }
        if (action == "add-sheet")
        {
            if (info.Sheets.Count >= 256)
            {
                throw new NotSupportedException("A workbook can contain up to 256 worksheets here.");
            }
            var nextId = info.Sheets.Max(item => int.Parse(item.Id)) + 1;
            var token = "sheet" + Guid.NewGuid().ToString("N");
            var part = "xl/worksheets/sheet-" + token + ".xml";
            sheets.Add(new XElement(S + "sheet", new XAttribute("name", name!), new XAttribute("sheetId", nextId), new XAttribute(R + "id", token)));
            var relations = document.ReadXml("xl/_rels/workbook.xml.rels");
            relations.Root!.Add(new XElement(Rel + "Relationship", new XAttribute("Id", token), new XAttribute("Type", R.NamespaceName + "/worksheet"), new XAttribute("Target", "worksheets/sheet-" + token + ".xml")));
            var types = document.ReadXml("[Content_Types].xml");
            types.Root!.Add(new XElement(types.Root.Name.Namespace + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
            document = document.WithXml(part, new XDocument(new XElement(S + "worksheet", new XElement(S + "sheetData"))))
                .WithXml("xl/_rels/workbook.xml.rels", relations).WithXml("[Content_Types].xml", types);
        }
        else if (sheet is null)
        {
            throw new ArgumentException("Worksheet not found.");
        }
        else if (action == "rename-sheet")
        {
            sheet.SetAttributeValue("name", name);
        }
        else if (action == "delete-sheet")
        {
            if (info.Sheets.Count(item => !item.Hidden) <= 1 && (string?)sheet.Attribute("state") is not ("hidden" or "veryHidden"))
            {
                throw new ArgumentException("Keep at least one visible worksheet.");
            }
            sheet.Remove();
        }
        else if (action == "move-sheet")
        {
            if (position < 0 || position >= info.Sheets.Count || info.Sheets.Any(item => item.Cells.Any(cell => cell.Formula?.Contains(':') == true && cell.Formula.Contains('!'))))
            {
                throw new ArgumentException("Invalid position, or workbook contains order-sensitive 3D formulas.");
            }
            var items = sheets.Elements().ToList();
            items.Remove(sheet);
            items.Insert(position, sheet);
            sheets.ReplaceNodes(items);
        }
        else
        {
            throw new ArgumentException("Unknown worksheet operation.");
        }
        var newIds = sheets.Elements().Select(item => (string)item.Attribute("sheetId")!).ToArray();
        int Remap(int oldIndex)
        {
            var index = oldIndex >= 0 && oldIndex < oldIds.Length ? Array.IndexOf(newIds, oldIds[oldIndex]) : -1;
            return index >= 0 ? index : 0;
        }
        foreach (var nameNode in book.Descendants(S + "definedName").Where(item => item.Attribute("localSheetId") is not null))
        {
            nameNode.SetAttributeValue("localSheetId", Remap((int)nameNode.Attribute("localSheetId")!));
        }
        foreach (var view in book.Descendants(S + "workbookView"))
        {
            var active = Remap((int?)view.Attribute("activeTab") ?? 0);
            if ((string?)sheets.Elements().ElementAt(active).Attribute("state") is "hidden" or "veryHidden")
            {
                active = sheets.Elements().TakeWhile(item => (string?)item.Attribute("state") is "hidden" or "veryHidden").Count();
            }
            view.SetAttributeValue("activeTab", active);
            view.SetAttributeValue("firstSheet", 0);
        }
        return Recalculate(document.WithXml("xl/workbook.xml", book));
    }
}
