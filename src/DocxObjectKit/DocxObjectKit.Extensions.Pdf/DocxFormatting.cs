using System.Xml.Linq;
using System.Globalization;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private XDocument? numbering;
    private readonly Dictionary<string, int[]> counters = new();
    private readonly Dictionary<string, bool[]> initializedCounters = new();
    private XDocument? theme;

    private static void MergeProperty(XElement target, XElement source)
    {
        var existing = target.Element(source.Name);
        if (existing is null || source.Name.LocalName is "b" or "i" or "u" or "vanish" or "pageBreakBefore" or "keepNext" or "keepLines")
        {
            existing?.Remove();
            target.Add(new XElement(source));
            return;
        }
        if (source.Name == W + "tabs")
        {
            foreach (var tab in source.Elements(W + "tab"))
            {
                existing.Elements(W + "tab").Where(item => Value(item, "pos") == Value(tab, "pos")).Remove();
                existing.Add(new XElement(tab));
            }
            return;
        }
        foreach (var attribute in source.Attributes())
        {
            existing.SetAttributeValue(attribute.Name, attribute.Value);
        }
        foreach (var child in source.Elements())
        {
            MergeProperty(existing, child);
        }
    }

    private string FontFamily(XElement style)
    {
        var fonts = style.Element(W + "rFonts");
        var explicitFont = Value(fonts, "ascii") ?? Value(fonts, "hAnsi");
        if (explicitFont is not null)
        {
            return explicitFont;
        }
        var themeFont = Value(fonts, "asciiTheme") ?? Value(fonts, "hAnsiTheme");
        if (themeFont is not null && document.PartNames.Contains("word/theme/theme1.xml"))
        {
            theme ??= document.ReadXml("word/theme/theme1.xml");
            return (string?)theme.Descendants(A + (themeFont.StartsWith("major", StringComparison.Ordinal) ? "majorFont" : "minorFont")).FirstOrDefault()?.Element(A + "latin")?.Attribute("typeface") ?? "Calibri";
        }
        return "Calibri";
    }

    private XElement? NumberingLevel(XElement properties, int? requestedLevel = null)
    {
        var numProperties = properties.Element(W + "numPr");
        var id = Value(numProperties?.Element(W + "numId"));
        if (id is null or "0" || !document.PartNames.Contains("word/numbering.xml"))
        {
            return null;
        }
        numbering ??= document.ReadXml("word/numbering.xml");
        var num = numbering.Root!.Elements(W + "num").FirstOrDefault(item => Value(item, "numId") == id);
        var level = requestedLevel?.ToString(CultureInfo.InvariantCulture) ?? Value(numProperties?.Element(W + "ilvl")) ?? "0";
        var levelOverride = num?.Elements(W + "lvlOverride").FirstOrDefault(item => Value(item, "ilvl") == level);
        var definition = levelOverride?.Element(W + "lvl") ?? numbering.Root.Elements(W + "abstractNum").FirstOrDefault(item => Value(item, "abstractNumId") == Value(num?.Element(W + "abstractNumId")))?.Elements(W + "lvl").FirstOrDefault(item => Value(item, "ilvl") == level);
        if (definition is null)
        {
            if (requestedLevel is null)
            {
                warnings.Add("A numbering definition is missing; the marker is omitted.");
            }
            return null;
        }
        var copy = new XElement(definition);
        if (levelOverride?.Element(W + "startOverride") is { } start)
        {
            copy.Elements(W + "start").Remove();
            copy.Add(new XElement(W + "start", new XAttribute(W + "val", Value(start) ?? "1")));
        }
        return copy;
    }

    private string ListMarker(XElement properties, XElement definition)
    {
        var id = Value(properties.Element(W + "numPr")?.Element(W + "numId"))!;
        var level = Math.Clamp((int)Number(definition, "ilvl"), 0, 8);
        if (!counters.TryGetValue(id, out var values))
        {
            values = new int[9];
            counters[id] = values;
            initializedCounters[id] = new bool[9];
        }
        var initialized = initializedCounters[id];
        values[level] = initialized[level] ? values[level] + 1 : (int)Number(definition.Element(W + "start"), "val", 1);
        initialized[level] = true;
        for (var deeper = level + 1; deeper < values.Length; deeper++)
        {
            var deeperDefinition = NumberingLevel(properties, deeper);
            var restart = (int)Number(deeperDefinition?.Element(W + "lvlRestart"), "val", deeper);
            if (restart > 0 && level < restart)
            {
                initialized[deeper] = false;
            }
        }
        var format = Value(definition.Element(W + "numFmt")) ?? "decimal";
        if (format == "none")
        {
            return "";
        }
        if (format == "bullet")
        {
            var bullet = Value(definition.Element(W + "lvlText")) ?? "\u2022";
            return bullet.Any(character => character >= 0xe000 && character <= 0xf8ff) ? "\u2022" : bullet;
        }
        var marker = Value(definition.Element(W + "lvlText")) ?? "%1.";
        for (var index = 0; index <= level; index++)
        {
            var referencedLevel = index == level ? definition : NumberingLevel(properties, index);
            var value = !initialized[index] ? (int)Number(referencedLevel?.Element(W + "start"), "val", 1) : values[index];
            var referencedFormat = Value(referencedLevel?.Element(W + "numFmt")) ?? "decimal";
            var text = FormatNumber(value, Enabled(definition, "isLgl") ? "decimal" : referencedFormat);
            marker = marker.Replace("%" + (index + 1), text, StringComparison.Ordinal);
        }
        if (format is not ("decimal" or "decimalZero" or "lowerLetter" or "upperLetter" or "lowerRoman" or "upperRoman"))
        {
            warnings.Add("Numbering format '" + format + "' uses decimal markers.");
        }
        return marker;
    }
    private static string FormatNumber(int value, string format)
    {
        if (format is "lowerLetter" or "upperLetter" && value > 0)
        {
            var text = "";
            while (value > 0)
            {
                value--;
                text = (char)('A' + value % 26) + text;
                value /= 26;
            }
            return format == "lowerLetter" ? text.ToLowerInvariant() : text;
        }
        if (format is "lowerRoman" or "upperRoman" && value is > 0 and < 4000)
        {
            var text = "";
            (int Value, string Text)[] symbols = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
            foreach (var symbol in symbols)
            {
                while (value >= symbol.Value)
                {
                    text += symbol.Text;
                    value -= symbol.Value;
                }
            }
            return format == "lowerRoman" ? text.ToLowerInvariant() : text;
        }
        return value.ToString(format == "decimalZero" ? "D2" : "D", CultureInfo.InvariantCulture);
    }

}
