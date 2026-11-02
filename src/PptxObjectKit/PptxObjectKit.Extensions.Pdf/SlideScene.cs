using System.Globalization;
using System.Xml.Linq;
using PptxObjectKit.Graphs;
using PptxObjectKit.Models;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Extensions.Pdf;

internal sealed record SlideScene(XDocument Xml, string Part, int Number, XDocument? Layout, string? LayoutPart, XDocument? Master, string? MasterPart, XDocument? Theme);

internal sealed partial class SlideRenderer : IDisposable
{
    private readonly PptxObjectDocument document;
    private readonly PptxPackageGraph graph;
    private readonly HashSet<string> warnings;
    private readonly CancellationToken cancellation;
    private SlideScene scene = null!;
    private readonly XElement? defaultTextStyle;

    internal SlideRenderer(PptxObjectDocument document, HashSet<string> warnings, CancellationToken cancellation)
    {
        this.document = document;
        graph = new(document);
        var presentation = document.Xml(graph.PresentationPart());
        defaultTextStyle = D(presentation, "defaultTextStyle").FirstOrDefault();
        if (D(presentation, "embeddedFontLst").Any())
        {
            warnings.Add("Presentation: embedded PPTX fonts are not loaded; host fonts or bundled Noto Sans are used.");
        }
        this.warnings = warnings;
        this.cancellation = cancellation;
    }

    private string? Related(string part, string kind, string? id = null)
    {
        var relation = graph.Relationships(part).FirstOrDefault(item => ((string?)item.Attribute("Type"))?.EndsWith("/" + kind, StringComparison.Ordinal) == true && (id is null || (string?)item.Attribute("Id") == id));
        if (relation is null || (string?)relation.Attribute("TargetMode") == "External")
        {
            return null;
        }
        return Target(part, relation);
    }

    internal SlideScene Load(string part, int number)
    {
        var layout = Related(part, "slideLayout");
        var master = layout is null ? null : Related(layout, "slideMaster");
        var theme = master is null ? null : Related(master, "theme");
        return new(document.Xml(part), part, number, layout is null ? null : document.Xml(layout), layout, master is null ? null : document.Xml(master), master, theme is null ? null : document.Xml(theme));
    }

    private static XElement? C(XContainer? parent, string name) => parent?.Elements().FirstOrDefault(item => item.Name.LocalName == name);

    private static IEnumerable<XElement> D(XContainer? parent, string name) => parent?.Descendants().Where(item => item.Name.LocalName == name) ?? [];

    private static float N(XElement? element, string name, float fallback = 0)
    {
        var value = (string?)element?.Attribute(name);
        if (value is null)
        {
            return fallback;
        }
        if (value is "true" or "false" or "on" or "off")
        {
            return value is "true" or "on" ? 1 : 0;
        }
        if (!float.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result) || !float.IsFinite(result) || Math.Abs(result) > 1e12)
        {
            throw new InvalidDataException($"Invalid DrawingML numeric value on {element?.Name.LocalName}/@{name}: {value}.");
        }
        return result;
    }

    private void Warn(string message) => warnings.Add($"Slide {scene.Number}: {message}");

    private static XElement? Placeholder(XElement shape, XDocument? document)
    {
        var ph = D(shape, "ph").FirstOrDefault();
        if (ph is null)
        {
            return null;
        }
        var shapes = D(document, "sp").Where(item => D(item, "ph").Any()).ToArray();
        return shapes.FirstOrDefault(item => ((string?)D(item, "ph").First().Attribute("idx") ?? "0") == ((string?)ph.Attribute("idx") ?? "0"))
            ?? shapes.FirstOrDefault(item => (string?)D(item, "ph").First().Attribute("type") == (string?)ph.Attribute("type"));
    }

    private static void Merge(XElement? target, XElement? source)
    {
        if (target is null || source is null)
        {
            return;
        }
        foreach (var attribute in source.Attributes())
        {
            if (target.Attribute(attribute.Name) is null)
            {
                target.Add(new XAttribute(attribute));
            }
        }
        var choices = new[] { new[] { "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill" }, new[] { "buNone", "buChar", "buAutoNum" }, new[] { "buFont", "buFontTx" }, new[] { "buClr", "buClrTx" }, new[] { "buSzPct", "buSzPts", "buSzTx" } };
        foreach (var child in source.Elements())
        {
            var exclusive = choices.FirstOrDefault(group => group.Contains(child.Name.LocalName));
            if (child.Name.LocalName == "defRPr" && C(target, "defRPr") is { } existing)
            {
                Merge(existing, child);
            }
            if (C(target, child.Name.LocalName) is null && !(exclusive?.Any(name => C(target, name) is not null) ?? false))
            {
                target.Add(new XElement(child));
            }
        }
    }

    private XElement Prepare(XElement shape)
    {
        var copy = new XElement(shape);
        var inherited = Placeholder(shape, scene.Layout);
        var master = Placeholder(inherited ?? shape, scene.Master);
        var props = C(copy, "spPr");
        Merge(props, C(inherited, "spPr"));
        Merge(props, C(master, "spPr"));
        var body = C(copy, "txBody");
        Merge(C(body, "bodyPr"), C(C(inherited, "txBody"), "bodyPr"));
        Merge(C(body, "bodyPr"), C(C(master, "txBody"), "bodyPr"));
        var type = (string?)D(shape, "ph").FirstOrDefault()?.Attribute("type") ?? "other";
        var styleName = type is "title" or "ctrTitle" ? "titleStyle" : type is "body" or "subTitle" or "obj" ? "bodyStyle" : "otherStyle";
        foreach (var paragraph in body?.Elements(A + "p") ?? [])
        {
            var pPr = paragraph.Element(A + "pPr");
            if (pPr is null)
            {
                pPr = new XElement(A + "pPr");
                paragraph.AddFirst(pPr);
            }
            var level = "lvl" + (N(pPr, "lvl") + 1).ToString(CultureInfo.InvariantCulture) + "pPr";
            foreach (var defaults in new[] { C(C(body, "lstStyle"), level), C(C(C(inherited, "txBody"), "lstStyle"), level), C(C(C(master, "txBody"), "lstStyle"), level) })
            {
                Merge(pPr, defaults);
            }
            // Shape font references outrank generic master/presentation defaults,
            // but explicit run and placeholder formatting still take precedence.
            var localRunDefaults = new XElement(C(pPr, "defRPr") ?? new XElement(A + "defRPr"));
            Merge(pPr, C(D(scene.Master, styleName).FirstOrDefault(), level));
            Merge(pPr, C(defaultTextStyle, level));
            foreach (var run in paragraph.Elements().Where(item => item.Name.LocalName is "r" or "fld"))
            {
                var rPr = run.Element(A + "rPr");
                if (rPr is null)
                {
                    rPr = new XElement(A + "rPr");
                    run.AddFirst(rPr);
                }
                Merge(rPr, localRunDefaults);
                var fontRef = C(C(shape, "style"), "fontRef") ?? C(C(inherited, "style"), "fontRef") ?? C(C(master, "style"), "fontRef");
                if (!rPr.Elements().Any(item => item.Name.LocalName.EndsWith("Fill", StringComparison.Ordinal)) && fontRef?.Elements().FirstOrDefault() is { } color)
                {
                    rPr.Add(new XElement(A + "solidFill", new XElement(color)));
                }
                Merge(rPr, C(pPr, "defRPr"));
                if (C(rPr, "latin") is null)
                {
                    rPr.Add(new XElement(A + "latin", new XAttribute("typeface", (string?)fontRef?.Attribute("idx") == "major" ? "+mj-lt" : "+mn-lt")));
                }
            }
        }
        return copy;
    }
}
