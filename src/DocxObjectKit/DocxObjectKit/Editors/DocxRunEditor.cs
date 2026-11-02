using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Editors;

internal static class DocxRunEditor
{
    // Preserve unchanged XML and formatting on both sides of the edited range.
    // Inserted text inherits the run at the insertion point.
    public static void Replace(XElement paragraph, string replacement)
    {
        if (paragraph.Descendants().Any(node => node.Name == W + "fldChar" || node.Name == W + "fldSimple"
            || node.Name == W + "instrText" || node.Name == W + "ins" || node.Name == W + "del"
            || node.Name == W + "drawing" || node.Name == W + "pict" || node.Name == W + "object"
            || node.Name == W + "footnoteReference" || node.Name == W + "endnoteReference"
            || node.Name == W + "br" && (string?)node.Attribute(W + "type") is "page" or "column"))
        {
            throw new NotSupportedException("Text replacement cannot modify paragraphs containing fields, revisions, notes, drawings or page/column breaks.");
        }
        var nodes = paragraph.Descendants().Where(node => node.Parent?.Name == W + "r"
            && (node.Name == W + "t" || node.Name == W + "tab" || node.Name == W + "br" || node.Name == W + "cr")).ToArray();
        string Text(XElement node) => node.Name == W + "t" ? node.Value : node.Name == W + "tab" ? "\t" : "\n";
        var original = string.Concat(nodes.Select(Text));
        if (original == replacement)
        {
            return;
        }
        var prefix = 0;
        while (prefix < original.Length && prefix < replacement.Length && original[prefix] == replacement[prefix])
        {
            prefix++;
        }
        var suffix = 0;
        while (suffix < original.Length - prefix && suffix < replacement.Length - prefix
            && original[original.Length - 1 - suffix] == replacement[replacement.Length - 1 - suffix])
        {
            suffix++;
        }
        // Never split a surrogate pair at the edit boundary.
        if (prefix > 0 && prefix < original.Length && char.IsHighSurrogate(original[prefix - 1]))
        {
            prefix--;
        }
        if (suffix > 0 && char.IsLowSurrogate(original[original.Length - suffix]))
        {
            suffix--;
        }
        var end = original.Length - suffix;
        var inserted = replacement.Substring(prefix, replacement.Length - prefix - suffix);
        var position = 0;
        var applied = false;
        foreach (var node in nodes)
        {
            var value = Text(node);
            var start = position;
            position += value.Length;
            if (position < prefix || start > end)
            {
                continue;
            }
            var keepBefore = Math.Clamp(prefix - start, 0, value.Length);
            var keepAfter = Math.Clamp(end - start, 0, value.Length);
            var updated = value[..keepBefore] + (applied ? "" : inserted) + value[keepAfter..];
            applied = true;
            if (updated != value)
            {
                node.ReplaceWith(Content(updated));
            }
        }
        if (!applied)
        {
            var run = paragraph.Elements(W + "r").LastOrDefault();
            if (run is null)
            {
                run = new XElement(W + "r");
                paragraph.Add(run);
            }
            run.Add(Content(inserted));
        }
    }

    private static IEnumerable<XElement> Content(string text)
    {
        var start = 0;
        for (var index = 0; index <= text.Length; index++)
        {
            if (index != text.Length && text[index] is not ('\n' or '\t'))
            {
                continue;
            }
            if (index > start)
            {
                yield return new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text[start..index]);
            }
            if (index < text.Length)
            {
                yield return new XElement(W + (text[index] == '\n' ? "br" : "tab"));
            }
            start = index + 1;
        }
    }
}
