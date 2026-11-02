using System.Xml.Linq;
using DocxObjectKit.Editors;
using DocxObjectKit.Extensions.Pdf;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

internal static class LayoutFeatureTests
{
    public static void Run(DocxObjectDocument source, Action<bool, string> check)
    {
        XElement P(string text, int height = 20) => new(W + "p", new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "line", height * 20), new XAttribute(W + "lineRule", "exact"), new XAttribute(W + "before", 0), new XAttribute(W + "after", 0))), new XElement(W + "r", new XElement(W + "t", text)));
        DocxObjectDocument Document(params XElement[] blocks)
        {
            var xml = source.ReadXml("word/document.xml");
            xml.Root!.Element(W + "body")!.ReplaceNodes(blocks, new XElement(W + "sectPr", new XElement(W + "pgSz", new XAttribute(W + "w", 7200), new XAttribute(W + "h", 6000)), new XElement(W + "pgMar", new XAttribute(W + "top", 400), new XAttribute(W + "bottom", 400), new XAttribute(W + "left", 400), new XAttribute(W + "right", 400))));
            return source.WithXml("word/document.xml", xml);
        }
        DocxLayoutResult Layout(DocxObjectDocument document)
        {
            using var engine = new DocxLayout(document);
            return engine.Build();
        }
        var lines = P("L0");
        for (var index = 1; index < 5; index++)
        {
            lines.Add(new XElement(W + "r", new XElement(W + "br"), new XElement(W + "t", "L" + index)));
        }
        var widows = Layout(Document(P("Filler", 175), new XElement(lines)));
        check(widows.Pages.Count == 2 && widows.Pages[0].Runs.Count(run => run.Text.StartsWith("L")) == 3 && widows.Pages[1].Runs.Count(run => run.Text.StartsWith("L")) == 2, "Widow control should keep two final lines together.");
        var twoLines = P("First");
        twoLines.Add(new XElement(W + "r", new XElement(W + "br"), new XElement(W + "t", "Second")));
        var orphans = Layout(Document(P("Filler", 235), twoLines));
        check(orphans.Pages[0].Runs.All(run => run.Text != "First") && orphans.Pages[1].Runs.Any(run => run.Text == "First"), "Orphan control should move the first line with its successor.");
        XElement Table(params XElement[] content) => new(W + "tbl", new XElement(W + "tblGrid", new XElement(W + "gridCol", new XAttribute(W + "w", 6000))), new XElement(W + "tr", new XElement(W + "tc", content)));
        var inner = Table(P("Nested content"));
        inner.Descendants(W + "tc").Single().AddFirst(new XElement(W + "tcPr", new XElement(W + "shd", new XAttribute(W + "fill", "ABCDEF"))));
        var nested = Layout(Document(Table(P("Before"), inner, P("After"))));
        check(nested.Pages[0].Runs.Any(run => run.Text == "Nested") && nested.Pages[0].Runs.Last().Text == "After", "Nested table content must remain in cell flow order.");
        check(nested.Pages[0].Rules.Any(rule => rule.Fill == "ABCDEF"), "Nested table cell fills must survive parent composition.");
        check(nested.Pages[0].Runs.All(run => run.X >= 20 && run.X + run.Width <= 340), "Nested tables must use the containing cell width.");
        var row = Table(P("RowOne"), P("RowTwo"), P("RowThree"));
        var split = Layout(Document(P("Filler", 235), new XElement(row)));
        check(split.Pages[0].Runs.Any(run => run.Text == "RowOne") && split.Pages[1].Runs.Any(run => run.Text == "RowTwo"), "Splittable table rows should use remaining page space.");
        check(split.Pages.SelectMany(page => page.Runs).Where(run => run.Text.StartsWith("Row")).Count() == 3, "Row pagination must neither lose nor duplicate text.");
        check(split.Pages.SelectMany(page => page.Runs).Where(run => run.Text.StartsWith("Row")).All(run => run.Y + run.Height <= 280.01), "Table pagination must not cut glyphs at the page edge.");
        row.Element(W + "tr")!.AddFirst(new XElement(W + "trPr", new XElement(W + "cantSplit")));
        var kept = Layout(Document(P("Filler", 235), row));
        check(kept.Pages[0].Runs.All(run => !run.Text.StartsWith("Row")), "cantSplit must move a fitting row intact.");
        var caption = P("Caption");
        caption.Element(W + "pPr")!.Add(new XElement(W + "keepNext"));
        var captionLayout = Layout(Document(P("Filler", 225), caption, Table(P("TableStart"))));
        check(captionLayout.Pages[0].Runs.All(run => run.Text != "Caption") && captionLayout.Pages[1].Runs.Any(run => run.Text == "TableStart"), "Keep-next must bind a caption to the first table row.");
        var reference = P("Reference");
        reference.Add(new XElement(W + "r", new XElement(W + "footnoteReference", new XAttribute(W + "id", 7))));
        var body = Enumerable.Range(0, 18).Select(index => P("Body" + index)).ToArray();
        var note = new XElement(W + "footnote", new XAttribute(W + "id", 7), P("Footnote explanation"));
        note.Element(W + "p")!.AddFirst(new XElement(W + "r", new XElement(W + "footnoteRef")));
        var noteDocument = Document(new[] { reference }.Concat(body).ToArray()).WithXml("word/footnotes.xml", new XDocument(new XElement(W + "footnotes", note)));
        var notes = Layout(noteDocument);
        var notePage = notes.Pages.Single(page => page.Runs.Any(run => run.Text == "Footnote"));
        check(notePage.Runs.Any(run => run.NoteKey == "footnote:7"), "Footnotes must appear on the reference page.");
        var noteTop = notePage.Runs.Where(run => run.Id.StartsWith("note-")).Min(run => run.Y);
        check(notePage.Runs.Where(run => !run.Id.StartsWith("note-")).All(run => run.Y + run.Height <= noteTop), "Footnote space must be reserved before body pagination.");
        check(notePage.TextObjects.All(item => !item.Id.StartsWith("note-")), "Notes must not expose invalid body editor object IDs.");
        var endReference = P("End reference");
        endReference.Add(new XElement(W + "r", new XElement(W + "endnoteReference", new XAttribute(W + "id", 9))));
        var endDocument = Document(endReference).WithXml("word/endnotes.xml", new XDocument(new XElement(W + "endnotes", new XElement(W + "endnote", new XAttribute(W + "id", 9), P("Endnote explanation")))));
        var end = Layout(endDocument);
        check(end.Pages.Count == 2 && end.Pages[1].Runs.Any(run => run.Text == "Endnote"), "Endnotes must render after document content.");
        var toc = new XElement(W + "sdt", new XElement(W + "sdtPr", new XElement(W + "docPartObj", new XElement(W + "docPartGallery", new XAttribute(W + "val", "Table of Contents")))), new XElement(W + "sdtContent", P("Stale contents")));
        var heading = P("Updated heading");
        heading.Element(W + "pPr")!.Add(new XElement(W + "outlineLvl", new XAttribute(W + "val", 0)), new XElement(W + "pageBreakBefore"));
        var tocDocument = DocxTocUpdater.Update(Document(toc, heading));
        var tocXml = tocDocument.ReadXml("word/document.xml");
        var tocContent = tocXml.Descendants(W + "sdtContent").Single();
        check(!tocContent.Value.Contains("Stale") && tocContent.Value.Contains("Updated heading"), "TOC refresh must regenerate heading titles.");
        check(tocContent.Descendants(W + "fldSimple").Single().Descendants(W + "t").Single().Value == "2", "TOC must use recomputed heading pages.");
        check(tocContent.Descendants(W + "hyperlink").Single().Attribute(W + "anchor")!.Value == tocXml.Descendants(W + "bookmarkStart").Single().Attribute(W + "name")!.Value, "TOC links must target persistent heading bookmarks.");
        check(DocxTocUpdater.Update(tocDocument).ReadXml("word/document.xml").ToString() == tocXml.ToString(), "TOC refresh should converge and be idempotent.");
        var rich = new XElement(W + "p", new XElement(W + "r", new XElement(W + "rPr", new XElement(W + "b")), new XElement(W + "t", "Bold ")),
            new XElement(W + "hyperlink", new XAttribute(W + "anchor", "target"), new XElement(W + "r", new XElement(W + "rPr", new XElement(W + "color", new XAttribute(W + "val", "123456"))), new XElement(W + "t", "link"))),
            new XElement(W + "r", new XElement(W + "rPr", new XElement(W + "i")), new XElement(W + "t", " italic")));
        var richDocument = Document(rich);
        var edited = DocxEditor.SetText(richDocument, "p0", "Bold new link italic").ReadXml("word/document.xml");
        check(edited.Descendants(W + "b").Any() && edited.Descendants(W + "i").Any() && edited.Descendants(W + "hyperlink").Single().Value == "link", "Text insertion must preserve unaffected runs and hyperlinks.");
        check(edited.Descendants(W + "r").First().Value == "Bold new ", "Inserted text inherits the insertion run.");
        var multiline = DocxEditor.SetText(richDocument, "p0", "Bold link\n\titalic").ReadXml("word/document.xml");
        check(multiline.Descendants(W + "br").Count() == 1 && multiline.Descendants(W + "tab").Count() == 1 && multiline.Descendants(W + "i").Any(), "Line breaks and tabs must preserve surrounding formatting.");
    }
}
