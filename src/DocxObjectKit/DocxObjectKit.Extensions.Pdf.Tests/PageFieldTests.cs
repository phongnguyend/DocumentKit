using System.Xml.Linq;
using DocxObjectKit.Extensions.Pdf;
using DocxObjectKit.Models;
using static DocxObjectKit.Models.DocxNamespaces;

internal static class PageFieldTests
{
    public static void Run(DocxObjectDocument source, Action<bool, string> check)
    {
        XElement P(string text) => new(W + "p", new XElement(W + "r", new XElement(W + "t", text)));
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
        var toc = new XElement(W + "sdt", new XElement(W + "sdtPr", new XElement(W + "docPartObj", new XElement(W + "docPartGallery", new XAttribute(W + "val", "Table of Contents")))), new XElement(W + "sdtContent", P("Old entry")));
        var heading = P("Numbered heading");
        heading.AddFirst(new XElement(W + "pPr", new XElement(W + "outlineLvl", new XAttribute(W + "val", 0)), new XElement(W + "pageBreakBefore"), new XElement(W + "numPr", new XElement(W + "ilvl", new XAttribute(W + "val", 0)), new XElement(W + "numId", new XAttribute(W + "val", 1)))));
        var numbered = Document(toc, heading).WithXml("word/numbering.xml", new XDocument(new XElement(W + "numbering",
            new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", 0), new XElement(W + "lvl", new XAttribute(W + "ilvl", 0), new XElement(W + "start", new XAttribute(W + "val", 3)), new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")), new XElement(W + "lvlText", new XAttribute(W + "val", "%1.")))),
            new XElement(W + "num", new XAttribute(W + "numId", 1), new XElement(W + "abstractNumId", new XAttribute(W + "val", 0))))));
        var doc = numbered.ReadXml("word/document.xml");
        doc.Descendants(W + "sectPr").Single().Add(new XElement(W + "pgNumType", new XAttribute(W + "fmt", "lowerRoman"), new XAttribute(W + "start", 4)));
        numbered = numbered.WithXml("word/document.xml", doc);
        var refreshed = DocxTocUpdater.Update(numbered);
        var tocContent = refreshed.ReadXml("word/document.xml").Descendants(W + "sdtContent").Single();
        check(tocContent.Descendants(W + "t").Any(text => text.Value == "3. Numbered heading"), "TOC refresh must include automatic heading numbering.");
        check(tocContent.Descendants(W + "fldSimple").Single().Value == "v", "TOC page references must honor Roman page labels and section restarts.");
        check(DocxTocUpdater.Update(refreshed).ReadXml("word/document.xml").ToString() == refreshed.ReadXml("word/document.xml").ToString(), "Numbered TOCs must remain idempotent.");
        var second = P("Second page");
        second.AddFirst(new XElement(W + "pPr", new XElement(W + "pageBreakBefore")));
        var pages = Document(P("First page"), second);
        doc = pages.ReadXml("word/document.xml");
        doc.Descendants(W + "sectPr").Single().Add(new XElement(W + "pgNumType", new XAttribute(W + "fmt", "lowerRoman"), new XAttribute(W + "start", 4)), new XElement(W + "footerReference", new XAttribute(W + "type", "default"), new XAttribute(R + "id", "pageFields")));
        var rels = pages.ReadXml("word/_rels/document.xml.rels");
        rels.Root!.Add(new XElement(rels.Root.Name.Namespace + "Relationship", new XAttribute("Id", "pageFields"), new XAttribute("Target", "pageFields.xml"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")));
        XElement Run(XElement content) => new(W + "r", content);
        var footer = new XDocument(new XElement(W + "ftr", new XElement(W + "p",
            new XElement(W + "fldSimple", new XAttribute(W + "instr", " page "), Run(new XElement(W + "t", "old"))),
            Run(new XElement(W + "t", " / ")),
            Run(new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))), Run(new XElement(W + "instrText", " NUM")), Run(new XElement(W + "instrText", "PAGES ")),
            Run(new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))), Run(new XElement(W + "t", "old")), Run(new XElement(W + "t", "cache")), Run(new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))),
            Run(new XElement(W + "t", " / ")),
            new XElement(W + "fldSimple", new XAttribute(W + "instr", " PAGE \\* ALPHABETIC "), Run(new XElement(W + "t", "old"))),
            new XElement(W + "fldSimple", new XAttribute(W + "instr", " DATE "), Run(new XElement(W + "t", "Preserved"))))));
        pages = pages.WithXml("word/document.xml", doc).WithXml("word/_rels/document.xml.rels", rels).WithXml("word/pageFields.xml", footer);
        var output = Layout(pages);
        check(output.PageLabels![1] == "iv" && output.PageLabels[2] == "v", "Layout must expose formatted labels separately from physical page numbers.");
        check(output.Pages[0].Runs.Any(run => run.Text == "iv") && output.Pages[1].Runs.Any(run => run.Text == "v"), "Footer PAGE must inherit section number format.");
        check(output.Pages.All(page => page.Runs.Any(run => run.Text == "2") && page.Runs.All(run => run.Text != "cache")), "Split complex NUMPAGES instructions must update and clear every cached fragment.");
        check(output.Pages[0].Runs.Any(run => run.Text == "D") && output.Pages[1].Runs.Any(run => run.Text == "E"), "Explicit field number formatting must override section format.");
        check(output.Pages.All(page => page.Runs.Any(run => run.Text == "Preserved")), "Unrelated field caches must be preserved.");
        var spaced = P("LineOne");
        spaced.AddFirst(new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "line", 360), new XAttribute(W + "lineRule", "atLeast"))));
        spaced.Add(Run(new XElement(W + "br")), Run(new XElement(W + "t", "LineTwo")));
        var spacedRuns = Layout(Document(spaced)).Pages[0].Runs.Where(run => run.Text.StartsWith("Line")).ToArray();
        check(Math.Abs(spacedRuns[1].Baseline - spacedRuns[0].Baseline - 18) < .01, "At-least line spacing uses twips, not an automatic-spacing multiplier.");
        var filler = P("Filler");
        filler.AddFirst(new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "line", 1600), new XAttribute(W + "lineRule", "exact"), new XAttribute(W + "after", 0))));
        var tall = P("TallOne");
        tall.AddFirst(new XElement(W + "pPr", new XElement(W + "spacing", new XAttribute(W + "line", 2000), new XAttribute(W + "lineRule", "exact"), new XAttribute(W + "before", 0))));
        tall.Add(Run(new XElement(W + "br")), Run(new XElement(W + "t", "TallTwo")), Run(new XElement(W + "br")), Run(new XElement(W + "t", "TallThree")));
        var shortFiller = new XElement(filler);
        shortFiller.Descendants(W + "spacing").Single().SetAttributeValue(W + "line", 3000);
        var shortParagraph = new XElement(tall);
        shortParagraph.Descendants(W + "spacing").Single().SetAttributeValue(W + "line", 800);
        var shortLayout = Layout(Document(shortFiller, shortParagraph));
        check(shortLayout.Pages[0].Runs.All(run => !run.Text.StartsWith("Tall")), "A three-line paragraph that fits a full page must not leave a single widow line.");
        var tallLayout = Layout(Document(filler, tall));
        check(tallLayout.Pages[0].Runs.All(run => !run.Text.StartsWith("Tall")) && tallLayout.Pages[1].Runs.Any(run => run.Text == "TallTwo"), "Widow/orphan preflight must measure full line advances rather than glyph ink bounds.");
    }
}

