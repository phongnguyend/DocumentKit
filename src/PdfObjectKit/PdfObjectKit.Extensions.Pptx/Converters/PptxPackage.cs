using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace PdfObjectKit.Converters;

internal sealed class PptxPackage : IDisposable
{
    internal static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    internal static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    internal static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    private readonly MemoryStream buffer = new();
    private readonly ZipArchive archive;
    private readonly List<XElement> contentTypes = [];
    private readonly List<XElement> slides = [];
    private readonly List<XElement> slideRelations = [];
    private int mediaCount;
    private long uncompressedBytes;

    public PptxPackage()
    {
        archive = new(buffer, ZipArchiveMode.Create, true);
        contentTypes.Add(new(Ct + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")));
        contentTypes.Add(new(Ct + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png")));
        contentTypes.Add(new(Ct + "Default", new XAttribute("Extension", "jpg"), new XAttribute("ContentType", "image/jpeg")));
        Relationships("_rels/.rels", Relation("rId1", "officeDocument", "ppt/presentation.xml"));
    }

    internal static XElement Relation(string id, string type, string target) => new(Rel + "Relationship", new XAttribute("Id", id), new XAttribute("Type", R.NamespaceName + "/" + type), new XAttribute("Target", target));

    internal static XElement Tree() => new(P + "spTree",
        new XElement(P + "nvGrpSpPr", new XElement(P + "cNvPr", new XAttribute("id", 1), new XAttribute("name", "")), new XElement(P + "cNvGrpSpPr"), new XElement(P + "nvPr")),
        new XElement(P + "grpSpPr", new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(A + "ext", new XAttribute("cx", 0), new XAttribute("cy", 0)), new XElement(A + "chOff", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(A + "chExt", new XAttribute("cx", 0), new XAttribute("cy", 0)))));

    internal string AddMedia(byte[] bytes, string extension)
    {
        var name = $"image{++mediaCount}.{extension}";
        Write("ppt/media/" + name, bytes);
        return "../media/" + name;
    }

    internal void AddSlide(XElement tree, List<XElement> images)
    {
        var number = slides.Count + 1;
        Part($"ppt/slides/slide{number}.xml", "slide", new(P + "sld", new XAttribute(XNamespace.Xmlns + "a", A), new XAttribute(XNamespace.Xmlns + "r", R),
            new XElement(P + "cSld", new XElement(P + "bg", new XElement(P + "bgPr", new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", "FFFFFF"))), new XElement(A + "effectLst"))), tree),
            new XElement(P + "clrMapOvr", new XElement(A + "masterClrMapping"))));
        Relationships($"ppt/slides/_rels/slide{number}.xml.rels", [Relation("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"), .. images]);
        slides.Add(new(P + "sldId", new XAttribute("id", 255 + number), new XAttribute(R + "id", "rId" + (number + 1))));
        slideRelations.Add(Relation("rId" + (number + 1), "slide", $"slides/slide{number}.xml"));
    }

    internal byte[] Finish(double width, double height)
    {
        Part("ppt/presentation.xml", "presentation.main", new(P + "presentation", new XAttribute(XNamespace.Xmlns + "a", A), new XAttribute(XNamespace.Xmlns + "r", R),
            new XElement(P + "sldMasterIdLst", new XElement(P + "sldMasterId", new XAttribute("id", 2147483648L), new XAttribute(R + "id", "rId1"))),
            new XElement(P + "sldIdLst", slides), new XElement(P + "sldSz", new XAttribute("cx", Emu(width)), new XAttribute("cy", Emu(height))),
            new XElement(P + "notesSz", new XAttribute("cx", 6858000), new XAttribute("cy", 9144000))));
        Relationships("ppt/_rels/presentation.xml.rels", [Relation("rId1", "slideMaster", "slideMasters/slideMaster1.xml"), .. slideRelations]);
        var map = new XElement(P + "clrMap");
        foreach (var (name, value) in new[] { ("bg1", "lt1"), ("tx1", "dk1"), ("bg2", "lt2"), ("tx2", "dk2"), ("accent1", "accent1"), ("accent2", "accent2"), ("accent3", "accent3"), ("accent4", "accent4"), ("accent5", "accent5"), ("accent6", "accent6"), ("hlink", "hlink"), ("folHlink", "folHlink") })
        {
            map.Add(new XAttribute(name, value));
        }
        Part("ppt/slideMasters/slideMaster1.xml", "slideMaster", new(P + "sldMaster", new XElement(P + "cSld", Tree()), map,
            new XElement(P + "sldLayoutIdLst", new XElement(P + "sldLayoutId", new XAttribute("id", 2147483649L), new XAttribute(R + "id", "rId1"))),
            new XElement(P + "txStyles", new XElement(P + "titleStyle"), new XElement(P + "bodyStyle"), new XElement(P + "otherStyle"))));
        Relationships("ppt/slideMasters/_rels/slideMaster1.xml.rels", Relation("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"), Relation("rId2", "theme", "../theme/theme1.xml"));
        Part("ppt/slideLayouts/slideLayout1.xml", "slideLayout", new(P + "sldLayout", new XAttribute("type", "blank"), new XAttribute("preserve", "1"),
            new XElement(P + "cSld", Tree()), new XElement(P + "clrMapOvr", new XElement(A + "masterClrMapping"))));
        Relationships("ppt/slideLayouts/_rels/slideLayout1.xml.rels", Relation("rId1", "slideMaster", "../slideMasters/slideMaster1.xml"));
        var colors = new XElement(A + "clrScheme", new XAttribute("name", "PDF"));
        var names = new[] { "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
        var values = new[] { "000000", "FFFFFF", "222222", "EEEEEE", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72" };
        for (var i = 0; i < names.Length; i++)
        {
            colors.Add(new XElement(A + names[i], new XElement(A + "srgbClr", new XAttribute("val", values[i]))));
        }
        XElement Font(string type) => new(A + type, new XElement(A + "latin", new XAttribute("typeface", "Arial")), new XElement(A + "ea", new XAttribute("typeface", "")), new XElement(A + "cs", new XAttribute("typeface", "")));
        XElement Fill() => new(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "phClr")));
        var formats = new XElement(A + "fmtScheme", new XAttribute("name", "PDF"),
            new XElement(A + "fillStyleLst", Enumerable.Range(0, 3).Select(_ => Fill())),
            new XElement(A + "lnStyleLst", Enumerable.Range(0, 3).Select(_ => new XElement(A + "ln", new XAttribute("w", 12700), Fill(), new XElement(A + "prstDash", new XAttribute("val", "solid"))))),
            new XElement(A + "effectStyleLst", Enumerable.Range(0, 3).Select(_ => new XElement(A + "effectStyle", new XElement(A + "effectLst")))),
            new XElement(A + "bgFillStyleLst", Enumerable.Range(0, 3).Select(_ => Fill())));
        contentTypes.Add(new(Ct + "Override", new XAttribute("PartName", "/ppt/theme/theme1.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml")));
        Xml("ppt/theme/theme1.xml", new(A + "theme", new XAttribute("name", "PDF"), new XElement(A + "themeElements", colors, new XElement(A + "fontScheme", new XAttribute("name", "PDF"), Font("majorFont"), Font("minorFont")), formats)));
        Xml("[Content_Types].xml", new(Ct + "Types", contentTypes));
        archive.Dispose();
        return buffer.ToArray();
    }

    internal static long Emu(double points) => checked((long)Math.Round(points * 12700));

    private void Part(string name, string type, XElement xml)
    {
        contentTypes.Add(new(Ct + "Override", new XAttribute("PartName", "/" + name), new XAttribute("ContentType", $"application/vnd.openxmlformats-officedocument.presentationml.{type}+xml")));
        Xml(name, xml);
    }

    private void Relationships(string name, params XElement[] relations) => Xml(name, new(Rel + "Relationships", relations));

    private void Xml(string name, XElement element) => Write(name, Encoding.UTF8.GetBytes(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), element).ToString()));

    private void Write(string name, byte[] bytes)
    {
        uncompressedBytes += bytes.LongLength;
        if (uncompressedBytes > 128L * 1024 * 1024)
        {
            throw new NotSupportedException("PowerPoint export exceeds the 128 MB uncompressed package limit. Export fewer pages.");
        }
        using var stream = archive.CreateEntry(name, CompressionLevel.Optimal).Open();
        stream.Write(bytes);
    }

    public void Dispose()
    {
        archive.Dispose();
        buffer.Dispose();
    }
}
