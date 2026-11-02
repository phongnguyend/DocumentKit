using System.IO.Compression;
using System.Text;
using PptxObjectKit.Readers;
using PptxObjectKit.Models;

internal static class RichPresentationFixture
{
    internal static PptxObjectDocument Create(byte[] bytes)
    {
        var document = PptxReader.Read(bytes);
        using var richPackage = new MemoryStream();
        richPackage.Write(bytes);
        using (var archive = new ZipArchive(richPackage, ZipArchiveMode.Update, true))
        {
            void Put(string name, string xml)
            {
                archive.GetEntry(name)?.Delete();
                using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write(xml);
            }
            var slideXml = Encoding.UTF8.GetString(document.GetPart("ppt/slides/slide1.xml"));
            slideXml = slideXml.Replace("</p:spTree>", """
                <p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tr>
                <a:tc><a:txBody><a:p><a:r><a:t>Cell | one</a:t></a:r></a:p></a:txBody></a:tc>
                <a:tc><a:txBody><a:p><a:r><a:t>Cell two</a:t></a:r></a:p></a:txBody></a:tc>
                </a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
                <p:pic><p:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="picture"/></p:blipFill></p:pic></p:spTree>
                """);
            Put("ppt/slides/slide1.xml", slideXml);
            Put("ppt/slides/_rels/slide1.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="notes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
                <Relationship Id="picture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image.png"/>
                </Relationships>
                """);
            Put("ppt/notesSlides/notesSlide1.xml", """
                <p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
                <p:txBody><a:p><a:r><a:t>Speaker details</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>
                """);
            using var image = archive.CreateEntry("ppt/media/image.png").Open();
            image.Write(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jQ1sAAAAASUVORK5CYII="));
        }
        return PptxReader.Read(richPackage.ToArray());
    }
}
