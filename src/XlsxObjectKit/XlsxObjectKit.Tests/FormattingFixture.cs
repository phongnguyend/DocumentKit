using System.Xml.Linq;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Models;
using static XlsxObjectKit.Models.XlsxNamespaces;

internal static class FormattingFixture
{
    public static XlsxObjectDocument Create()
    {
        var document = XlsxEditor.Create();
        var relationships = document.ReadXml("xl/_rels/workbook.xml.rels");
        relationships.Root!.Add(
            new XElement(Rel + "Relationship", new XAttribute("Id", "styles"), new XAttribute("Type", R.NamespaceName + "/styles"), new XAttribute("Target", "styles.xml")),
            new XElement(Rel + "Relationship", new XAttribute("Id", "theme"), new XAttribute("Type", R.NamespaceName + "/theme"), new XAttribute("Target", "theme/theme1.xml")));
        var book = document.ReadXml("xl/workbook.xml");
        book.Root!.AddFirst(new XElement(S + "workbookPr", new XAttribute("date1904", "1")));
        return document.WithXml("xl/workbook.xml", book).WithXml("xl/_rels/workbook.xml.rels", relationships)
            .WithXml("xl/styles.xml", XDocument.Parse($"""
            <styleSheet xmlns="{S}">
              <numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd"/></numFmts>
              <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><i/><strike/><u val="double"/><sz val="20"/><color theme="4"/><name val="Arial"/></font></fonts>
              <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.5"/></patternFill></fill></fills>
              <borders count="2"><border/><border><left style="thin"><color indexed="4"/></left><bottom style="double"><color rgb="FFFF0000"/></bottom></border></borders>
              <cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1"><alignment vertical="center"/></xf></cellStyleXfs>
              <cellXfs count="5">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                <xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
                <xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0"/>
                <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0"/>
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyFont="0" applyFill="0" applyBorder="0"><alignment wrapText="1" horizontal="left" indent="1"/></xf>
              </cellXfs>
            </styleSheet>
            """))
            .WithXml("xl/theme/theme1.xml", XDocument.Parse("""
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Test"><a:accent1><a:srgbClr val="008000"/></a:accent1></a:clrScheme></a:themeElements></a:theme>
            """))
            .WithXml("xl/worksheets/sheet1.xml", XDocument.Parse($"""
            <worksheet xmlns="{S}"><sheetFormatPr defaultRowHeight="22" defaultColWidth="18"/><cols><col min="2" max="2" width="20" customWidth="1" style="4"/></cols><sheetData>
              <row r="1" ht="45" customHeight="1"><c r="A1" s="1"><v>1234.5</v></c></row>
              <row r="2"><c r="A2" s="2"><v>0.125</v></c></row>
              <row r="3"><c r="A3" s="3"><v>0</v></c></row>
              <row r="4" ht="90" customHeight="1" s="4" customFormat="1"><c r="A4" t="inlineStr"><is><t>Wrapped sample text spanning several lines</t></is></c><c r="B4" s="0" t="inlineStr"><is><t>Explicit default</t></is></c></row>
              <row r="5"><c r="B5" t="inlineStr"><is><t>Column style</t></is></c></row>
            </sheetData></worksheet>
            """));
    }
}
