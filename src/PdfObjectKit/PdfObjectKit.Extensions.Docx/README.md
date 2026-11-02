# PdfObjectKit.Extensions.Docx

Optional PDF-to-Word export for PdfObjectKit. The core library does not depend on this extension.

```csharp
using PdfObjectKit.Converters;
using PdfObjectKit.Readers;

var pdf = PdfReader.Read(File.ReadAllBytes("input.pdf"));
var result = PdfDocxConverter.Convert(pdf, new() { IncludeImages = true, PageBreaks = true });
File.WriteAllBytes("output.docx", result.Bytes);
```

Exports editable reflowable paragraphs, heuristic headings and tables, literal list markers,
and supported JPEG/PNG images (including supported soft masks). References PDF font names
and sizes; no fonts are embedded. PageNumbers accepts 1?100 distinct one-based pages in
output order; null selects all pages. Cancellation is checked between pages and blocks.

Output uses Letter pages with standard margins. Source page breaks are optional; Word
can create additional pages while reflowing. This is not a fixed-layout replica: colors,
vector drawings, image clipping/rotation, vertical table merges, annotations, interactive
fields and metadata are not preserved. Scanned pages remain images unless OCR text is
present. Reading order, headings and tables are inferred and require review. Diagnostics
report unsupported extraction and omitted images. Never use this export as redaction.
