# PdfObjectKit.Extensions.Pptx

Optional editable PowerPoint export for PdfObjectKit. Reference this package when using `PdfPptxConverter`, `PdfPptxOptions`, or `PdfPptxResult`. Their existing namespaces (`PdfObjectKit.Converters`, `PdfObjectKit.Options`, and `PdfObjectKit.Models`) are preserved.

```csharp
using PdfObjectKit.Converters;
using PdfObjectKit.Readers;

var document = PdfReader.Read(File.ReadAllBytes("input.pdf"));
var result = PdfPptxConverter.Convert(document);
File.WriteAllBytes("output.pptx", result.Bytes);
```

The extension references PdfObjectKit and StbImageSharp (managed JPEG decoding for transparent pictures). PdfObjectKit itself has no package dependency on this extension or StbImageSharp. Office and native image libraries are not required.

Export creates editable text boxes, pictures and supported vector shapes. Font substitution, arbitrary clipping and advanced graphics can differ from the PDF. Review `result.Warnings` and the generated slides. See the repository's PdfObjectKit README for options and limitations.

## Tests

Run the extension's console regression suite from the repository root:

```sh
dotnet run --project src/PdfObjectKit/PdfObjectKit.Extensions.Pptx.Tests --configuration Release
```

The suite validates editable text grouping, graphics, transparent images, and Open XML packages. It links the core tests' fixture-builder source without referencing the core test assembly. Open XML validation dependencies live in this test project, not the core test project.
