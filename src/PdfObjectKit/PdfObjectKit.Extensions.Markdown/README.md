# PdfObjectKit.Extensions.Markdown

Optional Markdown conversion for PdfObjectKit. Reference this package alongside the core
library. Converter types use the `PdfObjectKit.Extensions.Markdown` namespace.
The core package has no dependency on Markdown rendering.

```csharp
using PdfObjectKit.Extensions.Markdown;

var result = PdfMarkdownConverter.Convert(document);
```

The playground and CLI include this extension. Existing conversion options and
output behavior are unchanged. PDF document analysis and structured models remain
in core for reuse by other exports; `PdfMarkdownOptions` derives from the shared
`PdfAnalysisOptions` type.

Run the dedicated tests from the repository root:

```powershell
dotnet run --project src/PdfObjectKit/PdfObjectKit.Extensions.Markdown.Tests -c Release
```
