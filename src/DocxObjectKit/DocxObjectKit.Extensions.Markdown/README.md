# DocxObjectKit.Extensions.Markdown

Optional Markdown conversion for DocxObjectKit. Reference this package alongside the core
library. Converter types use the `DocxObjectKit.Extensions.Markdown` namespace.
The playground and CLI include this extension. Options and output behavior are unchanged.

```csharp
using DocxObjectKit.Extensions.Markdown;

var markdown = DocxMarkdownConverter.Convert(document);
```

Run the dedicated tests from the repository root:

```powershell
dotnet run --project src/DocxObjectKit/DocxObjectKit.Extensions.Markdown.Tests -c Release
```
