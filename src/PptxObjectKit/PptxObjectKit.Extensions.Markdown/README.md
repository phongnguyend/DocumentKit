# PptxObjectKit.Extensions.Markdown

Optional Markdown conversion for PptxObjectKit. Reference this package alongside the core
library. Converter types use the `PptxObjectKit.Extensions.Markdown` namespace.
The playground and CLI include this extension. Options and output behavior are unchanged.

```csharp
using PptxObjectKit.Extensions.Markdown;

var markdown = PptxMarkdownConverter.Convert(document);
```

Run the dedicated tests from the repository root:

```powershell
dotnet run --project src/PptxObjectKit/PptxObjectKit.Extensions.Markdown.Tests -c Release
```
