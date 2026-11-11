# PptxObjectKit

A first version of a .NET 11 library and playground for opening, inspecting, editing, and saving `.pptx` presentations. It is independent of PdfObjectKit and LibreOffice. The core library uses ZIP and XML APIs from .NET and has no NuGet dependencies. The optional PDF extension uses SkiaSharp for native PDF output.

## Projects

| Project | Purpose |
| --- | --- |
| `PptxObjectKit` | Package models, reader/writer, slide extraction, editors, Markdown conversion |
| `PptxObjectKit.Application` | Shared edit dispatch for API and CLI |
| `PptxObjectKit.Cli` | Inspect, rewrite, edit and convert to Markdown/PDF |
| `PptxObjectKit.Extensions.Pdf` | Native DrawingML-to-PDF rendering with selectable text |
| `PptxObjectKit.Extensions.Pdf.Tests` | PDF output regression tests using PdfObjectKit to read the result |
| `PptxObjectKit.Playground.Api` | Stateless ASP.NET Core minimal API |
| `PptxObjectKit.Playground.Ui` | React 19 / TypeScript playground |
| `PptxObjectKit.Tests` | Console regression suite and a small synthetic presentation |

## Playground

The shell follows PdfObjectKit: shared branding and Open action above browser-style tabs, a document toolbar with Slides/Document menus and download/history actions, thumbnails on the left, an inspector on the right, and slide navigation at the bottom.

- Open or drop one or more `.pptx` files, up to 50 MB each, into separate browser-style document tabs. Add, close, or drag tabs to reorder; arrow keys navigate the tab strip. Each tab preserves its slide selection, edits, inspector, and undo/redo history. Closing a tab with edits not yet downloaded requires in-app confirmation.
- Navigate using slide thumbnails with slide numbers and titles. Previews load as the sidebar scrolls and refresh after edits. Slide order follows presentation relationships, including decks whose filenames do not match their order.
- View an experimental canvas preview. Select text (the default) copies words; Select objects lets you click an editable text object or image on the slide and edit it in the inspector. The icon switcher follows PdfObjectKit.
- Select a text object in the inspector and replace its text. The first run's formatting in each paragraph is retained; mixed formatting within that paragraph is simplified. The UI limits replacement text to 4,000 characters.
- Add images using the same picker as PdfObjectKit: upload/drop raster files (up to 5 MB), draw a transparent image with brush/color/undo controls, or generate a QR code locally. Preview before applying, adjust placement dimensions, or use the source pixel size. WebP/GIF/BMP/AVIF are normalized to PNG. Both insertion and replacement support all three sources. Select images on the slide in object mode or from the inspector, download the original image, replace an instance, drag the selected image to move it or use its eight resize handles (release to save; Escape cancels), edit its position/size in points, or delete it with confirmation. Image changes participate in undo/redo. Replacement preserves position, crop, and rotation and uses a new resource so shared images are unaffected. Grouped images support replacement/download/deletion; their position/size controls are disabled. Inherited layout/master images are not editable in the slide inspector. Deletion retains unreferenced media and is not sanitization.
- Recognize text in a selected image with **Recognize text (OCR)**. English, Vietnamese, and combined recognition run locally using bundled language files. Review on a canvas, select/move/edit individual lines, switch to plain text, copy, or retry. **Apply to slide** creates editable text boxes at the image placement, either transparent while keeping the image or visible while replacing it. Undo/redo and Markdown extraction include OCR edits. Fonts/layout are approximate; cropped images cannot yet be applied (copying recognized text remains available). Retained media is not sanitized.
- In **Select objects** mode, drag an ungrouped text box to move it or use its eight handles to resize. Release saves, Escape cancels, and undo/redo are supported. Layout/master-inherited placement is resolved and written as a slide-local override; font sizes, text, rotation and wrapping settings are retained. Grouped text remains editable but cannot yet be dragged/resized.
- Reorder slides by dragging the slide list or using Move earlier / Move later.
- Use **Slides > Duplicate slide** to insert an independent copy after the current slide, including speaker notes and owned package parts. **Blank slide after current** and **Blank slide at end** insert empty white slides at the presentation size. The new slide opens automatically; all three actions support undo/redo.
- Delete slides using an in-app confirmation; deleting the last slide is rejected.
- Undo/redo up to ten recent edits and download the current PPTX.
- View Markdown as rendered content or source, include/exclude notes and pictures, reapply conversion options, copy or download the result.

Files are processed in memory by the API and are not saved there. Changes are kept in the browser session until downloaded. Opening more files creates new tabs. Download each edited presentation to keep its changes before leaving the page.

### Run locally

Requires .NET 11 and Node.js 22.12+.

From this directory, in separate terminals:

```powershell
dotnet run --project PptxObjectKit.Playground.Api --urls http://localhost:5090
```

```powershell
cd PptxObjectKit.Playground.Ui
npm ci
npm run dev
```

Open http://localhost:5180. Copy `.env.example` to `.env` in the UI folder to change the development API URL or UI port.

### Docker

```powershell
docker compose up --build
```

Open http://localhost:8124. API is exposed at http://localhost:5090. Build contexts are this solution directory for .NET and the UI directory for Nginx/Vite. No Office installation is required.

## Library

```csharp
using PptxObjectKit.Readers;
using PptxObjectKit.Writers;
using PptxObjectKit.Extractors;
using PptxObjectKit.Editors;
using PptxObjectKit.Converters;

var document = PptxReader.Read(File.ReadAllBytes("input.pptx"));
var inspection = PptxSlideExtractor.Extract(document);
var markdown = PptxMarkdownConverter.Convert(document,
    new PptxMarkdownOptions(IncludeNotes: true, IncludeImages: true));
var reordered = PptxSlideEditor.Reorder(document, [2, 1]);
File.WriteAllBytes("output.pptx", PptxWriter.Write(reordered));
```

Editors return a new document and leave the input untouched. Slide numbers are one-based. Text object IDs come from `PptxSlideExtractor`. Unchanged XML/media parts retain their exact uncompressed bytes; ZIP container bytes and compression may change.

Markdown follows slide and shape order, includes paragraphs and explicit bullets, basic tables, PNG/JPEG images embedded as data URIs, and optional speaker notes. It is not a reconstruction of visual reading order or full rich-text formatting. Merged table cells, charts, SmartArt and embedded document contents are not converted.

## CLI

```powershell
dotnet run --project PptxObjectKit.Cli -- inspect input.pptx inspection.json
dotnet run --project PptxObjectKit.Cli -- to-markdown input.pptx output.md
dotnet run --project PptxObjectKit.Cli -- rewrite input.pptx output.pptx
dotnet run --project PptxObjectKit.Cli -- edit input.pptx output.pptx edit.json
```

The output must not already exist. Example `edit.json`:

```json
{ "kind": "text", "slideNumber": 1, "objectId": "2", "text": "Updated title" }
```

Other edits: `{ "kind": "reorder", "order": [2, 1] }` and `{ "kind": "delete-slide", "slideNumber": 1 }`.

## Native PDF conversion

Choose **Document > Export PDF** in the playground. Export all slides or the current slide, optionally exclude hidden slides, review conversion notes, preview the generated PDF, and download it. The current edited presentation is exported; the source PPTX is unchanged. Changing options clears the old result.

This conversion runs entirely in .NET using a DrawingML renderer and SkiaSharp's PDF writer. It does not invoke PowerPoint, LibreOffice, or a browser. Text is written as text with embedded fonts, including inherited solid and gradient fills; supported shapes and SVG paths remain vectors, and pictures remain images. PNG diagrams retain their source pixels using lossless PDF image encoding rather than low-quality JPEG recompression.

```csharp
using PptxObjectKit.Readers;
using PptxObjectKit.Extensions.Pdf;

var document = PptxReader.Read(File.ReadAllBytes("input.pptx"));
var result = PptxPdfConverter.Convert(document, new PptxPdfOptions
{
    Slides = [1, 3, 2], // Omit for all slides; numbers are one-based.
    IncludeHiddenSlides = false,
    FailOnWarnings = false
});
File.WriteAllBytes("output.pdf", result.Bytes);
foreach (var warning in result.Warnings)
{
    Console.WriteLine(warning);
}
```

```powershell
dotnet run --project PptxObjectKit.Cli -- to-pdf input.pptx output.pdf
# Optional JSON: { "slides": [1, 3], "includeHiddenSlides": false, "failOnWarnings": true }
dotnet run --project PptxObjectKit.Cli -- to-pdf input.pptx output.pdf options.json
```

The CLI refuses to overwrite an existing output and writes conversion notes to stderr. `FailOnWarnings` rejects output when a known limitation or font substitution is detected; it is not a certification of PowerPoint fidelity.

Supported rendering includes layout/master placeholder geometry, theme colors and fonts, template artwork, nested group transforms, rotated/flipped shapes with readable text, text size/bold/italic/color, basic word wrapping, horizontal alignment and vertical anchoring, paragraph spacing, common bullets/numbering, underline/strikethrough, raster pictures and cropping, basic SVG paths, common shape geometry, custom path segments, linear/radial gradients, table cell text/fills/borders, and saved SmartArt drawings. Export is limited to 500 selected slides and 64 megapixels per decoded raster image. Cancellation is supported between slides, shapes, and paragraphs.

**Fidelity limits:** this is an initial native converter, not a complete Office layout engine. Charts, SmartArt without a saved drawing, unsupported preset/custom geometry, shadows/3D, advanced SVG/CSS, vertical/multicolumn text, full complex-script shaping, advanced numbering, and some table/theme effects are incomplete. Unsupported content is reported where detected and may be omitted or approximated. Animations/transitions are static; media playback, speaker notes, accessibility structure and interactive hyperlinks are not exported. Visually review output before distributing it.

Fonts available to Skia on the host are used when possible; unavailable fonts fall back to bundled Noto Sans and produce conversion notes. Embedded PPTX fonts are not loaded. Font availability and metrics can differ across platforms, including Linux containers, so line wrapping may differ. The extension includes Linux native assets and a fallback font; PowerPoint/LibreOffice are not runtime dependencies. Noto Sans license is included under `PptxObjectKit.Extensions.Pdf/Fonts`.

## API

Inspection, Markdown and PDF endpoints accept raw PPTX bytes. The edit endpoint accepts multipart form data with a `file` part containing PPTX bytes and an `edit` field containing JSON.

| Endpoint | Result |
| --- | --- |
| `GET /api/health` | Health status |
| `POST /api/documents/inspect` | Slide dimensions, text objects, notes, warnings |
| `POST /api/documents/markdown?notes=true&images=true` | JSON with `markdown` |
| `POST /api/documents/pdf?slides=1,2&includeHidden=true&strict=false` | JSON with `pdfBase64`, `pageCount`, `warnings`; omit `slides` for all |
| `POST /api/documents/edit` | Updated PPTX bytes |

Malformed/unsupported documents return HTTP 422. Requests are limited to 50 MB, 10,000 package parts, 32 MB per expanded part, 128 MB total expanded size, and 2,000 slides. External resources are not fetched. XML DTDs are prohibited.

## Current limits

This is an initial toolkit, not a PowerPoint replacement or a production-ready rendering engine.

- Supports Transitional OOXML `.pptx`; legacy `.ppt`, encrypted packages, and Strict OOXML are not supported.
- Preview supports basic text, pictures, paths, rectangles/ellipses, gradients, transforms and opacity. Theme fill/line/font references and common elbow connectors with triangle/open arrowheads are supported. It resolves slide ordering, placeholder geometry and text defaults from layouts/masters, theme colors and fonts, template artwork, nested group transforms, image cropping, saved SmartArt drawings, common rounded-corner/brace shapes, readable text inside flipped shapes, word and long-URL wrapping across formatted runs (including non-breaking spaces), hyperlink colors/underlines, paragraph alignment/spacing, basic numbered and bulleted lists, common Wingdings symbol fallbacks, table cells with fills/borders and header/banded-row styling, and vertical text anchoring. Advanced shape geometry, full theme style references, advanced list numbering and symbol fonts, complex table style overrides, charts, SmartArt without a saved drawing, animation, video, and exact Office font layout remain incomplete. Unsupported content can be absent or approximate in preview but is preserved in the saved package.
- Signed packages can be read but not edited.
- Slide deletion removes the slide from the presentation sequence. Unreferenced parts are retained to avoid damaging shared assets/internal links. **Deletion is not redaction or content sanitization.**
- Creating a presentation from scratch is not yet supported.
- No server-side user authentication, persistence or collaboration. Add deployment access controls appropriate to your environment.

## Validation

```powershell
dotnet build PptxObjectKit.slnx -c Release
dotnet run --project PptxObjectKit.Tests -c Release
dotnet run --project PptxObjectKit.Extensions.Pdf.Tests -c Release
cd PptxObjectKit.Playground.Ui
npm run build
npx playwright test
```

Local browser tests use installed Chrome. CI installs Chromium and runs library and browser tests, then builds service images. The fixture is synthetic and contains no user documents.

## Markdown extension

Markdown conversion is supplied by the optional
[PptxObjectKit.Extensions.Markdown](PptxObjectKit.Extensions.Markdown/README.md) package.
Add a reference to it when using `PptxMarkdownConverter`; import the
`PptxObjectKit.Extensions.Markdown` namespace. Conversion behavior is unchanged.
The playground and CLI include the extension. Core-only consumers do not need it.
