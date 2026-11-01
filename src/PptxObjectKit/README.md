# PptxObjectKit

A first version of a .NET 10 library and playground for opening, inspecting, editing, and saving `.pptx` presentations. It is independent of PdfObjectKit and LibreOffice. The library uses ZIP and XML APIs from .NET; it has no NuGet dependencies.

## Projects

| Project | Purpose |
| --- | --- |
| `PptxObjectKit` | Package models, reader/writer, slide extraction, editors, Markdown conversion |
| `PptxObjectKit.Application` | Shared edit dispatch for API and CLI |
| `PptxObjectKit.Cli` | Inspect, rewrite, edit and convert to Markdown |
| `PptxObjectKit.Playground.Api` | Stateless ASP.NET Core minimal API |
| `PptxObjectKit.Playground.Ui` | React 19 / TypeScript playground |
| `PptxObjectKit.Tests` | Console regression suite and a small synthetic presentation |

## Playground

The shell follows PdfObjectKit: shared branding and Open action above browser-style tabs, a document toolbar with Slides/Document menus and download/history actions, thumbnails on the left, an inspector on the right, and slide navigation at the bottom.

- Open or drop one or more `.pptx` files, up to 50 MB each, into separate browser-style document tabs. Add, close, or drag tabs to reorder; arrow keys navigate the tab strip. Each tab preserves its slide selection, edits, inspector, and undo/redo history. Closing a tab with edits not yet downloaded requires in-app confirmation.
- Navigate using slide thumbnails with slide numbers and titles. Previews load as the sidebar scrolls and refresh after edits. Slide order follows presentation relationships, including decks whose filenames do not match their order.
- View an experimental canvas preview or use the selectable-text layer to copy slide text.
- Select a text object in the inspector and replace its text. The first run's formatting in each paragraph is retained; mixed formatting within that paragraph is simplified. The UI limits replacement text to 4,000 characters.
- Reorder slides by dragging the slide list or using Move earlier / Move later.
- Delete slides using an in-app confirmation; deleting the last slide is rejected.
- Undo/redo up to ten recent edits and download the current PPTX.
- View Markdown as rendered content or source, include/exclude notes and pictures, reapply conversion options, copy or download the result.

Files are processed in memory by the API and are not saved there. Changes are kept in the browser session until downloaded. Opening more files creates new tabs. Download each edited presentation to keep its changes before leaving the page.

### Run locally

Requires .NET 10 and Node.js 22.12+.

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

## API

Inspection and Markdown endpoints accept raw PPTX bytes. The edit endpoint accepts multipart form data with a `file` part containing PPTX bytes and an `edit` field containing JSON.

| Endpoint | Result |
| --- | --- |
| `GET /api/health` | Health status |
| `POST /api/documents/inspect` | Slide dimensions, text objects, notes, warnings |
| `POST /api/documents/markdown?notes=true&images=true` | JSON with `markdown` |
| `POST /api/documents/edit` | Updated PPTX bytes |

Malformed/unsupported documents return HTTP 422. Requests are limited to 50 MB, 10,000 package parts, 32 MB per expanded part, 128 MB total expanded size, and 2,000 slides. External resources are not fetched. XML DTDs are prohibited.

## Current limits

This is an initial toolkit, not a PowerPoint replacement or a production-ready rendering engine.

- Supports Transitional OOXML `.pptx`; legacy `.ppt`, encrypted packages, and Strict OOXML are not supported.
- Preview supports basic text, pictures, paths, rectangles/ellipses, gradients, transforms and opacity. Theme fill/line/font references and common elbow connectors with triangle/open arrowheads are supported. It resolves slide ordering, placeholder geometry and text defaults from layouts/masters, theme colors and fonts, template artwork, nested group transforms, image cropping, saved SmartArt drawings, common rounded-corner/brace shapes, readable text inside flipped shapes, word and long-URL wrapping across formatted runs (including non-breaking spaces), hyperlink colors/underlines, paragraph alignment/spacing, basic numbered and bulleted lists, common Wingdings symbol fallbacks, table cells with fills/borders and header/banded-row styling, and vertical text anchoring. Advanced shape geometry, full theme style references, advanced list numbering and symbol fonts, complex table style overrides, charts, SmartArt without a saved drawing, animation, video, and exact Office font layout remain incomplete. Unsupported content can be absent or approximate in preview but is preserved in the saved package.
- Signed packages can be read but not edited.
- Slide deletion removes the slide from the presentation sequence. Unreferenced parts are retained to avoid damaging shared assets/internal links. **Deletion is not redaction or content sanitization.**
- No creation of presentations, duplication of slides, image editing, object movement/resizing yet.
- No server-side user authentication, persistence or collaboration. Add deployment access controls appropriate to your environment.

## Validation

```powershell
dotnet build PptxObjectKit.slnx -c Release
dotnet run --project PptxObjectKit.Tests -c Release
cd PptxObjectKit.Playground.Ui
npm run build
npx playwright test
```

Local browser tests use installed Chrome. CI installs Chromium and runs library and browser tests, then builds service images. The fixture is synthetic and contains no user documents.
