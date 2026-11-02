# DocxObjectKit

A .NET 10 DOCX object library and React 19 playground following the PptxObjectKit
project structure and interaction patterns. Edits preserve unrelated ZIP parts.
The core uses .NET ZIP/XML APIs; native layout/PDF rendering lives in a separate
SkiaSharp extension. No Word, LibreOffice, browser printing, or PDF screenshot
conversion is used for PDF export.

| Project | Responsibility |
| --- | --- |
| DocxObjectKit | Package reader/writer, semantic extraction, paragraph/image editors, Markdown conversion |
| DocxObjectKit.Extensions.Pdf | Shared paginated layout and searchable native PDF output |
| DocxObjectKit.Application | Inspection and edit operations shared by API and CLI |
| DocxObjectKit.Playground.Api | Stateless document inspection, edit, Markdown and PDF endpoints |
| DocxObjectKit.Playground.Ui | Browser tabs, page thumbnails, native canvas/text renderer and inspector |
| DocxObjectKit.Tests | Package preservation and editing tests |
| DocxObjectKit.Extensions.Pdf.Tests | Pagination, layout, OCR application and PDF tests |
| DocxObjectKit.Cli | Inspect, edit, Markdown and PDF commands |

## Run

Use separate terminals from this directory:

```powershell
dotnet run --project DocxObjectKit.Playground.Api
cd DocxObjectKit.Playground.Ui
npm ci
npm run dev
```

Open http://localhost:5182. API: http://localhost:5092. Override with `DOCX_API_URL`
and `DOCX_UI_PORT` in the UI `.env` (see `.env.example`). Alternatively,
`docker compose up --build` serves the UI on http://localhost:8126.

## Playground features

- Open multiple DOCX files in browser-style tabs; reorder tabs, close with unsaved
  change confirmation, and retain each document's current page, selection mode,
  zoom and undo/redo history independently.
- Page thumbnails, page navigation and fit-page/50-200% zoom. The workspace scrolls
  without making the whole application scroll.
- Select text for copying, or select paragraph/image objects for editing. Text edits
  preserve the first run's formatting and simplify mixed formatting. Add/delete
  paragraphs. Moving/resizing a paragraph converts it into a page-positioned Word
  frame; table text and paragraphs crossing pages cannot be dragged independently.
- Add images by upload/drop, draw on an aspect-matched canvas, or generate a QR code.
  Download, replace, delete, move and resize images. Pointer release saves the edit;
  Escape cancels a drag. Moving an inline image makes it a page-relative floating
  image. Replacing one occurrence does not modify other uses of its original media.
- OCR images locally with English/Vietnamese Tesseract assets. Edit/move OCR lines,
  copy/download text or apply visible text frames to DOCX, optionally replacing the
  source image. OCR layout is approximate and should be reviewed.
- View Markdown as sanitized HTML or source, toggle images, reapply conversion,
  copy and download. Supports paragraphs, headings, simple list markers, inline
  bold/italic, tables and embedded PNG/JPEG images.
- Native PDF export for all pages or the current page, conversion notes, strict
  warning rejection, preview and download. Text remains searchable/selectable.
- Undo/redo is session-local; download the DOCX to retain edits.

## Layout fidelity and limitations

This is an initial renderer, **not full Microsoft Word layout parity**. It handles
basic style inheritance, font sizes/colors/bold/italic/underline, paragraph margins,
spacing and alignment, wrapping, automatic pagination, explicit page breaks, basic
tables, inline pictures and page-relative floating pictures/text frames. Browser
and PDF share the same server layout; available client fonts can still differ.
Bundled Noto Sans is the fallback font (license in `Fonts`).

Section margins and next-page breaks, equal/unequal column widths, sequential
column flow and explicit column breaks, decimal/letter list markers, cached content
controls, basic headers/footers, table grid widths/spans, per-edge solid borders (visibility, color and thickness), vertical merges and
page-spanning rows are supported. Floating text boxes use approximate placement
and reserve space when wrapping is enabled. Floating pictures support page, margin,
column and margin-area anchors, offsets, and left/center/right/top/bottom alignment.
Footer frames honor margin-relative right/center alignment; text boxes honor their
insets and vertical alignment. First-page and odd/even header/footer selection and
simple PAGE/NUMPAGES fields follow current pagination. PAGE honors section numbering
restarts and continues across sections without a restart; navigation and NUMPAGES
still use physical page indices/counts. Header/footer images are read-only. Tables in unequal columns use the narrowest
column width; end-of-section column balancing and separator rules are not implemented.

Footnotes/endnotes, equations, complex-script shaping, tracked changes, field
recalculation, mixed page sizes, image cropping/rotation, nested
tables, advanced table styles and drawing content remain incomplete. EMF pictures
using solid brushes, filled line/Bezier paths and path clipping are decoded natively
to transparent PNGs for both browser preview and PDF. The original package image
is preserved. Other EMF operations, EMF+ and patterned brushes produce conversion
notes instead of silently drawing incomplete images.
Detected unsupported features produce notes and package data is retained.
DOCX pages flow automatically: slide-style page
reorder/duplicate/delete controls are deliberately not exposed. Inserting on a page
requires a complete paragraph or picture anchor. This does not implement redaction
or promise removal of hidden package content.

Only transitional `.docx` packages are accepted (up to 50 MB, bounded expanded ZIP
size and part count). Legacy `.doc`, encrypted and macro-enabled documents are not
supported. External relationships are not fetched. The API has no persistent file
storage or authentication; deployment hardening is the host application's concern.

## API / CLI

Raw DOCX POST endpoints: `/api/documents/inspect`, `/markdown?images=false`,
`/pdf?pages=1,2&strict=true`. PDF returns `pdfBase64`, `pageCount`, `warnings`.
`/edit` accepts multipart `file` and JSON `edit`, returning edited DOCX bytes.
Edit kinds: `text`, `add-text`, `delete-text`, `transform-text`, `add-image`,
`replace-image`, `resize-image`, `delete-image`, `apply-image-ocr`.
Object IDs come from the latest inspection and are invalidated by structural edits.

```powershell
dotnet run --project DocxObjectKit.Cli -- inspect input.docx
dotnet run --project DocxObjectKit.Cli -- markdown input.docx output.md
dotnet run --project DocxObjectKit.Cli -- to-pdf input.docx output.pdf
dotnet run --project DocxObjectKit.Cli -- edit input.docx edited.docx request.json
```

CLI output files are created without overwriting existing files.

## Verify

```powershell
dotnet build DocxObjectKit.slnx -c Release
dotnet run --project DocxObjectKit.Tests -c Release
dotnet run --project DocxObjectKit.Extensions.Pdf.Tests -c Release
cd DocxObjectKit.Playground.Ui
npm run build
npm run test:e2e
```

Fixtures are generated synthetic documents; no private user documents are included.

Paragraph layout uses font line metrics, collapsed adjacent spacing and contextual
spacing. TOC entries honor left/right/center tab stops and dot/hyphen leaders; TOC
page references remain cached rather than being regenerated. Table formatting
resolves inherited styles, first/last row/column and banded-cell formatting, with
direct cell/run overrides and inherited cell margins. Matching page counts do not
guarantee identical Word layout.
