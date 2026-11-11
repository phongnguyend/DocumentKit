# DocxObjectKit

A .NET 11 DOCX object library and React 19 playground following the PptxObjectKit
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
  preserve unchanged runs and hyperlinks outside the changed range; inserted text
  inherits formatting at the insertion point. Paragraphs containing fields, notes,
  revisions, drawings or page/column breaks reject plain-text replacement. Add/delete
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
column flow and explicit column breaks, decimal, Roman and alphabetic list markers (including multilevel starts and restart rules), cached content
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

Equations, complex-script shaping, tracked changes, general field recalculation,
mixed page sizes, advanced table styles and drawing content remain incomplete. EMF pictures
using solid brushes, filled line/Bezier paths and path clipping are decoded natively
to transparent PNGs for both browser preview and PDF. The original package image
is preserved. Other unsupported EMF operations, EMF+ Only and patterned brushes produce conversion
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
page references stay cached until **Document > Refresh table of contents** is used.
Refresh supports one TOC content control using outline levels (default 1-3),
preserves a TOCHeading title, adds heading bookmarks, regenerates entries and
PAGEREF caches, and iterates pagination until page numbers stabilize. Section
numbering restarts and decimal/Roman/alphabetic page labels are reflected in TOC
page references. Automatic list numbers on headings are included in regenerated
entry titles. Custom style,
caption, bookmark-range and multiple/unwrapped TOCs are rejected. This updates
the saved DOCX; unsupported fields retain their existing caches. Table formatting
resolves inherited styles, first/last row/column and banded-cell formatting, with
direct cell/run overrides and inherited cell margins. Matching page counts do not
guarantee identical Word layout.

Native image preview and PDF export support source cropping (including negative padding), rotation and horizontal/vertical flips. The source image bytes remain unchanged; downloading an image returns the full original image.

Body text paragraphs honor `keepLines` and consecutive `keepNext` chains when the
kept content fits in one page/column. Oversized paragraphs still split. Measurement
preserves numbering counters and creates no duplicate selection objects. Explicit
page/column breaks, positioned frames and drawings are excluded from this preflight.
Widow/orphan control measures full line advances (including blank lines) and keeps two lines together at ordinary paragraph boundaries
unless disabled. Keep-next also binds captions/headings to the first table row.
Trailing spaces do not create extra lines/pages or shift centered/right-aligned
text; subsequent words still wrap when the intervening spaces exceed the line width.
Terminal manual page breaks do not introduce a blank paragraph line on the next
page unless `splitPgBreakAndParaMark` requests it. Word 2013+ compatibility mode
suppresses paragraph space-before at ordinary page starts, while retaining it at
document/section starts. Legacy compatibility keeps its earlier spacing behavior.
Nested tables flow inside their parent cells (up to 16 levels). Splittable rows
use available page space without cutting ordinary text lines; cantSplit keeps
rows together when they fit. Oversized rows still split. Complex vertical merges
across page boundaries and oversized cell images remain approximate.

Plain-text table cells paginate independently with widow/orphan and keep-lines
constraints. Repeated headers stay intact, and a row starts on the next page when
its declared minimum height cannot fit. Top-aligned text in vertical merges can
occupy the combined row span without inflating the first row, when its text fits
on the starting page; complex or cross-page merged content retains conservative layout.
Inline image paragraphs reserve baseline descender space and declared effect extents
instead of adding a blank text line. Effect extents also participate in image alignment;
this does not imply rendering all drawing effects. Text wrapping allows at most two
twips (0.1 pt) of font-measurement tolerance at the line edge.

Legacy VML image previews in embedded objects (including Visio) are displayed as
read-only images. On Windows, EMFs outside the portable decoder subset use the
native GDI renderer and are rasterized on a white background. No Office/Visio
installation is required. Linux containers also render supported Visio EMF previews
through portable Skia playback: saved state, transforms, solid pens/brushes,
polygons, rectangular clips, DIB images, and positioned text. EMF+ Dual uses its
complete EMF fallback stream; EMF+ Only and unsupported drawing operations still
produce conversion warnings. Fonts may be substituted when unavailable, and glyph
indices require the original font (apart from the known Calibri space glyph).
This is a diagram-preview subset, not complete EMF/EMF+ support. Embedded objects are preserved,
not executed or edited. Legacy floating placement remains approximate.

## Notes, fonts and remaining layout limits

Footnote references and note text are rendered, with bounded re-pagination to
reserve space at the bottom of the reference page. Endnotes appear on a new page
after the document. Notes are read-only in the playground. Note numbering currently
uses continuous decimal labels. Footnotes occupying over half a page are placed
after the document with a warning; continuation footnotes and section-specific
note numbering/placement remain unsupported. Re-pagination may retain conservative
spacing on pages from which a reference moved, also reported as a warning.

When the requested font is unavailable, installed metric-compatible substitutes
are preferred: Arial/Liberation Sans, Times New Roman/Liberation Serif,
Courier New/Liberation Mono, Calibri/Carlito and Cambria/Caladea. Otherwise the
bundled Noto Sans fallback is used. Both native API Docker images install Liberation,
Carlito and Caladea. Substitutions are reported; compatible metrics
do not guarantee identical glyphs, shaping or Microsoft Word pagination.

The shared edit API and CLI accept `{"kind":"refresh-toc"}` for TOC regeneration.

Header/footer PAGE and NUMPAGES fields support simple and complex field encodings,
including instructions split across runs and cached results split into multiple
fragments. PAGE inherits section numbering formats; explicit Arabic, Roman or
alphabetic formatting switches override them. Physical navigation indices remain
unchanged. Minimum (`atLeast`) paragraph line spacing uses point-based minimums,
while automatic spacing uses a line-height multiplier.

Body DrawingML horizontal and vertical line shapes render as vector rules in both native preview and PDF, preserving anchor positions, stroke thickness, direct RGB colors and theme colors. Diagonal/rotated lines and advanced line effects remain limited.

## Markdown extension

Markdown conversion is supplied by the optional
[DocxObjectKit.Extensions.Markdown](DocxObjectKit.Extensions.Markdown/README.md) package.
Add a reference to it when using `DocxMarkdownConverter`; import the
`DocxObjectKit.Extensions.Markdown` namespace. Conversion behavior is unchanged.
The playground and CLI include the extension. Core-only consumers do not need it.
