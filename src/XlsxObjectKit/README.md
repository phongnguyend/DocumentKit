# XlsxObjectKit

Native .NET 10 XLSX inspection and editing, with a React 19 playground following
the layout and immutable document-editing approach of the other ObjectKits.
This is an initial spreadsheet editor with a from-scratch .NET calculation engine.
It supports a documented formula subset, rather than full Excel compatibility.

## Projects

| Project | Purpose |
| --- | --- |
| `XlsxObjectKit` | ZIP/XML reader and writer, workbook extraction, cell and worksheet editing; no runtime package dependencies |
| `XlsxObjectKit.Application` | Shared inspection and edit operations for the API and CLI |
| `XlsxObjectKit.Playground.Api` | ASP.NET Core minimal API; documents are processed in memory |
| `XlsxObjectKit.Playground.Ui` | React playground with Lucide icons and the shared ObjectKit UI patterns |
| `XlsxObjectKit.Cli` | Create, inspect and edit workbooks |
| `XlsxObjectKit.Tests` | Synthetic fixtures, round-trip checks and independent Open XML schema validation |

## Playground

- Open multiple XLSX files by upload or drag/drop, or create a workbook.
- Click **+** in the tab strip to open a blank upload tab. Choosing or dropping a
  workbook fills that tab; additional files open in separate tabs.
- A blank upload tab is present on startup and restored when the last tab is closed.
- Browser-style document tabs preserve each workbook's worksheet, selection,
  scroll position and undo/redo history. Unsaved tabs use an in-app close dialog.
- Worksheets on the left, a native virtualized grid in the center, Inspector on
  the right, and toolbar actions aligned right. Panels scroll within the viewport.
- Select cells, navigate with arrow keys, double-click or press Enter/F2 to edit,
  press Enter to save, Escape to cancel, and Delete to clear.
- Use the formula bar or Inspector to edit text, numbers, booleans or formulas.
  Explicit Text type preserves leading zeros and literal strings beginning with `=`.
  Existing cells select their stored value type automatically (formula cells stay
  Formula regardless of their calculated result). Blank cells use Automatic.
  Editing preserves the selected type; choose another type to convert a value.
- Edits recalculate supported formulas and their dependencies automatically.
  **Recalculate** refreshes cached results in uploaded workbooks. Formula expressions
  remain editable and calculated values are saved in the downloaded XLSX.
- Drag across cells to select a rectangular range in the current grid window. The outline and status show its address and cell count. Copy exports tab-separated values; Delete clears the range in one undoable edit (up to 10,000 cells). Click a cell, use arrow keys, or press Escape to leave the range. Editing and Inspector formatting still target the active cell.
- Paste tab-separated rows/columns starting at the active cell (up to 10,000 cells).
- Apply bold, italic, text/background colors and alignment. Existing font size,
  number format and border definitions are preserved when applying these controls.
- Render font family/size, underline and strikethrough, solid fills, theme/indexed
  colors and tints, cell borders, horizontal/vertical alignment, indentation,
  wrapping and basic text rotation. Row/column styles and base cell styles are
  inherited, with explicit cell styles taking precedence.
- Display built-in and custom number formats (including dates, times, percentages,
  currency and scientific notation) through [SheetJS SSF](https://github.com/SheetJS/ssf).
  Both the 1900 and 1904 date systems are respected. The formula bar and cell editor
  retain the original values; display formatting does not change workbook data.
- Set the selected row's height and column's width in the Inspector.
- **Format Painter** copies the selected cell's full style to another cell without
  changing values, formulas, cached results, dimensions or merged regions. Click
  once to paint one cell or range, double-click to keep painting, and press Escape to cancel.
  It works across worksheets in the same workbook, copies inherited styles, and
  each application is one undoable edit. Keyboard users can navigate to a destination
  with arrow keys and press Enter. Drag across destination cells to paint a rectangular range (up to 10,000 cells) in one undoable edit. Escape cancels a pending drag. Merged cells retain their existing structure; only their anchor cells are painted.
- Add, rename, delete and reorder worksheets from the Worksheet menu.
- **Add image** opens the same upload, hand-drawing and QR-code picker used by
  the other ObjectKit playgrounds. The selected cell determines its initial position.
  Use the original image dimensions or enter placement and size in pixels.
- Select a picture to replace, download or delete it from the Inspector. Drag
  the picture to move it, or use its eight handles to resize; release to save.
  Escape cancels a drag, Delete removes the selected image, and undo/redo applies
  to image changes. Replacement retains placement and size. Cell data and formula
  caches are not changed by image edits.
- Existing embedded PNG/JPEG pictures with one-cell, two-cell or absolute anchors
  are displayed. New or moved pictures use absolute anchors (fixed pixel positions,
  independent of later row/column resizing). Grouped, linked, vector and in-cell
  pictures are not rendered. Cropped/rotated previews are approximate and cannot
  be moved/resized in the playground. Deletion removes the drawing placement;
  it is not secure removal of old media bytes from the package.
- Download the edited XLSX. Undo/redo retains up to 20 snapshots, with a 100 MB
  history budget excluding the current snapshot (one prior snapshot is retained).

The grid opens a window of 10,000 rows and 26 columns. Use **Go to cell** to jump
to another window, including distant addresses such as `AA10001`. Only visible
rows are rendered. Merged ranges show their top-left cell; edit that anchor.

## Run locally

From the repository root:

```powershell
dotnet run --project src/XlsxObjectKit/XlsxObjectKit.Playground.Api
```

In another terminal:

```powershell
cd src/XlsxObjectKit/XlsxObjectKit.Playground.Ui
npm ci
npm run dev
```

Open <http://localhost:5184>. Vite proxies `/api` to <http://localhost:5094>.
Copy `.env.example` to `.env` to override either port.

## CLI and API

```powershell
dotnet run --project src/XlsxObjectKit/XlsxObjectKit.Cli -- new workbook.xlsx
dotnet run --project src/XlsxObjectKit/XlsxObjectKit.Cli -- inspect workbook.xlsx
dotnet run --project src/XlsxObjectKit/XlsxObjectKit.Cli -- edit workbook.xlsx edit.json edited.xlsx
dotnet run --project src/XlsxObjectKit/XlsxObjectKit.Cli -- calculate workbook.xlsx calculated.xlsx
```

Example `edit.json`:

```json
{
  "action": "cells",
  "sheetId": "1",
  "cells": [
    { "address": "A1", "value": "Revenue", "type": "text" },
    { "address": "B1", "value": "1200.50", "type": "number" },
    { "address": "B2", "value": "=B1*2", "type": "formula" }
  ]
}
```

API endpoints:

- `GET /api/health`
- `POST /api/documents/new`: return a new XLSX.
- `POST /api/documents/inspect`: raw XLSX request body; return workbook JSON.
- `POST /api/documents/edit`: multipart `file` and `edit` JSON; return edited XLSX.

Edit actions are `cells`, `dimensions`, `add-sheet`, `rename-sheet`,
`delete-sheet`, `move-sheet`, `recalculate` and `apply-format`. A `format` on a cell request changes only styling.
Worksheet positions are zero-based. Sheet IDs are stable workbook IDs, not positions.
`apply-format` takes `sourceSheetId`, `sourceAddress`, the destination `sheetId`, and
an `addresses` array (up to 10,000 cells). The core API is `XlsxEditor.ApplyFormat`.

Image actions are `add-image`, `replace-image`, `transform-image` and `delete-image`.
They take `sheetId`, `imageId` (except insertion), `imageBase64` (insertion/replacement),
and `x`, `y`, `width`, `height` in pixels (insertion/transform). PNG/JPEG inputs are
limited to 5 MB. The core API is `XlsxImageEditor.Edit`; inspection exposes `Images`
on each worksheet. The UI normalizes other browser-decodable raster formats to PNG.

## Native formula calculation

`XlsxObjectKit.Calculation.XlsxCalculator.Calculate(document)` returns a new document,
counts of calculated/unresolved formulas, and diagnostics. It does not use ClosedXML,
Excel, LibreOffice or a third-party calculation library. Parsing, evaluation,
dependency traversal and cycle detection are implemented in the core library.

Supported syntax includes numeric/text/boolean constants, parentheses, arithmetic
(`+ - * / ^ %`), comparisons, concatenation (`&`), A1 references with `$` anchors,
rectangular ranges, and cross-sheet references such as `'Other sheet'!A1`.
Use English function names and comma-separated arguments.

Supported functions:

- Aggregates: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`.
- Logical: `IF`, `IFERROR`, `AND`, `OR`, `NOT`, `TRUE`, `FALSE`.
- Math: `ABS`, `INT`, `SQRT`, `POWER`, `MOD`, `ROUND`.
- Text: `LEN`, `LEFT`, `RIGHT`, `MID`, `TRIM`, `LOWER`, `UPPER`, `CONCAT`, `CONCATENATE`.

Normal calculation errors are saved as Excel error values, such as `#DIV/0!`,
`#VALUE!`, `#NUM!` and `#REF!`. Unsupported formulas, invalid syntax, circular
references and resource limits produce diagnostics and clear the affected formula's
cached result. Dependent formulas remain unresolved; `IFERROR` does not hide an
unsupported calculation behind a fabricated fallback result. `IF` evaluates only
its selected branch. Workbook notes display calculation diagnostics after edits.

Calculation limits are 10,000 formulas per workbook, 100,000 cells per range,
200,000 evaluation/reference operations and bounded expression/dependency depth.
Diagnostics are capped at 100 per calculation and the API displays the first 10.
Formulas use double-precision arithmetic; precision, coercion and locale edge cases
can differ from Excel. No external workbooks, network data or code are loaded.

## Docker Compose

```powershell
cd src/XlsxObjectKit
docker compose up --build
```

Open <http://localhost:8128>. The API is exposed at port 5094.
The optional `cli` profile mounts `./files` at `/files`:

```powershell
docker compose --profile cli run --build --rm cli inspect workbook.xlsx
```

## Validation

```powershell
dotnet build src/XlsxObjectKit/XlsxObjectKit.slnx -c Release
dotnet run --project src/XlsxObjectKit/XlsxObjectKit.Tests -c Release --no-build
cd src/XlsxObjectKit/XlsxObjectKit.Playground.Ui
npm run build
npx playwright test
```

Local browser tests use Chrome. CI installs Playwright Chromium and sets `CI`.
The dedicated `xlsxobjectkit.yml` workflow builds/tests the solution and UI,
runs browser tests, and builds the Compose images.

## Current boundaries

- Transitional `.xlsx` only: no legacy XLS, encrypted packages, macro-enabled
  workbooks, Strict OOXML, chart sheets or digital-signature editing.
- Formula support is limited to the subset above. Named ranges, structured table
  references, whole-row/column ranges, lookups, date/time functions, external links,
  dynamic arrays and iterative calculation are not implemented. Workbooks containing
  shared/array/data-table formulas are not recalculated by this engine yet; their
  formula caches are cleared with a diagnostic. ISO date cells are unsupported as
  calculation inputs (numeric Excel date serials can participate in arithmetic).
- Rename/delete is rejected when formulas or named ranges exist, since their
  references cannot yet be rewritten safely. Reordering rejects detected 3D formulas.
- Charts, tables, conditional formatting, filters and other package parts
  are preserved but not rendered or edited. Preserving parts does not guarantee
  that advanced features remain semantically consistent after changing their inputs.
- Fonts use installed browser/system fonts, with fallback when unavailable. Exact
  Excel font metrics, auto-fit row heights, shrink-to-fit, rich-text runs, pattern/
  gradient fills, diagonal borders, conditional/number-format colors, hidden rows/
  columns and frozen panes are not implemented. Unsupported number formats fall
  back to stored values; locale-specific formats can differ from Excel. Wrapped
  and rotated text uses the stored row/column dimensions and can be clipped.
  Merged ranges crossing the current virtual window are not fully represented.
- Row/column insertion/deletion, image OCR,
  Markdown and native PDF export are future extensions. They are not provided by
  this first implementation.
- Limits: 50 MB uploaded ZIP, 128 MB expanded, 32 MB per part, 10,000 parts,
  200,000 stored cells, 256 worksheets and 12 open tabs. Protected workbooks/sheets
  reject edits. Deleted sheets' orphan package parts are retained; deletion is not
  secure data removal.
- The playground has no authentication or persistent storage. Deploy behind your
  own access controls and resource limits if exposing it beyond local development.
