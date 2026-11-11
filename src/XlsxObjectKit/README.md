# XlsxObjectKit

Native .NET 11 XLSX inspection and editing, with a React 19 playground following
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
- Worksheet tabs along the bottom and a full-width native virtualized grid, with
  Cell value, Format, Size, Image and Workbook notes actions in the top toolbar, and toolbar actions aligned right. Panels scroll within the viewport.
- Select cells, navigate with arrow keys, double-click or press Enter/F2 to edit,
  press Enter to save, Escape to cancel, and Delete to clear.
- Use the formula bar or Cell value action to edit text, numbers, booleans or formulas.
  Explicit Text type preserves leading zeros and literal strings beginning with `=`.
  Existing cells select their stored value type automatically (formula cells stay
  Formula regardless of their calculated result). Blank cells use Automatic.
  Editing preserves the selected type; choose another type to convert a value.
- Unwrapped text spills into empty neighboring cells, following left, right or
  center alignment. Values, formulas, merged regions and the active cell editor
  stop overflow. Wrapped/rotated text, merged cells and numbers remain clipped
  to their own bounds. Overflow does not change cell selection or saved values.
  Right-aligned text extends left. Imported cell reading order is respected;
  contextual Arabic/Hebrew text uses RTL direction, while explicit alignment
  remains independent of reading order.
- Edits recalculate supported formulas and their dependencies automatically.
  **Recalculate** refreshes cached results in uploaded workbooks. Formula expressions
  remain editable and calculated values are saved in the downloaded XLSX.
- Drag across cells to select a rectangular range in the current grid window. The outline and status show its address and cell count. Copy exports tab-separated values; Delete clears the range in one undoable edit (up to 10,000 cells). Click a cell, use arrow keys, or press Escape to leave the range. Editing and Format action still target the active cell.
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
- Set the selected row's height and column's width in the Size action.
- Drag the right edge of a column header or the bottom edge of a row header to
  resize it. A guide shows the proposed size; release to save as one undoable edit.
  Escape cancels. Focus a resize handle and use arrow keys for one-unit changes.
  Resizing a column does not alter row heights, and vice versa.
- **Format Painter** copies the selected cell's full style to another cell without
  changing values, formulas, cached results, dimensions or merged regions. Click
  once to paint one cell or range, double-click to keep painting, and press Escape to cancel.
  It works across worksheets in the same workbook, copies inherited styles, and
  each application is one undoable edit. Keyboard users can navigate to a destination
  with arrow keys and press Enter. Drag across destination cells to paint a rectangular range (up to 10,000 cells) in one undoable edit. Escape cancels a pending drag. Merged cells retain their existing structure; only their anchor cells are painted.
- Add, rename, delete and reorder worksheets from the Worksheet menu.
- Select a rectangle and choose **Merge cells**. The top-left value and cell styles
  are preserved; merging rejects other populated cells rather than silently dropping
  their data. Select a merged cell or range and choose **Unmerge cells** to separate
  all intersecting merged areas. Both actions support undo/redo and keep formula caches.
  Partial overlaps, Excel table cells and shared/array formula ranges cannot be merged.
- **Add image** opens the same upload, hand-drawing and QR-code picker used by
  the other ObjectKit playgrounds. The selected cell determines its initial position.
  Use the original image dimensions or enter placement and size in pixels.
- Select a picture to replace, download or delete it from the Image action. Drag
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

The grid opens a window of 10,000 rows. Horizontal scrolling reaches the last
column containing data, a merged range or an image, with A?Z as the minimum.
**Go to cell** can temporarily extend that area up to XFD and jump to another row
window, including addresses such as `AA10001`. Only nearby rows and columns are
rendered. Merged ranges show their top-left cell; edit that anchor.

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
`merge-cells` and `unmerge-cells` take `sheetId` and an A1-style `range` such as
`A1:C2`. The core API is `XlsxEditor.MergeCells(document, sheetId, range, unmerge)`.

Image actions are `add-image`, `replace-image`, `transform-image` and `delete-image`.
They take `sheetId`, `imageId` (except insertion), `imageBase64` (insertion/replacement),
and `x`, `y`, `width`, `height` in pixels (insertion/transform). PNG/JPEG inputs are
limited to 5 MB. The core API is `XlsxImageEditor.Edit`; inspection exposes `Images`
on each worksheet. The UI normalizes other browser-decodable raster formats to PNG.

## Worksheet colors

Worksheet colors are available from **Worksheet > Worksheet color**. Choose a
preset or custom color, or **No color** to clear it. The bottom bar uses Excel-style tabs: solid colors for inactive sheets and a
tinted active tab with a colored underline. Colors support undo/redo
and are saved in downloaded workbooks. Existing RGB, theme/tint and indexed tab
colors are displayed. API/CLI: `set-sheet-color` with `sheetId` and `tabColor`
(`#RRGGBB`, or `null` to clear); core API: `XlsxEditor.SetSheetColor`.

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
  array/data-table formulas are not recalculated by this engine yet; their
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

Click a column letter or row number to select the entire column or row.
Drag across header labels in either direction to select multiple columns or rows;
release to keep the selection, or press Escape to cancel. The
selection includes cells beyond the visible grid window. Merged cells do not expand
header selections into adjacent rows or columns. Click a cell or press Escape to return to cell selection.
Delete clears populated cells in the selected rows or columns (up to 10,000 per
operation), with undo support. Header borders still resize independently.

### Insert rows and columns

Select a cell, range, or row/column headers, then use **Insert > Insert rows
above/below** or **Insert > Insert columns left/right**. The selected range's
height or width determines how many rows or columns are inserted (up to 10,000).
Changes save immediately and support undo/redo.

Insertion shifts cell values and styles, row heights, column widths, merged
ranges, worksheet ranges, frozen panes, and cell-anchored images. Absolute image
anchors remain fixed. A1 formula references (including absolute references,
cross-sheet references, and defined names) update, followed by recalculation of
supported formulas. Formula string literals remain unchanged. Insertion that
would push cell content beyond worksheet limits is rejected.

Core API: `XlsxStructureEditor.Insert(document, sheetId, rows: true, position: 2,
count: 1)`; positions are one-based. The shared application/API/CLI edit actions
are `insert-rows` and `insert-columns`, with `position` and optional `count`.

Current insertion limitations: protected/signed workbooks, Excel tables, charts,
pivot tables, external links, shared/array/data-table formulas, structured or 3-D
references, worksheet extensions, legacy drawings/comments/controls, and shared
drawing parts are rejected explicitly. New cells are blank; new row dimensions
use worksheet defaults and column formatting may inherit a spanning column style.

### Delete rows and columns

Select cells or drag across row/column headers, then use **Delete > Delete selected
rows** or **Delete > Delete selected columns**. Entire rows/columns are removed,
remaining cells shift up/left, and the change supports undo/redo. Up to 10,000
rows or columns can be deleted per operation. The keyboard Delete key still
clears contents without removing rows or columns.

Ranges contract, fully deleted ranges disappear, and formulas referencing deleted
cells become `#REF!`. Surviving cell styles, row heights, column widths, filters,
merged ranges, and image anchors are adjusted. Two-cell images whose anchors
collapse completely are removed; absolute images remain fixed. The same complex
workbook restrictions listed for insertion apply to deletion.

Core API: `XlsxStructureEditor.Delete(document, sheetId, rows: true, position: 2,
count: 1)`. Shared application/API/CLI actions: `delete-rows` and `delete-columns`,
with a one-based `position` and optional `count` (default 1).

### Hide and unhide rows or columns

Use **Visibility > Hide selected rows/columns** after selecting cells or headers.
Row-header selections show only row visibility actions; column-header selections
show only column visibility actions. Selection-specific actions require entire rows or columns selected using their
headers. The Visibility menu stays visible and is disabled when no actions apply,
keeping the toolbar stable during selection. Unhide all rows/columns remains available whenever hidden rows/columns
exist, regardless of the current selection.
Hidden rows/columns collapse in the grid while retaining their values, formulas,
formatting, and original sizes. Arrow navigation skips hidden headers. Changes
are saved in the XLSX and support undo/redo.

Use **Unhide selected rows/columns** on a range spanning hidden headers, or use
**Unhide all rows/columns** to restore visibility throughout the worksheet.
Cell-anchored images follow collapsed dimensions; absolute images remain fixed.

Core API: `XlsxVisibilityEditor.SetHidden(document, sheetId, rows: true,
position: 2, count: 3, hidden: true)`. Shared API/CLI actions are `hide-rows`,
`hide-columns`, `unhide-rows`, and `unhide-columns`, with one-based `position`
and `count`. Hide is limited to 10,000 rows/columns per action; unhide supports
the full worksheet. Signed/protected sheets cannot be modified.

### Conditional formatting

Select cells or headers and choose **Conditional formatting**. Add numeric
comparisons (greater/less than, inclusive comparisons, equal/not equal, and
between/not between), set text/fill colors, and enter an A1 range. Changes save
immediately, support undo/redo, and remain native conditional-formatting rules
when downloaded. The dialog lists worksheet rules in priority order and lets you
remove rules; replace a rule by removing it and adding its replacement.

The grid evaluates supported `cellIs` rules with literal bounds or same-sheet
A1-reference arithmetic, including percentage thresholds. Differential font and
solid-fill colors support RGB, theme/tint and indexed colors; imported differential
background-color fills are supported too. Rules use cached numeric formula results and
refresh when values recalculate. Blanks, text, booleans and errors are not treated
as numbers. Earlier rules win for conflicting properties; stop-if-true is honored
for supported rules. Base cell formatting is preserved.

Simple imported expression rules support relative A1 references, text/numeric comparisons,
and `AND`, `OR`, `NOT`, including differential bold/italic. These rules use cached
values of stored cells and refresh after recalculation; blank unstored cells are
not evaluated. Expression rules are not authored in the dialog yet.

Other formula-based conditions, text/date rule types, duplicate-value rules, color scales,
data bars, icon sets and other differential
styles are preserved but not previewed or authored yet. Unsupported imported
rules are labelled in the dialog; mixed supported/unsupported rules can display
differently from Excel. Selection formatting and format painter continue to copy
base styles rather than conditional rules.

Shared API/CLI actions: `add-conditional-rule` (`range`, `operator`, `value`,
optional `secondValue`, `color`, `background`) and `remove-conditional-rule`
(`ruleId` from inspection). Core API: `XlsxConditionalFormatEditor.Edit`.

### Shared formulas and weekday text

Inspection expands shared formula followers into their effective A1 formulas.
Application edits and recalculation normalize shared groups into individual formulas
so chained dependencies recalculate and remain editable. The calculation engine
supports `TEXT(serial, "dddd")` and `TEXT(serial, "ddd")` for English weekday names
in both workbook date systems. Other `TEXT` formats remain unsupported.
