# Nori

`@byfungsi/nori` is a TypeScript spreadsheet library with a canonical JSON model, a platform-independent runtime, an XLSX adapter, and composable React views. This is an initial working milestone, not full Excel compatibility. Nothing has been published to npm by this repository setup.

## Public entrypoints

Use the root for `parseXlsx`, `createWorkbook`, and common model types. Focused APIs live at `/model`, `/xlsx`, `/core`, `/formula`, `/pivot`, and `/react`.

React is an optional peer dependency. Install React only when importing `/react`. Headless layers never import React or DOM APIs. Internal workspace packages are bundled into this package.

Parsers and commands return explicit Result values. See the source repository documentation for API examples, compatibility limits and roadmap. This initial version supports a basic XLSX-to-runtime-to-React slice; it does not claim complete Excel compatibility.

`/react` exports `Spreadsheet` for editing and `SpreadsheetPreview` for compact read-only chat attachments. Both support multiple sheets and workbook styles. The editor includes drag/Shift selection, row/column resizing, merge/unmerge and imported frozen panes. Preview sheet switching is local to the preview.

The XLSX adapter preserves merged cells, dimensions, hidden rows/columns, worksheet value/custom filters, and saved value-sort metadata. Advanced and table-level filters/sorts are not yet supported.
