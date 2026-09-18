# Importing XLSX

The XLSX adapter accepts `Uint8Array` bytes and returns a canonical workbook snapshot. It is synchronous and has no file-picker or filesystem dependency.

```ts
import { parseXlsx, createWorkbook } from "@byfungsi/nori";

// Browser boundary: `file` is a File selected by the host.
const imported = parseXlsx(new Uint8Array(await file.arrayBuffer()));
if (!imported.ok) {
  console.error(imported.error.reason, imported.error.message);
} else {
  // Retain these notes in your UI; successful import is not lossless fidelity.
  console.info(imported.value.warnings);
  const created = createWorkbook(imported.value.snapshot);
  if (!created.ok) console.error(created.error.message);
  else showWorkbook(created.value); // Your application's callback.
}
```

In Node, pass `new Uint8Array(await readFile(path))` from `node:fs/promises`. A native application supplies bytes through its platform file adapter; Nori has no React Native binding yet.

## Limits and errors

`parseXlsx(bytes, options)` accepts `maxInputBytes` (20,000,000 default), `maxUncompressedBytes` (100,000,000), and `maxCells` (200,000). Values must be positive safe integers. `XlsxError.reason` is `invalid-file`, `unsupported`, or `limit`.

Declared ZIP sizes limit ordinary input, not hostile resource consumption. For public uploads, the host owns upload policy and CPU/memory isolation. Worker scheduling belongs to a platform adapter; core does not assume Web Workers.

## Layout and saved views

Multiple sheets, dimensions, merged regions, frozen panes, hidden rows/columns, worksheet value/custom filters and saved value-sort metadata are supported. Saved sorted rows remain in their existing order; edits do not automatically re-sort. Advanced and table-level filters/sorts emit warnings. See the precise [support matrix](/support).

The parser never exposes XML nodes, ZIP entries or OOXML names as canonical model fields. If you need to construct a workbook without XLSX, use `parseWorkbookSnapshot` or `createEmptySnapshot`.

## Output and persistence

Use `JSON.stringify(workbook.exportSnapshot())` to save canonical document data, and `createWorkbook(JSON.parse(json))` to rehydrate with validation. There is no `writeXlsx` API yet. Formula caches in snapshots may differ from current calculated results; use `getCellValue` for display or downstream processing.
