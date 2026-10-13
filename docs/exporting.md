# Exporting XLSX

`writeXlsx(snapshot, options)` returns `Result<Uint8Array, XlsxWriteError>`. It serializes the canonical workbook model without filesystem, DOM, native dependencies, or network access. The host owns saving and sharing the bytes.

```ts
import { writeXlsx } from "@byfungsi/nori/xlsx";

const result = writeXlsx(workbook.exportSnapshot(), {
  maxCells: 100_000,
  maxOutputBytes: 10_000_000,
});
if (!result.ok) {
  console.error(result.error.reason, result.error.message);
} else {
  saveWorkbookBytes(result.value); // Your host's callback.
}
```

## Supported output

| Data          | Export behavior                                                                                      |
| ------------- | ---------------------------------------------------------------------------------------------------- |
| Workbook      | Multiple sheets, names/order, logical dimensions, and 1900/1904 date metadata                        |
| Values        | Text, finite numbers, booleans, blanks, and Excel-compatible error codes                             |
| Formulas      | Ordinary formula text and supplied numeric/string/boolean/error/blank caches                         |
| Styles        | Bold/italic, RGB text/fill, alignment, wrapping, and nonempty number-format codes                    |
| Layout        | Explicit/default widths/heights, sparse blank cells, merges, frozen panes, and hidden rows/columns   |
| Saved views   | Supported worksheet value/custom filters and value-sort metadata; existing row order stays unchanged |
| Text encoding | XML characters, Unicode, escaped controls, and literal SpreadsheetML escape sequences                |

Sheet IDs are runtime identities and regenerate on XLSX import. Pixel-based column widths use the same seven-unit digit-width approximation as import; fractional widths are quantized on re-import. Numeric custom-filter criteria use Excel's text representation and re-import as strings. Explicit normal font flags may re-import as omitted defaults, and RGB colors may normalize letter case.

Import warnings still matter: a snapshot cannot reconstruct features discarded during import. Keep original bytes and retain diagnostics when exporting edits. Charts, drawings, tables, macros, named ranges, external-link parts, rich text runs, unsupported styles, and other data outside the canonical model are not recreated. Export produces a new document and does not promise lossless preservation of arbitrary XLSX files.

## Formula caches

Serialization never evaluates a formula. `exportSnapshot()` contains imported or command-provided caches, which may be stale after edits. Refresh supported caches explicitly when calculated results are required:

```ts
import { createWorkbook } from "@byfungsi/nori/core";
import { writeXlsx } from "@byfungsi/nori/xlsx";

const created = createWorkbook(snapshot, { calculation: "manual" });
if (created.ok) {
  const book = created.value;
  // Reads and edits preserve supplied caches until this explicit operation.
  const calculated = book.recalculate({ maxCellVisits: 10_000 });
  if (calculated.ok) {
    const output = writeXlsx(book.exportSnapshot());
    // Save/share only a successful output through your host's authority.
  }
}
```

Automatic calculation remains the default for existing consumers. `getCellValue` in automatic mode calculates on demand; it does not refresh exported caches. Manual mode uses saved values for cell reads and filtered-row navigation, including after edits, undo, and redo. The host must label those caches appropriately and admit calculation explicitly.

`recalculate` evaluates current formulas, refreshes scalar caches, and commits one undoable snapshot update with one notification. Formula errors remain cell values. Array-valued results cannot fit scalar caches and return `WorkbookCalculationError` with reason `unsupported`. Invalid options, an exceeded budget, or command rejection return `invalid-options`, `limit`, or `command`; failure preserves the snapshot and calculation cache.

The default `maxCellVisits` is 100,000. It counts source-cell indexing, dependency edges, and recursive cell reads, including memoized reads. This is a work bound, not a wall-clock limit or isolation for custom function callbacks. Hosts must still bound incoming snapshots, formula size, available functions, and total execution resources.

By default, the written workbook requests manual calculation and preserves its supplied caches. Set `recalculateOnOpen: true` to explicitly request Excel's automatic recalculation on open. This flag does not calculate anything in Nori, and other spreadsheet applications may handle calculation settings differently. Plain cell strings beginning with `=` remain strings.

Nori's `#CYCLE!` and `#ERROR!` are not portable Excel error codes. Export rejects them with `unsupported`; the host must explicitly resolve or replace such values. Unsupported formula grammar can remain stored as ordinary formula text with a valid supplied cache, without Nori evaluating it.

## Limits and errors

| Option                 |                                                                Default |
| ---------------------- | ---------------------------------------------------------------------: |
| `maxCells`             |                                                 200,000 retained cells |
| `maxSheets`            |                                                                    256 |
| `maxStyles`            |                             4,096 unique styles, including the default |
| `maxLayoutEntries`     | 200,000 dimensions, merges, filter entries/values, and sort conditions |
| `maxUncompressedBytes` |                                       100,000,000 serialized XML bytes |
| `maxOutputBytes`       |                                                   20,000,000 ZIP bytes |

Limits must be positive safe integers. The writer also caps styles at 65,000, cell text at 32,767 UTF-16 code units, number-format strings at 4,096 code units, and column widths at 1,785 logical units. Unsupported sheet names, lone UTF-16 surrogates, invalid formula XML, and full-grid freeze panes return explicit failures. XML fragments are counted before retaining complete parts; compressed size is checked after ZIP encoding. ZIP timestamps and sparse-cell ordering are deterministic.

`XlsxWriteError.reason` is `invalid-snapshot`, `invalid-options`, `unsupported`, `limit`, or `archive`. Limits bound document output and model cardinality; they are not a hostile-input CPU/memory isolation mechanism. The host must bound incoming snapshots and choose limits appropriate to its execution environment.

## Verification

Focused tests cover canonical re-import, Unicode/control/literal-escape property cases, formula caches, styles, layout, saved views, malformed input, limits, and deterministic output. An independent Openpyxl reader verifies values, formula caches, dates, relationships, styles, and saved layout. To run that check, use a Python environment with Openpyxl 3.1.5:

```sh
NORI_XLSX_PYTHON=/path/to/venv/bin/python npx vitest run tests/xlsx-export.test.ts tests/xlsx-export-interoperability.test.ts tests/xlsx.test.ts
```

The independent-reader test explicitly skips when `NORI_XLSX_PYTHON` is absent. Its ordinary-text interoperability evidence does not replace the SpreadsheetML escaped-string contract for extended control characters. See [Microsoft's escaped-string specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bd0aa042-434a-4ca7-b25f-4e1fd25a954d).
