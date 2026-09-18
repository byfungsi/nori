# Public API notes

## Model

`parseWorkbookSnapshot(unknown)` validates, copies, and freezes version 1 snapshots. Every sheet has a stable branded ID, unique case-insensitive name, logical extents, and an A1-keyed sparse cell record. Cell values are number/string/boolean/null or a tagged spreadsheet error. Formula text and workbook styles are optional. `createEmptySnapshot()` makes a 100 × 26 empty sheet.

`parseAddress('$b$12')` returns zero-based `{ row: 11, column: 1 }`; `addressOf(position)` encodes already-valid coordinates. Addresses are bounded to XFD1048576. For arbitrary numeric inputs, validate coordinates through the model boundary before calling `addressOf`. `parseSheetId` refines external IDs; most callers obtain IDs directly from parsed snapshots.

## Import

```ts
const result = parseXlsx(bytes, {
  maxInputBytes: 20_000_000,
  maxUncompressedBytes: 100_000_000,
  maxCells: 200_000,
});
```

The same values are defaults. Limits must be positive safe integers. Limits bound ordinary archives using their declared ZIP sizes; this synchronous adapter is not an adversarial archive sandbox. Applications accepting hostile documents should isolate CPU/memory work and enforce their own upload policy.

A successful result is `{ snapshot, warnings }`. Failure is `XlsxError` with `reason: 'invalid-file' | 'unsupported' | 'limit'`. Never discard warnings in a fidelity-sensitive import workflow.

## Runtime

```ts
const created = createWorkbook(snapshot, {
  historyLimit: 100,
  // functions: an optional function registry copied at construction
});
```

- `getState()` is referentially stable until a state transition. Subscribe callbacks are synchronous and called once per successful commit.
- `subscribe(listener)` returns an unsubscribe function. Keep listeners nonthrowing.
- `dispatch(command)` returns a Result. Set cells with explicit scalar input or `{ value: null, formula: 'SUM(A1:A3)' }`. `clearCell` deletes a sparse entry. Set cells outside current extents grow the extents, within Excel bounds.
- `undo()` / `redo()` return false when unavailable.
- `getCellValue(sheetId, address)` returns current calculated scalar/array data. Snapshot cell values are caches, not authoritative calculated output.
- `setActiveSheet(sheetId)` clears selection. `setSelection(selection | null)` sets explicit corners; state exposes normalized `range`, `anchor`, and `focus`. `selectCell(sheetId, position, { extend: true })` extends from the current anchor. `moveSelection(direction, { extend, bounds })` supports platform-neutral arrow navigation and skips hidden/filtered entries. None of these operations enter document history.
- `exportSnapshot()` returns immutable JSON-compatible model data. Pass parsed JSON back to `createWorkbook` for safe rehydration.

`getVisibleRange(layout, viewport, extent)` computes an inclusive rectangle from logical units; no pixel, DOM, canvas, or native measurement objects cross its contract.

## Formula

```ts
import {
  createFunctionRegistry,
  parseFormula,
  evaluateFormula,
} from "@byfungsi/nori/formula";
import { cellError } from "@byfungsi/nori/model";

const functions = createFunctionRegistry();
functions.set("DOUBLE", (args) =>
  typeof args[0] === "number" ? args[0] * 2 : cellError("#VALUE!"),
);
const ast = parseFormula("DOUBLE(21)");
if (ast.ok) {
  const result = evaluateFormula(
    ast.value,
    {
      cell: () => cellError("#REF!"),
      range: () => cellError("#REF!"),
    },
    functions,
  ); // 42
}
```

Function names are uppercased by the parser, so register uppercase names. `IF` and `IFERROR` are reserved evaluator-owned lazy forms. Changing the source registry after workbook construction does not change that workbook. `collectReferences` returns syntactic references, retaining compact ranges and both conditional branches. `DependencyGraph` manages opaque keys and transitive dependent traversal; core owns workbook-specific key construction.

## Pivot

```ts
const result = createPivot(
  [
    { region: "West", revenue: 100 },
    { region: "West", revenue: 50 },
  ],
  {
    rows: ["region"],
    values: [{ id: "total", field: "revenue", aggregate: "sum" }],
  },
);
// result.value.groups[0] = { key: ['West'], values: { total: 150 } }
// result.value.totals = { total: 150 }
```

Rows may contain multiple grouping fields. Groups retain first-seen order; numeric/string keys remain distinct. All configured fields must exist in every record. Numeric measures ignore text, blanks, booleans, and error values. Count counts nonblank, non-error values, including text. Empty numeric groups produce sum/count 0 and average/min/max null. No grouping fields produces one group for nonempty data. Grand totals aggregate original rows (an average of all records, not an average of group averages).

Source records are typed semantic input; parse less-trusted external records before calling this API. Measures must have unique IDs. Results are immutable.

## Layout and merge commands

```ts
workbook.dispatch({ type: "resizeColumn", sheetId, column: 1, width: 240 });
workbook.dispatch({ type: "resizeRow", sheetId, row: 3, height: 64 });
// null removes an override and restores the sheet default.
workbook.dispatch({ type: "resizeColumn", sheetId, column: 1, width: null });
workbook.dispatch({
  type: "mergeCells",
  sheetId,
  range: { start: { row: 0, column: 0 }, end: { row: 1, column: 2 } },
});
workbook.dispatch({ type: "unmergeCells", sheetId, range });
```

Sizes are positive logical units up to 10,000. UI gestures clamp to a practical minimum of 40 for columns and 20 for rows; imported smaller dimensions are retained. `getColumnWidth` and `getRowHeight` resolve overrides and defaults. `getVisibleRange` accepts sparse `columnWidths`/`rowHeights` in addition to uniform base dimensions.

`SheetSnapshot.merges` contains nonoverlapping normalized inclusive rectangles. Merge commands reject single cells, out-of-bounds or partial-overlap ranges, and any content outside the new top-left anchor. Explicit merge commands can combine fully contained existing merges only when the content invariant holds. Unmerge removes every region intersecting the given range. All errors leave the document unchanged.

`SheetSnapshot.frozen` stores leading row/column counts. `hiddenRows` and `hiddenColumns` are sorted, unique, zero-based indexes. `autoFilter` stores a range and semantic value-list or custom comparison criteria; `sort` stores value-sort columns and directions. `getVisibleRows(sheet, { start, end, limit, read })` accepts a scalar reader for current formula values; without it, it uses snapshot caches. Hidden rows explicitly saved in the file remain hidden even if an edited value later matches a filter.
