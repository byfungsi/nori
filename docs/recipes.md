# Headless recipes

Examples below assume a validated live `workbook` and its `sheetId`, obtained from `workbook.getState().activeSheetId` as in the [quick start](/getting-started). All changes return Results; inspect failures before reporting success.

## Subscribe and clean up

```ts
const unsubscribe = workbook.subscribe(() => {
  const state = workbook.getState();
  console.log(state.revision, workbook.getCellValue(state.activeSheetId, "A3"));
});
// On host teardown:
unsubscribe();
```

Listeners run synchronously. Keep them nonthrowing. A throwing listener can interrupt later notifications after a command has already committed.

## One history entry for multiple edits

```ts
const result = workbook.dispatch({
  type: "batch",
  commands: [
    { type: "setCell", sheetId, address: "B1", cell: { value: 10 } },
    {
      type: "setCell",
      sheetId,
      address: "B2",
      cell: { value: null, formula: "B1*2" },
    },
  ],
});
if (!result.ok) console.error(result.error.message);
else console.log(workbook.getCellValue(sheetId, "B2")); // 20
workbook.undo(); // Boolean: false if unavailable.
workbook.redo();
```

A failed batch is atomic: no partial data changes, history or notifications. Nested batches are not supported. `historyLimit: 0` disables history when creating the runtime.

## Select and resize without a renderer

```ts
const selected = workbook.selectCell(sheetId, { row: 0, column: 0 });
const extended = workbook.selectCell(
  sheetId,
  { row: 2, column: 1 },
  { extend: true },
);
const resized = workbook.dispatch({
  type: "resizeColumn",
  sheetId,
  column: 1,
  width: 240,
});
for (const result of [selected, extended, resized]) {
  if (!result.ok) console.error(result.error.message);
}
```

Selection is transient and does not enter history. Layout mutations do. Coordinates are zero-based; address strings use Excel's A1 convention. Merge closure and hidden/filter navigation live in core. [API details](/api#layout-and-merge-commands).

## Extend formulas

The [formula API example](/api#formula) shows a `DOUBLE` function. Register uppercase names before constructing the workbook and pass `{functions}` to `createWorkbook`. The registry is copied at construction. Functions receive scalar/array arguments and return `FormulaResult`; `IF` and `IFERROR` are evaluator-owned lazy forms.

Functions should be pure and synchronous. Thrown defects propagate to the host. Return `cellError('#VALUE!')` for a modeled spreadsheet error. Arrays are supported as values but do not spill into neighboring cells.

## Build a pivot

```ts
import { createPivot } from "@byfungsi/nori/pivot";
const result = createPivot(
  [
    { region: "West", amount: 100 },
    { region: "West", amount: 50 },
    { region: "East", amount: 20 },
  ],
  {
    rows: ["region"],
    values: [{ id: "total", field: "amount", aggregate: "sum" }],
  },
);
if (!result.ok) console.error(result.error.message);
else console.log(result.value.groups, result.value.totals); // grand total 170
```

Pivots consume records, not cell addresses. Extract calculated cell values in your host adapter when using workbook data. Groups preserve first-seen order. Measures support sum/count/average/min/max. This does not import Excel pivot caches or render a pivot widget.
