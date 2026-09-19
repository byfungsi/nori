# Currency formatting

Available since **0.2.0**. Nori displays currency in both `Spreadsheet` and `SpreadsheetPreview`. Formatting uses `Cell.style.numberFormat`; the stored value and formula results stay numeric. No React theme setting or `renderCell` override is required.

## Set a currency format

```ts
const sheetId = workbook.getState().activeSheetId;
const result = workbook.dispatch({
  type: "setCell",
  sheetId,
  address: "B2",
  cell: {
    value: 1234.5,
    style: { numberFormat: '"Rp "#,##0.00' },
  },
});
if (!result.ok) throw result.error;
// Both views display: Rp 1,234.50
// workbook.getCellValue(sheetId, "B2") remains 1234.5.
```

`setCell` replaces the cell. When formatting an existing cell, spread its current contents and style so you retain formulas and other formatting:

```ts
const sheet = workbook.getState().snapshot.sheets.find((s) => s.id === sheetId);
const cell = sheet?.cells["B2"] ?? { value: null };
workbook.dispatch({
  type: "setCell",
  sheetId,
  address: "B2",
  cell: { ...cell, style: { ...cell.style, numberFormat: "$#,##0.00" } },
});
```

These commands use normal history and subscriptions. Editing a formatted cell still edits its raw value; enter `1234.5`, not a string such as `$1,234.50`. Formatting does not parse currency text or perform currency conversion.

## Supported examples

| Format                  | Value   | Display      |
| ----------------------- | ------- | ------------ |
| `$#,##0.00`             | 1234.5  | $1,234.50    |
| `$#,##0.00`             | -1234.5 | -$1,234.50   |
| `"Rp "#,##0`            | 1234.5  | Rp 1,235     |
| `"USD "#,##0.00`        | 1234.5  | USD 1,234.50 |
| `£#,##0.00`             | 1234.5  | £1,234.50    |
| `¥#,##0`                | 1234.5  | ¥1,235       |
| `#,##0.00" EUR"`        | 1234.5  | 1,234.50 EUR |
| `$#,##0.00;($#,##0.00)` | -1234.5 | ($1,234.50)  |
| `$0.00;($0.00);"-"`     | 0       | -            |
| `[$€-407]#,##0.00`      | 1234.5  | €1,234.50    |

Currency symbols may be unquoted. Text currency codes/labels must be quoted or escaped. Decimal precision is explicit in the format; Nori does not infer a currency's minor units.

Supported numeric patterns are `0` and `#,##0`, optionally followed by a decimal point and 1–10 zeros, optionally followed by `%`. Existing integer, decimal, and percent display remains supported. Values round for display only; a value such as `1.005` with `$0.00` displays `$1.01` without changing the stored number.

## Sections and Excel import

One section applies to all numbers, with a minus sign automatically added for negatives. Two sections specify positive/zero and negative display. Three specify positive, negative, and zero display. An empty selected section hides its display; quoted literal-only sections can show a zero as a dash. Semicolons inside quotes or escaped with a backslash remain literal.

A fourth section is reserved for text in Excel; Nori leaves text values unchanged. Known color tags such as `[Red]` are accepted but do not set text color. Set `Cell.style.color` separately if desired. Excel underscore padding is accepted but omitted because this formatter returns plain text.

XLSX import preserves custom format strings. The adapter also recognizes common English/US built-in currency formats 5–8 and grouped-number formats 3–4. Built-in currency IDs are locale-dependent in Excel; Nori uses dollar formats for these IDs. An explicit custom currency symbol is more reliable for international workbooks.

## Headless formatting

```ts
import { formatNumber } from "@byfungsi/nori/model";

formatNumber(1234.5, '"Rp "#,##0.00'); // "Rp 1,234.50"
formatNumber(1234.5, "mm-dd-yy"); // undefined: unsupported
```

This helper has no React or DOM dependency. It returns `undefined` for unsupported formats or non-finite numbers. The React `formatCellValue` helper uses it and falls back to the raw numeric string. Both the full grid and preview use that same display path, including calculated cells, labels, and tooltips.

## Current limits

- Output uses comma grouping and a decimal point, independent of the browser's language. An `Rp` label does **not** select Indonesian separators. The locale suffix in `[$€-407]` is preserved by import but does not control output locale.
- Accounting fill/repetition (`*`), conditional sections (`[>100]`), date/time, scientific, fractions, scaling commas, optional decimal placeholders, and full Excel format parity are not supported. Unsupported formats fall back to the raw numeric value.
- Currency does not imply financial decimal arithmetic; calculations still use JavaScript numbers.
- Text values are not coerced into amounts. CSV has no format metadata; apply a number format after import to numeric cells.
- There is no built-in currency-picker toolbar or automatic currency detection.

For Excel's full format syntax, see [Microsoft's custom-number-format guidelines](https://support.microsoft.com/en-us/excel/review-guidelines-for-customizing-a-number-format). Nori implements the subset described here.
