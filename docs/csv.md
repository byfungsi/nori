# Importing CSV

CSV import is available from the root package or `@byfungsi/nori/csv`. It produces one canonical worksheet and works without React, DOM, Node APIs, or an XLSX parser dependency.

```ts
import { parseCsv } from "@byfungsi/nori/csv";
import { createWorkbook } from "@byfungsi/nori/core";

const imported = parseCsv("Region,Revenue\r\nWest,1200\r\nEast,850", {
  sheetName: "Revenue",
  valueMode: "infer",
});
if (!imported.ok) {
  console.error(
    imported.error.reason,
    imported.error.offset,
    imported.error.message,
  );
} else {
  const created = createWorkbook(imported.value.snapshot);
  if (!created.ok) console.error(created.error.message);
  else
    console.log(
      created.value.getCellValue(created.value.getState().activeSheetId, "B2"),
    ); // 1200
}
```

## Input and options

`parseCsv(text, options?)` accepts decoded text and returns `Result<CsvImport, CsvError>`. A successful `CsvImport` contains `snapshot`. It does not contain XLSX-specific warnings. In a browser, decode a file with `await file.text()`; for Node, use `readFile(path, 'utf8')`. Other encodings must be decoded by the host first.

| Option               | Default      | Behavior                                                                                                                |
| -------------------- | ------------ | ----------------------------------------------------------------------------------------------------------------------- |
| `delimiter`          | `','`        | Explicit comma, semicolon (`';'`), tab (`'\t'`) or pipe (`'                                                             | '`); no auto-detection |
| `sheetName`          | `'Sheet1'`   | Must satisfy canonical sheet-name validation                                                                            |
| `valueMode`          | `'text'`     | `'text'` preserves nonempty field contents; `'infer'` converts canonical finite numbers and case-insensitive TRUE/FALSE |
| `maxInputCharacters` | `20_000_000` | Positive safe integer; UTF-16 code-unit input budget                                                                    |
| `maxCells`           | `200_000`    | Positive safe integer; counts all fields, including blanks                                                              |

The first record stays in row 1; it is not consumed as metadata. Ragged records are accepted, with the widest record setting the sheet's width. Empty fields are sparse blanks, and empty input becomes an empty 1×1 worksheet. Blank records are retained; one final line ending does not append an extra record.

## Quoting and values

Quoted delimiters, doubled quotes, embedded newlines, CRLF/LF/CR record endings, and one leading Unicode BOM are supported. Whitespace inside a field is preserved. A quote in an unquoted field, an unclosed quote, or any characters between a closing quote and its delimiter/newline are rejected. Parsing is strict: even whitespace after a closing quote is invalid.

Text mode keeps `00123`, `TRUE`, `123` and `=1+1` as text. Infer mode preserves leading-zero identifiers, surrounding whitespace, unsafe integer values and nonfinite numeric strings as text. It does not infer dates, currencies, percentages, or locale-specific numbers. Decimal conversion follows JavaScript number precision.

CSV never creates formulas, even in infer mode. `=SUM(A1:A2)` remains a literal string. To intentionally create a formula, issue a `setCell` command after import. CSV cannot represent styles, merges, multiple sheets, hidden dimensions, or frozen panes.

`CsvError.reason` is `invalid-csv`, `invalid-options`, or `limit`. `offset` is a zero-based UTF-16 location (zero for option/input-budget failures). Failed imports return no partial snapshot. Excel worksheet dimension limits also apply.

## Try it live

The [playground](/playground) accepts `.csv` and `.xlsx` files and includes a downloadable CSV sample. It uses infer mode so ordinary amounts become numbers while identifiers and formula-like strings remain text. Uploaded files stay in the browser. CSV export and streaming parsing are not implemented.
