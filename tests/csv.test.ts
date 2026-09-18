import { it, expect } from "vitest";
import fc from "fast-check";
import { parseCsv } from "@byfungsi/nori/csv";
import { createWorkbook } from "@byfungsi/nori";
import { unwrap } from "./helpers";
const sheet = (text: string) => unwrap(parseCsv(text)).snapshot.sheets[0];
it("handles BOM, quoted commas, doubled quotes, embedded newlines and CRLF", () => {
  const result = sheet('\uFEFFName,Note\r\n"A,B","He said ""yes""\nnext"\r\n');
  expect(result?.cells["A2"]?.value).toBe("A,B");
  expect(result?.cells["B2"]?.value).toBe('He said "yes"\nnext');
  expect(result?.rowCount).toBe(2);
});
it("preserves identifiers, whitespace, formulas and numeric-looking text by default", () => {
  const result = sheet("001,123,TRUE,=SUM(A1:A2),  x  ,9007199254740993");
  expect(Object.values(result?.cells ?? {}).map((c) => c.value)).toEqual([
    "001",
    "123",
    "TRUE",
    "=SUM(A1:A2)",
    "  x  ",
    "9007199254740993",
  ]);
  const book = unwrap(createWorkbook(unwrap(parseCsv("=1+1")).snapshot));
  expect(book.getCellValue(book.getState().activeSheetId, "A1")).toBe("=1+1");
});
it("supports explicit delimiter and conservative inference", () => {
  const result = unwrap(
    parseCsv("001;12;-2.5;TRUE;9007199254740993;1e309;=2+2", {
      delimiter: ";",
      valueMode: "infer",
      sheetName: "Import",
    }),
  ).snapshot.sheets[0];
  expect(Object.values(result?.cells ?? {}).map((c) => c.value)).toEqual([
    "001",
    12,
    -2.5,
    true,
    "9007199254740993",
    "1e309",
    "=2+2",
  ]);
  expect(result?.name).toBe("Import");
  expect(
    unwrap(parseCsv("a\tb", { delimiter: "\t" })).snapshot.sheets[0]
      ?.columnCount,
  ).toBe(2);
});
it("preserves ragged rows and empty field extents with sparse storage", () => {
  expect(sheet("a,,\nb\n\n")?.cells).toEqual({
    A1: { value: "a" },
    A2: { value: "b" },
  });
  expect(sheet("a,,\nb\n\n")?.rowCount).toBe(3);
  expect(sheet("a,,")?.columnCount).toBe(3);
  expect(sheet("")?.cells).toEqual({});
  expect(sheet('""')?.columnCount).toBe(1);
  expect(sheet("a\rb")?.rowCount).toBe(2);
});
it.each(['"unfinished', 'a"b', '"closed"junk', '"a" ,b'])(
  "rejects malformed quoting: %s",
  (text) => {
    const result = parseCsv(text);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.reason).toBe("invalid-csv");
  },
);
it("enforces budgets including empty fields and validates options", () => {
  for (const result of [
    parseCsv(",,,", { maxCells: 3 }),
    parseCsv("long", { maxInputCharacters: 2 }),
    parseCsv(",".repeat(16384)),
  ]) {
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.reason).toBe("limit");
  }
  for (const options of [
    { maxCells: 0 },
    { maxInputCharacters: NaN },
    { sheetName: "" },
  ])
    expect(parseCsv("a", options).ok).toBe(false);
});
it("round-trips arbitrary quoted text records", () => {
  fc.assert(
    fc.property(
      fc.array(fc.array(fc.string(), { minLength: 1, maxLength: 6 }), {
        minLength: 1,
        maxLength: 8,
      }),
      (rows) => {
        const csv = rows
          .map((row) =>
            row.map((v) => '"' + v.replaceAll('"', '""') + '"').join(","),
          )
          .join("\r\n");
        const parsed = unwrap(parseCsv(csv)).snapshot.sheets[0];
        expect(parsed?.rowCount).toBe(rows.length);
        const book = unwrap(createWorkbook(unwrap(parseCsv(csv)).snapshot));
        rows.forEach((row, r) =>
          row.forEach((value, c) => {
            expect(
              book.getCellValue(
                book.getState().activeSheetId,
                String.fromCharCode(65 + c) + (r + 1),
              ),
            ).toBe(value === "" ? null : value);
          }),
        );
      },
    ),
    { numRuns: 100 },
  );
});
