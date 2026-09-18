import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { createWorkbook } from "@nori-internal/core";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { parseXlsx } from "@nori-internal/xlsx";
import { basicBytes, unwrap } from "./helpers";
function replace(
  path: string,
  transform: (text: string) => string,
): Uint8Array {
  const files = unzipSync(basicBytes);
  const file = files[path];
  if (!file) throw Error("fixture part");
  files[path] = strToU8(transform(strFromU8(file)));
  return zipSync(files);
}
describe("XLSX adapter", () => {
  it("imports sparse cells, rich shared strings, multiple sheets, formula caches and styles", () => {
    const { snapshot } = unwrap(parseXlsx(basicBytes));
    expect(snapshot.sheets).toHaveLength(2);
    expect(snapshot.sheets[0]?.cells["A2"]?.value).toBe("West");
    expect(snapshot.sheets[0]?.cells["F10"]?.value).toBe(true);
    expect(snapshot.sheets[0]?.cells["B6"]).toMatchObject({
      value: 2650,
      formula: "SUM(B2:B4)",
      style: {
        fontWeight: "bold",
        background: "#DDF4D8",
        numberFormat: "0.00",
      },
    });
    expect(Object.keys(snapshot.sheets[0]?.cells ?? {})).toHaveLength(11);
  });
  it("rejects malformed ZIP/XML and missing relationships", () => {
    expect(parseXlsx(new Uint8Array([1, 2, 3])).ok).toBe(false);
    expect(parseXlsx(replace("xl/workbook.xml", () => "<broken>")).ok).toBe(
      false,
    );
    expect(
      parseXlsx(
        replace("xl/_rels/workbook.xml.rels", (s) =>
          s.replace('Id="rId1"', 'Id="missing"'),
        ),
      ).ok,
    ).toBe(false);
  });
  it("resolves workbook relationship targets instead of assuming sheet filenames", () => {
    const files = unzipSync(basicBytes);
    const sheet = files["xl/worksheets/sheet1.xml"];
    if (!sheet) throw Error("fixture");
    files["custom/sales.xml"] = sheet;
    delete files["xl/worksheets/sheet1.xml"];
    const rels = files["xl/_rels/workbook.xml.rels"];
    if (!rels) throw Error("fixture");
    files["xl/_rels/workbook.xml.rels"] = strToU8(
      strFromU8(rels).replace("worksheets/sheet1.xml", "../custom/sales.xml"),
    );
    expect(
      unwrap(parseXlsx(zipSync(files))).snapshot.sheets[0]?.cells["B2"]?.value,
    ).toBe(1200);
  });
  it("imports merged cells while retaining unsupported shared-formula caches", () => {
    const input = replace("xl/worksheets/sheet1.xml", (s) =>
      s
        .replace("<f>SUM(B2:B4)</f>", '<f t="shared" si="0">SUM(B2:B4)</f>')
        .replace(
          "</worksheet>",
          '<mergeCells><mergeCell ref="C1:D1"/></mergeCells></worksheet>',
        ),
    );
    const result = unwrap(parseXlsx(input));
    expect(result.warnings.map((w) => w.feature)).not.toContain("merged cells");
    expect(result.snapshot.sheets[0]?.merges).toEqual([
      { start: { row: 0, column: 2 }, end: { row: 0, column: 3 } },
    ]);
    expect(result.snapshot.sheets[0]?.cells["B6"]?.formula).toBeUndefined();
    expect(result.snapshot.sheets[0]?.cells["B6"]?.value).toBe(2650);
  });
  it("handles 1904 metadata and XML entity text", () => {
    const input = replace("xl/workbook.xml", (s) =>
      s.replace('date1904="0"', 'date1904="1"'),
    );
    expect(unwrap(parseXlsx(input)).snapshot.dateSystem).toBe("1904");
    const text = replace("xl/sharedStrings.xml", (s) =>
      s.replace("Region", "Region &amp; area"),
    );
    expect(unwrap(parseXlsx(text)).snapshot.sheets[0]?.cells["A1"]?.value).toBe(
      "Region & area",
    );
  });
  it.each([
    { maxInputBytes: 10 },
    { maxUncompressedBytes: 10 },
    { maxCells: 1 },
    { maxCells: -1 },
  ])("enforces resource limits %j", (options) => {
    const result = parseXlsx(basicBytes, options);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.reason).toBe("limit");
  });
  it("rejects invalid shared indexes, duplicate cells and DTDs", () => {
    expect(
      parseXlsx(
        replace("xl/worksheets/sheet1.xml", (s) =>
          s.replace("<v>0</v>", "<v>99</v>"),
        ),
      ).ok,
    ).toBe(false);
    expect(
      parseXlsx(
        replace("xl/worksheets/sheet1.xml", (s) =>
          s.replace('r="B2"', 'r="A2"'),
        ),
      ).ok,
    ).toBe(false);
    expect(
      parseXlsx(replace("xl/workbook.xml", (s) => "<!DOCTYPE workbook>" + s))
        .ok,
    ).toBe(false);
  });
});

it("imports an independently generated openpyxl workbook with Unicode, styles and the 1904 epoch", () => {
  const imported = unwrap(
    parseXlsx(new Uint8Array(readFileSync("tests/fixtures/openpyxl.xlsx"))),
  );
  expect(imported.snapshot.dateSystem).toBe("1904");
  expect(imported.snapshot.sheets.map((sheet) => sheet.name)).toEqual([
    "Data & notes",
    "Résumé",
  ]);
  expect(imported.snapshot.sheets[0]?.cells["A1"]?.style).toMatchObject({
    fontWeight: "bold",
    color: "#123456",
    background: "#EEEEAA",
  });
  expect(imported.snapshot.sheets[1]?.cells["B1"]?.value).toBe("雪 & <text>");
  const runtime = unwrap(createWorkbook(imported.snapshot));
  const summary = imported.snapshot.sheets[1];
  if (!summary) throw new Error("fixture");
  expect(runtime.getCellValue(summary.id, "A1")).toBe(60);
});

it("preserves imported geometry, merges, visibility, freeze panes and saved sorting/filtering", () => {
  const imported = unwrap(
    parseXlsx(new Uint8Array(readFileSync("tests/fixtures/layout.xlsx"))),
  );
  const sheet = imported.snapshot.sheets[0];
  if (!sheet) throw new Error("fixture");
  expect(sheet.columnWidths?.["0"]).toBe(168);
  expect(sheet.columnWidths?.["3"]).toBe(336);
  expect(sheet.rowHeights?.["0"]).toBe(40);
  expect(sheet.merges).toEqual([
    { start: { row: 0, column: 0 }, end: { row: 1, column: 3 } },
  ]);
  expect(sheet.cells["A1"]?.style?.wrapText).toBe(true);
  expect(sheet.frozen).toEqual({ rows: 4, columns: 1 });
  expect(sheet.hiddenColumns).toEqual([2]);
  expect(sheet.hiddenRows).toEqual([5]);
  expect(sheet.autoFilter?.columns).toHaveLength(2);
  expect(sheet.sort?.conditions).toEqual([
    { column: 1, direction: "descending" },
  ]);
  expect([
    sheet.cells["B5"]?.value,
    sheet.cells["B6"]?.value,
    sheet.cells["B7"]?.value,
  ]).toEqual([450, 300, 200]);
  expect(
    imported.warnings.some((w) =>
      /merged cells|column dimensions|sheet views/.test(w.feature),
    ),
  ).toBe(false);
});

it("warns on unsupported filters while retaining the workbook saved hidden rows", () => {
  const bytes = replace("xl/worksheets/sheet1.xml", (source) =>
    source
      .replace('<row r="3">', '<row r="3" hidden="1">')
      .replace(
        "</worksheet>",
        '<autoFilter ref="A1:B4"><filterColumn colId="0"><dynamicFilter type="today"/></filterColumn></autoFilter></worksheet>',
      ),
  );
  const imported = unwrap(parseXlsx(bytes));
  expect(imported.snapshot.sheets[0]?.hiddenRows).toEqual([2]);
  expect(imported.warnings.some((w) => w.feature.includes("dynamic"))).toBe(
    true,
  );
});
