import { expect, it } from "vitest";
import { formatNumber } from "@byfungsi/nori/model";
import {
  formatCellValue,
  Spreadsheet,
  SpreadsheetPreview,
} from "@byfungsi/nori/react";
import { createWorkbook, createEmptySnapshot } from "@byfungsi/nori";
import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";

it.each([
  [1234.5, "$#,##0.00", "$1,234.50"],
  [-1234.5, "$#,##0.00", "-$1,234.50"],
  [-1234.5, '"$"#,##0.00_);[Red]("$"#,##0.00)', "($1,234.50)"],
  [1234, '"Rp "#,##0', "Rp 1,234"],
  [1234.5, '#,##0.00" EUR"', "1,234.50 EUR"],
  [1234.5, "[$€-407]#,##0.00", "€1,234.50"],
  [12.5, "\\$0.00", "$12.50"],
  [0, '$0.00;($0.00);"-"', "-"],
  [0, "$0.00", "$0.00"],
  [1.005, "$0.00", "$1.01"],
  [0.125, "0.00%", "12.50%"],
  [12, '"a;b"0', "a;b12"],
  [12, "0;;", "12"],
  [-12, "0;;", ""],
])("formats %s using %s", (value, format, expected) => {
  expect(formatNumber(value, format)).toBe(expected);
});
it.each([
  "[>100]$0",
  "[$-409]0",
  "* $0",
  "mm-dd-yy",
  "0.0E+00",
  "0,",
  "0.##",
  '"broken',
  "0;0;0;@;0",
])("falls back for unsupported %s", (format) => {
  expect(formatNumber(123, format)).toBeUndefined();
  expect(
    formatCellValue(123, { value: 123, style: { numberFormat: format } }),
  ).toBe("123");
});
it("preserves text values and bounds untrusted formatting", () => {
  expect(
    formatCellValue("123", { value: "123", style: { numberFormat: "$0.00" } }),
  ).toBe("123");
  expect(formatNumber(Infinity, "$0")).toBeUndefined();
  expect(formatNumber(1, "0".repeat(4097))).toBeUndefined();
});
it("renders currency in both views while keeping arithmetic numeric", () => {
  const result = createWorkbook(createEmptySnapshot());
  if (!result.ok) throw result.error;
  const workbook = result.value,
    sheetId = workbook.getState().activeSheetId;
  expect(
    workbook.dispatch({
      type: "setCell",
      sheetId,
      address: "A1",
      cell: { value: 12.5, style: { numberFormat: '"Rp "#,##0.00' } },
    }).ok,
  ).toBe(true);
  expect(
    workbook.dispatch({
      type: "setCell",
      sheetId,
      address: "A2",
      cell: { value: null, formula: "A1*2", style: { numberFormat: "$0.00" } },
    }).ok,
  ).toBe(true);
  expect(workbook.getCellValue(sheetId, "A2")).toBe(25);
  for (const component of [Spreadsheet, SpreadsheetPreview]) {
    const html = renderToStaticMarkup(createElement(component, { workbook }));
    expect(html).toContain("Rp 12.50");
    expect(html).toContain("$25.00");
  }
  expect(workbook.exportSnapshot().sheets[0]?.cells["A1"]?.value).toBe(12.5);
});
