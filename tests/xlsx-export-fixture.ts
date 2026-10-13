import {
  parseWorkbookSnapshot,
  type WorkbookSnapshot,
} from "@nori-internal/model";
import { unwrap } from "./helpers";

/** Canonical workbook exercises cached formulas and the supported layout without recalculation. */
export function xlsxExportSnapshot(
  dateSystem: "1900" | "1904" = "1904",
): WorkbookSnapshot {
  return unwrap(
    parseWorkbookSnapshot({
      version: 1,
      dateSystem,
      sheets: [
        {
          id: "data",
          name: "Data & notes",
          rowCount: 20,
          columnCount: 10,
          cells: {
            A1: {
              value: "雪 & <text> 🦄",
              style: {
                fontWeight: "bold",
                fontStyle: "italic",
                color: "#123456",
                background: "#EEEEAA",
                horizontalAlign: "center",
                verticalAlign: "middle",
                wrapText: true,
                numberFormat: "@",
              },
            },
            B1: { value: true },
            C1: { value: "" },
            D1: { value: null, style: { background: "#FF0000" } },
            A2: { value: 2.5, style: { numberFormat: '"$"#,##0.00' } },
            B2: { value: { kind: "error", code: "#DIV/0!" } },
            C2: { value: 6, formula: "SUM(A2:A3)" },
            D2: {
              value: "cached & value",
              formula: 'IF(B1,"cached & value","")',
            },
            E2: { value: false, formula: "1=2" },
            F2: { value: null, formula: "A4" },
            A3: { value: 3.5 },
            B3: { value: 1, style: { numberFormat: "yyyy-mm-dd" } },
            G4: { value: null },
            C5: { value: "merged" },
            H10: { value: "=SUM(A1:A9)" },
          },
          defaultColumnWidth: 70,
          defaultRowHeight: 24,
          columnWidths: { "0": 84, "2": 105 },
          rowHeights: { "1": 32, "3": 40 },
          hiddenRows: [3],
          hiddenColumns: [4],
          frozen: { rows: 1, columns: 2 },
          merges: [
            { start: { row: 4, column: 2 }, end: { row: 5, column: 3 } },
          ],
          autoFilter: {
            range: { start: { row: 0, column: 0 }, end: { row: 2, column: 5 } },
            columns: [
              {
                column: 0,
                condition: {
                  kind: "values",
                  values: ["雪 & <text> 🦄", "literal _x0041_"],
                  includeBlank: true,
                },
              },
              {
                column: 1,
                condition: {
                  kind: "custom",
                  join: "and",
                  comparisons: [
                    { operator: "greaterThan", value: "1" },
                    { operator: "lessThanOrEqual", value: "3" },
                  ],
                },
              },
            ],
          },
          sort: {
            range: { start: { row: 1, column: 0 }, end: { row: 2, column: 5 } },
            caseSensitive: false,
            conditions: [
              { column: 0, direction: "ascending" },
              { column: 2, direction: "descending" },
            ],
          },
        },
        {
          id: "summary",
          name: "Résumé",
          rowCount: 2,
          columnCount: 2,
          cells: {
            A1: { value: 6, formula: "'Data & notes'!C2" },
            B2: { value: "second sheet" },
          },
        },
      ],
    }),
  );
}
