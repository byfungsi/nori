import { describe, expect, it } from "vitest";
import fc from "fast-check";
import { createWorkbook } from "@byfungsi/nori/core";
import { createWorkbook as createFromRoot } from "@byfungsi/nori";
import { createFunctionRegistry } from "@byfungsi/nori/formula";
import { parseWorkbookSnapshot, parseSheetId } from "@byfungsi/nori/model";
import { parseXlsx, writeXlsx } from "@byfungsi/nori/xlsx";
import { unwrap, workbook } from "./helpers";

const sheetId = unwrap(parseSheetId("budget"));
const snapshot = (a = 2, b = 3, formula = "=SUM(A1:A2)") =>
  unwrap(
    parseWorkbookSnapshot({
      version: 1,
      dateSystem: "1900",
      sheets: [
        {
          id: sheetId,
          name: "Budget",
          rowCount: 3,
          columnCount: 2,
          cells: {
            A1: { value: a },
            A2: { value: b },
            B1: { value: 999, formula, style: { fontWeight: "bold" } },
          },
        },
      ],
    }),
  );

describe("explicit workbook calculation", () => {
  it("keeps manual reads, edits, selection, undo and redo inert until explicitly recalculated", () => {
    const functions = createFunctionRegistry();
    let calls = 0;
    functions.set("PROBE", () => {
      calls += 1;
      return 42;
    });
    const book = unwrap(
      createWorkbook(snapshot(2, 3, "=PROBE()"), {
        calculation: "manual",
        functions,
      }),
    );
    expect(book.getCellValue(sheetId, "$B$1")).toBe(999);
    unwrap(
      book.dispatch({
        type: "setCell",
        sheetId,
        address: "A1",
        cell: { value: 100 },
      }),
    );
    unwrap(book.selectCell(sheetId, { row: 0, column: 0 }));
    unwrap(book.moveSelection("down"));
    expect(book.undo()).toBe(true);
    expect(book.redo()).toBe(true);
    expect(book.getCellValue(sheetId, "B1")).toBe(999);
    expect(calls).toBe(0);
    unwrap(book.recalculate());
    expect(calls).toBe(1);
    expect(book.getCellValue(sheetId, "B1")).toBe(42);
    expect(book.exportSnapshot().sheets[0]?.cells.B1).toEqual({
      value: 42,
      formula: "=PROBE()",
      style: { fontWeight: "bold" },
    });
    expect(book.undo()).toBe(true);
    expect(book.getCellValue(sheetId, "B1")).toBe(999);
    expect(book.redo()).toBe(true);
    expect(book.getCellValue(sheetId, "B1")).toBe(42);
    expect(calls).toBe(1);
  });

  it("retains existing automatic read behavior and refreshes exported caches only explicitly", () => {
    const book = unwrap(createFromRoot(snapshot()));
    expect(book.getCellValue(sheetId, "B1")).toBe(5);
    expect(book.exportSnapshot().sheets[0]?.cells.B1?.value).toBe(999);
    let notifications = 0;
    book.subscribe(() => {
      notifications += 1;
    });
    unwrap(book.recalculate());
    expect(book.exportSnapshot().sheets[0]?.cells.B1?.value).toBe(5);
    expect(notifications).toBe(1);
    expect(book.getState().revision).toBe(1);
  });

  it("refreshes edited cross-sheet dependencies in the exported snapshot", () => {
    const book = unwrap(
      createWorkbook(workbook().exportSnapshot(), { calculation: "manual" }),
    );
    const [sales, summary] = book.exportSnapshot().sheets;
    if (sales === undefined || summary === undefined)
      throw new Error("Expected two fixture sheets");
    const importedCache = book.getCellValue(summary.id, "B1");
    unwrap(
      book.dispatch({
        type: "setCell",
        sheetId: sales.id,
        address: "B2",
        cell: { value: 100 },
      }),
    );
    expect(book.getCellValue(summary.id, "B1")).toEqual(importedCache);
    unwrap(book.recalculate());
    expect(book.getCellValue(summary.id, "B1")).toBe(1550);
    expect(book.exportSnapshot().sheets[1]?.cells.B1?.value).toBe(1550);
  });

  it("returns formula failures as cached cell values instead of hiding them", () => {
    const book = unwrap(
      createWorkbook(snapshot(2, 0, "=A1/A2"), { calculation: "manual" }),
    );
    unwrap(book.recalculate());
    expect(book.getCellValue(sheetId, "B1")).toEqual({
      kind: "error",
      code: "#DIV/0!",
    });
    expect(book.exportSnapshot().sheets[0]?.cells.B1?.formula).toBe("=A1/A2");
  });

  it.each(["manual", "automatic"] as const)(
    "rejects an exhausted budget atomically in %s mode and can retry",
    (calculation) => {
      const book = unwrap(createWorkbook(snapshot(), { calculation }));
      const before = book.getState();
      let notifications = 0;
      book.subscribe(() => {
        notifications += 1;
      });
      const beforeValue = book.getCellValue(sheetId, "B1");
      const failed = book.recalculate({ maxCellVisits: 2 });
      expect(failed.ok).toBe(false);
      if (!failed.ok) expect(failed.error.reason).toBe("limit");
      expect(book.getState()).toBe(before);
      expect(book.getCellValue(sheetId, "B1")).toEqual(beforeValue);
      expect(notifications).toBe(0);
      unwrap(book.recalculate({ maxCellVisits: 100 }));
      expect(book.getCellValue(sheetId, "B1")).toBe(5);
      expect(notifications).toBe(1);
    },
  );

  it("rejects array caches without changing snapshot, history or previously calculated values", () => {
    const functions = createFunctionRegistry();
    functions.set("ARRAY", () => ({ kind: "array", values: [[1, 2]] }));
    const book = unwrap(
      createWorkbook(snapshot(2, 3, "=ARRAY()"), { functions }),
    );
    const before = book.getState();
    const value = book.getCellValue(sheetId, "B1");
    const failed = book.recalculate();
    expect(failed.ok).toBe(false);
    if (!failed.ok) expect(failed.error.reason).toBe("unsupported");
    expect(book.getState()).toBe(before);
    expect(book.getCellValue(sheetId, "B1")).toBe(value);
    expect(book.undo()).toBe(false);
  });

  it("uses cached formula values for manual filtered-row navigation", () => {
    const initial = snapshot();
    const sheet = initial.sheets[0];
    if (sheet === undefined) throw new Error("Expected fixture sheet");
    const functions = createFunctionRegistry();
    let calls = 0;
    functions.set("PROBE", () => {
      calls += 1;
      return 42;
    });
    const book = unwrap(
      createWorkbook(
        {
          ...initial,
          sheets: [
            {
              ...sheet,
              cells: {
                ...sheet.cells,
                B2: { value: 999, formula: "=PROBE()" },
              },
              autoFilter: {
                range: {
                  start: { row: 0, column: 0 },
                  end: { row: 2, column: 1 },
                },
                columns: [
                  {
                    column: 1,
                    condition: {
                      kind: "values",
                      values: ["999"],
                      includeBlank: false,
                    },
                  },
                ],
              },
            },
          ],
        },
        { calculation: "manual", functions },
      ),
    );
    unwrap(book.selectCell(sheetId, { row: 0, column: 0 }));
    unwrap(book.moveSelection("down"));
    expect(book.getState().selection?.focus).toEqual({ row: 1, column: 0 });
    expect(calls).toBe(0);
  });

  it("rejects an invalid function result without committing any refreshed caches", () => {
    const functions = createFunctionRegistry();
    functions.set("INVALID", () => Number.NaN);
    const book = unwrap(
      createWorkbook(snapshot(2, 3, "=INVALID()"), {
        calculation: "manual",
        functions,
      }),
    );
    const before = book.getState();
    const failed = book.recalculate();
    expect(failed.ok).toBe(false);
    if (!failed.ok) expect(failed.error.reason).toBe("command");
    expect(book.getState()).toBe(before);
    expect(book.getCellValue(sheetId, "B1")).toBe(999);
    expect(book.undo()).toBe(false);
  });

  it.each([0, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY])(
    "rejects invalid calculation budget %s",
    (maxCellVisits) => {
      const book = unwrap(createWorkbook(snapshot()));
      const before = book.getState();
      expect(book.recalculate({ maxCellVisits }).ok).toBe(false);
      expect(book.getState()).toBe(before);
    },
  );

  it("preserves refreshed formula caches through XLSX serialization for generated numeric inputs", () => {
    fc.assert(
      fc.property(
        fc.integer({ min: -10_000, max: 10_000 }),
        fc.integer({ min: -10_000, max: 10_000 }),
        (a, b) => {
          const book = unwrap(
            createWorkbook(snapshot(a, b), { calculation: "manual" }),
          );
          unwrap(book.recalculate());
          const imported = unwrap(
            parseXlsx(unwrap(writeXlsx(book.exportSnapshot()))),
          );
          expect(imported.snapshot.sheets[0]?.cells.B1).toMatchObject({
            value: a + b,
            formula: "SUM(A1:A2)",
            style: { fontWeight: "bold" },
          });
        },
      ),
      { numRuns: 100 },
    );
  });
});
