import { describe, it, expect } from "vitest";
import fc from "fast-check";
import { createWorkbook, getVisibleRange } from "@nori-internal/core";
import { cellError, parseSheetId } from "@nori-internal/model";
import { unwrap, workbook } from "./helpers";
describe("workbook runtime", () => {
  it("recalculates cross-sheet dependencies and undo/redo", () => {
    const book = workbook();
    const [sales, summary] = book.getState().snapshot.sheets;
    if (!sales || !summary) throw Error("fixture");
    expect(book.getCellValue(summary.id, "B1")).toBe(2650);
    unwrap(
      book.dispatch({
        type: "setCell",
        sheetId: sales.id,
        address: "b2",
        cell: { value: 100 },
      }),
    );
    expect(book.getCellValue(summary.id, "B1")).toBe(1550);
    expect(book.undo()).toBe(true);
    expect(book.getCellValue(summary.id, "B1")).toBe(2650);
    expect(book.redo()).toBe(true);
    expect(book.getCellValue(summary.id, "B1")).toBe(1550);
  });
  it("commits batches atomically and notifies once", () => {
    const book = workbook(),
      id = book.getState().activeSheetId;
    let count = 0;
    const stop = book.subscribe(() => count++);
    const before = book.getState();
    const result = book.dispatch({
      type: "batch",
      commands: [
        { type: "setCell", sheetId: id, address: "A1", cell: { value: 9 } },
        { type: "clearCell", sheetId: id, address: "A0" },
      ],
    });
    expect(result.ok).toBe(false);
    expect(book.getState()).toBe(before);
    expect(count).toBe(0);
    unwrap(
      book.dispatch({
        type: "batch",
        commands: [
          { type: "clearCell", sheetId: id, address: "A1" },
          { type: "setCell", sheetId: id, address: "B2", cell: { value: 4 } },
        ],
      }),
    );
    expect(count).toBe(1);
    expect(book.getCellValue(id, "A1")).toBeNull();
    book.undo();
    expect(count).toBe(2);
    stop();
    book.redo();
    expect(count).toBe(2);
  });
  it("redo is discarded by a new command and history is bounded", () => {
    const original = workbook();
    const book = unwrap(
      createWorkbook(original.exportSnapshot(), { historyLimit: 1 }),
    );
    const id = book.getState().activeSheetId;
    for (const value of [1, 2])
      unwrap(
        book.dispatch({
          type: "setCell",
          sheetId: id,
          address: "A1",
          cell: { value },
        }),
      );
    expect(book.undo()).toBe(true);
    expect(book.undo()).toBe(false);
    unwrap(
      book.dispatch({
        type: "setCell",
        sheetId: id,
        address: "A1",
        cell: { value: 3 },
      }),
    );
    expect(book.redo()).toBe(false);
  });
  it("detects cycles, missing sheets, parse errors and large ranges", () => {
    const book = workbook(),
      id = book.getState().activeSheetId;
    unwrap(
      book.dispatch({
        type: "batch",
        commands: [
          {
            type: "setCell",
            sheetId: id,
            address: "D1",
            cell: { value: null, formula: "D2" },
          },
          {
            type: "setCell",
            sheetId: id,
            address: "D2",
            cell: { value: null, formula: "D1" },
          },
          {
            type: "setCell",
            sheetId: id,
            address: "E1",
            cell: { value: null, formula: "Missing!A1" },
          },
          {
            type: "setCell",
            sheetId: id,
            address: "E2",
            cell: { value: null, formula: "SUM(" },
          },
          {
            type: "setCell",
            sheetId: id,
            address: "E3",
            cell: { value: null, formula: "SUM(A1:XFD1048576)" },
          },
        ],
      }),
    );
    expect(book.getCellValue(id, "D1")).toEqual(cellError("#CYCLE!"));
    expect(book.getCellValue(id, "E1")).toEqual(cellError("#REF!"));
    expect(book.getCellValue(id, "E2")).toEqual(cellError("#ERROR!"));
    expect(book.getCellValue(id, "E3")).toEqual(cellError("#NUM!"));
    unwrap(
      book.dispatch({
        type: "setCell",
        sheetId: id,
        address: "D2",
        cell: { value: 7 },
      }),
    );
    expect(book.getCellValue(id, "D1")).toBe(7);
  });
  it("selection is copied, bounds checked, and not in history", () => {
    const book = workbook(),
      id = book.getState().activeSheetId;
    const position = { row: 1, column: 1 };
    unwrap(
      book.setSelection({
        sheetId: id,
        range: { start: position, end: position },
      }),
    );
    position.row = 999;
    expect(book.getState().selection?.range.start.row).toBe(1);
    expect(book.getState().canUndo).toBe(false);
    expect(
      book.setSelection({
        sheetId: id,
        range: { start: position, end: position },
      }).ok,
    ).toBe(false);
    expect(book.setActiveSheet(unwrap(parseSheetId("missing"))).ok).toBe(false);
    const second = book.getState().snapshot.sheets[1];
    if (!second) throw Error("fixture");
    unwrap(book.setActiveSheet(second.id));
    expect(book.getState().selection).toBeNull();
  });
  it("exports defensively frozen JSON that rehydrates", () => {
    const book = workbook();
    expect(
      unwrap(
        createWorkbook(JSON.parse(JSON.stringify(book.exportSnapshot()))),
      ).exportSnapshot(),
    ).toEqual(book.exportSnapshot());
    expect(Object.isFrozen(book.exportSnapshot().sheets[0]?.cells["B2"])).toBe(
      true,
    );
  });
  it("undo restores arbitrary scalar edits", () =>
    fc.assert(
      fc.property(fc.integer(), (value) => {
        const book = workbook(),
          id = book.getState().activeSheetId,
          before = book.exportSnapshot();
        unwrap(
          book.dispatch({
            type: "setCell",
            sheetId: id,
            address: "C5",
            cell: { value },
          }),
        );
        expect(book.getCellValue(id, "C5")).toBe(value);
        book.undo();
        expect(book.exportSnapshot()).toEqual(before);
      }),
      { numRuns: 25 },
    ));
  it("maps viewports without DOM and rejects invalid measurements", () => {
    expect(
      unwrap(
        getVisibleRange(
          { rowHeight: 20, columnWidth: 80 },
          { x: 80, y: 40, width: 160, height: 60 },
          { rows: 100, columns: 10 },
        ),
      ),
    ).toEqual({ start: { row: 2, column: 1 }, end: { row: 4, column: 2 } });
    expect(
      getVisibleRange(
        { rowHeight: 0, columnWidth: 80 },
        { x: 0, y: 0, width: 80, height: 20 },
        { rows: 10, columns: 10 },
      ).ok,
    ).toBe(false);
  });
});
