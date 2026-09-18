import { describe, it, expect } from "vitest";
import fc from "fast-check";
import {
  createEmptySnapshot,
  parseWorkbookSnapshot,
  expandRangeToMerges,
} from "@nori-internal/model";
import {
  createWorkbook,
  getVisibleRange,
  getVisibleRows,
  getVisibleColumns,
  getColumnWidth,
  getRowHeight,
} from "@nori-internal/core";
import { unwrap } from "./helpers";
function bookWith(sheet: object = {}) {
  const base = createEmptySnapshot();
  return unwrap(
    createWorkbook({
      ...base,
      sheets: [{ ...base.sheets[0], rowCount: 12, columnCount: 8, ...sheet }],
    }),
  );
}
const range = (
  row: number,
  column: number,
  endRow: number,
  endColumn: number,
) => ({ start: { row, column }, end: { row: endRow, column: endColumn } });
describe("selection and document interactions", () => {
  it("keeps an anchor across reversed Shift selection, repeated extension and Shift-arrow contraction", () => {
    const book = bookWith(),
      id = book.getState().activeSheetId;
    unwrap(book.selectCell(id, { row: 4, column: 4 }));
    unwrap(book.selectCell(id, { row: 1, column: 1 }, { extend: true }));
    expect(book.getState().selection).toMatchObject({
      anchor: { row: 4, column: 4 },
      focus: { row: 1, column: 1 },
      range: range(1, 1, 4, 4),
    });
    unwrap(book.selectCell(id, { row: 2, column: 2 }, { extend: true }));
    unwrap(book.moveSelection("right", { extend: true }));
    expect(book.getState().selection).toMatchObject({
      anchor: { row: 4, column: 4 },
      focus: { row: 2, column: 3 },
      range: range(2, 3, 4, 4),
    });
    expect(book.getState().canUndo).toBe(false);
  });
  it("expands a selection through all intersecting merges and skips across a merge on arrow movement", () => {
    const book = bookWith({ merges: [range(1, 1, 3, 2), range(2, 3, 2, 4)] }),
      id = book.getState().activeSheetId;
    unwrap(book.setSelection({ sheetId: id, range: range(0, 0, 1, 3) }));
    expect(book.getState().selection?.range).toEqual(range(0, 0, 3, 4));
    unwrap(book.selectCell(id, { row: 3, column: 2 }));
    expect(book.getState().selection?.range).toEqual(range(1, 1, 3, 2));
    unwrap(book.moveSelection("right"));
    expect(book.getState().selection?.focus).toEqual({ row: 1, column: 3 });
  });
  it("merges and unmerges safely, redirects edits to the anchor, and supports undo/redo and JSON rehydration", () => {
    const book = bookWith({ cells: { B2: { value: "Title" } } }),
      id = book.getState().activeSheetId;
    unwrap(
      book.dispatch({
        type: "mergeCells",
        sheetId: id,
        range: range(1, 1, 2, 3),
      }),
    );
    unwrap(
      book.dispatch({
        type: "setCell",
        sheetId: id,
        address: "D3",
        cell: { value: "Renamed" },
      }),
    );
    expect(book.getCellValue(id, "B2")).toBe("Renamed");
    expect(book.getCellValue(id, "D3")).toBeNull();
    expect(
      unwrap(
        createWorkbook(JSON.parse(JSON.stringify(book.exportSnapshot()))),
      ).exportSnapshot(),
    ).toEqual(book.exportSnapshot());
    unwrap(
      book.dispatch({
        type: "unmergeCells",
        sheetId: id,
        range: range(2, 3, 2, 3),
      }),
    );
    expect(book.exportSnapshot().sheets[0]?.merges).toEqual([]);
    book.undo();
    expect(book.exportSnapshot().sheets[0]?.merges).toEqual([
      range(1, 1, 2, 3),
    ]);
    book.redo();
    expect(book.getCellValue(id, "B2")).toBe("Renamed");
  });
  it("rejects destructive merges, partial overlaps, invalid dimensions and failed batches atomically", () => {
    const book = bookWith({
        cells: { A1: { value: "A" }, B1: { value: "B" } },
        merges: [range(3, 3, 4, 4)],
      }),
      id = book.getState().activeSheetId,
      before = book.getState();
    expect(
      book.dispatch({
        type: "mergeCells",
        sheetId: id,
        range: range(0, 0, 0, 1),
      }).ok,
    ).toBe(false);
    expect(
      book.dispatch({
        type: "mergeCells",
        sheetId: id,
        range: range(2, 2, 3, 3),
      }).ok,
    ).toBe(false);
    expect(
      book.dispatch({
        type: "batch",
        commands: [
          { type: "resizeColumn", sheetId: id, column: 0, width: 240 },
          { type: "resizeRow", sheetId: id, row: 0, height: 0 },
        ],
      }).ok,
    ).toBe(false);
    expect(book.getState()).toBe(before);
  });
  it("resizes both axes, resets to imported defaults, and restores sizes through history", () => {
    const book = bookWith({ defaultColumnWidth: 90, defaultRowHeight: 24 }),
      id = book.getState().activeSheetId;
    unwrap(
      book.dispatch({
        type: "batch",
        commands: [
          { type: "resizeColumn", sheetId: id, column: 2, width: 240 },
          { type: "resizeRow", sheetId: id, row: 3, height: 72 },
        ],
      }),
    );
    const sheet = book.exportSnapshot().sheets[0];
    if (!sheet) throw Error("fixture");
    expect(getColumnWidth(sheet, 2)).toBe(240);
    expect(getRowHeight(sheet, 3)).toBe(72);
    book.undo();
    expect(book.exportSnapshot().sheets[0]?.columnWidths).toBeUndefined();
    book.redo();
    unwrap(
      book.dispatch({
        type: "resizeColumn",
        sheetId: id,
        column: 2,
        width: null,
      }),
    );
    const reset = book.exportSnapshot().sheets[0];
    if (!reset) throw Error("fixture");
    expect(getColumnWidth(reset, 2)).toBe(90);
  });
  it("maps viewports with sparse dimensions at exact boundaries", () => {
    expect(
      unwrap(
        getVisibleRange(
          {
            rowHeight: 20,
            columnWidth: 80,
            rowHeights: { 0: 40 },
            columnWidths: { 0: 160 },
          },
          { x: 160, y: 40, width: 80, height: 20 },
          { rows: 10, columns: 10 },
        ),
      ),
    ).toEqual(range(1, 1, 1, 1));
  });
  it("range expansion is idempotent for arbitrary selections", () => {
    const sheet = bookWith({
      merges: [range(1, 1, 3, 2), range(5, 4, 7, 6)],
    }).exportSnapshot().sheets[0];
    if (!sheet) throw Error("fixture");
    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: 11 }),
        fc.integer({ min: 0, max: 7 }),
        fc.integer({ min: 0, max: 11 }),
        fc.integer({ min: 0, max: 7 }),
        (a, b, c, d) => {
          const once = expandRangeToMerges(sheet, range(a, b, c, d));
          expect(expandRangeToMerges(sheet, once)).toEqual(once);
        },
      ),
    );
  });
  it("validates merge invariants and view metadata at the snapshot boundary", () => {
    const base = bookWith().exportSnapshot(),
      sheet = base.sheets[0];
    for (const patch of [
      { merges: [range(0, 0, 1, 1), range(1, 1, 2, 2)] },
      { merges: [range(0, 0, 0, 0)] },
      { merges: [range(0, 0, 99, 2)] },
      { columnWidths: { "-1": 200 } },
      { hiddenRows: [100] },
      { frozen: { rows: 99, columns: 0 } },
    ])
      expect(
        parseWorkbookSnapshot({ ...base, sheets: [{ ...sheet, ...patch }] }).ok,
      ).toBe(false);
  });
});
describe("saved sheet views", () => {
  it("combines hidden dimensions with value and custom filters and navigates visible rows", () => {
    const book = bookWith({
        cells: {
          A1: { value: "Region" },
          B1: { value: "Revenue" },
          A2: { value: "West" },
          B2: { value: 5 },
          A3: { value: "East" },
          B3: { value: 12 },
          A4: { value: "West" },
          B4: { value: 20 },
          A5: { value: "West" },
          B5: { value: 50 },
        },
        hiddenRows: [4],
        hiddenColumns: [2],
        autoFilter: {
          range: range(0, 0, 11, 1),
          columns: [
            {
              column: 0,
              condition: {
                kind: "values",
                values: ["West"],
                includeBlank: false,
              },
            },
            {
              column: 1,
              condition: {
                kind: "custom",
                join: "and",
                comparisons: [{ operator: "greaterThan", value: 10 }],
              },
            },
          ],
        },
      }),
      id = book.getState().activeSheetId;
    const sheet = book.exportSnapshot().sheets[0];
    if (!sheet) throw Error("fixture");
    expect(getVisibleRows(sheet)).toEqual([0, 3]);
    expect(getVisibleColumns(sheet)).toEqual([0, 1, 3, 4, 5, 6, 7]);
    unwrap(book.selectCell(id, { row: 0, column: 1 }));
    unwrap(book.moveSelection("down"));
    expect(book.getState().selection?.focus.row).toBe(3);
    unwrap(book.moveSelection("right"));
    expect(book.getState().selection?.focus.column).toBe(3);
  });
  it("starts keyboard navigation at a visible cell and leaves fully hidden views unselected", () => {
    const book = bookWith({ hiddenRows: [0, 1], hiddenColumns: [0] });
    unwrap(book.moveSelection("right"));
    expect(book.getState().selection?.focus).toEqual({ row: 2, column: 1 });
    const hidden = bookWith({ rowCount: 2, hiddenRows: [0, 1] });
    unwrap(hidden.moveSelection("down"));
    expect(hidden.getState().selection).toBeNull();
  });
  it("matches wildcard escaping, OR comparisons and explicit blanks", () => {
    const sheet = bookWith({
      rowCount: 5,
      cells: {
        A1: { value: "Name" },
        A2: { value: "North" },
        A3: { value: "south" },
        A4: { value: "N*" },
        A5: { value: null },
      },
      autoFilter: {
        range: range(0, 0, 4, 0),
        columns: [
          {
            column: 0,
            condition: {
              kind: "custom",
              join: "or",
              comparisons: [
                { operator: "equal", value: "n~*" },
                { operator: "equal", value: "S?uth" },
              ],
            },
          },
        ],
      },
    }).exportSnapshot().sheets[0];
    if (!sheet) throw Error("fixture");
    expect(getVisibleRows(sheet)).toEqual([0, 2, 3]);
    expect(
      getVisibleRows({
        ...sheet,
        autoFilter: {
          range: range(0, 0, 4, 0),
          columns: [
            {
              column: 0,
              condition: { kind: "values", values: [], includeBlank: true },
            },
          ],
        },
      }),
    ).toEqual([0, 4]);
  });
});
