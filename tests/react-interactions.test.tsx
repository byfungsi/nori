// @vitest-environment jsdom
import { afterEach, it, expect } from "vitest";
import {
  act,
  cleanup,
  fireEvent,
  render,
  screen,
  within,
} from "@testing-library/react";
import { createWorkbook } from "@nori-internal/core";
import { createEmptySnapshot } from "@nori-internal/model";
import {
  Spreadsheet,
  SpreadsheetPreview,
  WorkbookProvider,
  SheetGrid,
} from "@nori-internal/react";
import { unwrap, workbook } from "./helpers";
afterEach(cleanup);
function empty(sheet: object = {}) {
  const snapshot = createEmptySnapshot();
  return unwrap(
    createWorkbook({
      ...snapshot,
      sheets: [
        { ...snapshot.sheets[0], rowCount: 6, columnCount: 5, ...sheet },
      ],
    }),
  );
}
function pointer(
  node: Element,
  type: string,
  x: number,
  y: number,
  options: { shiftKey?: boolean; button?: number } = {},
) {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    clientX: x,
    clientY: y,
    button: options.button ?? 0,
    buttons: type === "pointerup" ? 0 : 1,
    ...options,
  });
  Object.defineProperties(event, {
    pointerId: { value: 1 },
    isPrimary: { value: true },
    pointerType: { value: "mouse" },
  });
  fireEvent(node, event);
}
// jsdom supplies events but no layout. These rectangles model the real table's
// geometry; actual pointer capture, scrolling and dimensions are browser-verified.
function tableGeometry(table: HTMLElement) {
  for (const th of table.querySelectorAll<HTMLElement>(
    "thead [data-column-index]",
  ))
    Object.defineProperty(th, "getBoundingClientRect", {
      configurable: true,
      value: () =>
        new DOMRect(42 + Number(th.dataset["columnIndex"]) * 144, 0, 144, 32),
    });
  for (const tr of table.querySelectorAll<HTMLElement>(
    "tbody tr[data-row-index]",
  ))
    Object.defineProperty(tr, "getBoundingClientRect", {
      configurable: true,
      value: () =>
        new DOMRect(0, 32 + Number(tr.dataset["rowIndex"]) * 34, 762, 34),
    });
}
it("drag-selects a rectangle without collapsing on click and stops at pointer cancel", () => {
  const book = empty();
  render(<Spreadsheet workbook={book} />);
  const table = screen.getByRole("table"),
    scroll = table.parentElement;
  if (!scroll) throw Error("fixture");
  tableGeometry(table);
  pointer(screen.getByRole("button", { name: "B2:" }), "pointerdown", 258, 83);
  pointer(scroll, "pointermove", 546, 151);
  pointer(scroll, "pointerup", 546, 151);
  fireEvent.click(screen.getByRole("button", { name: "D4:" }), { detail: 1 });
  expect(book.getState().selection?.range).toEqual({
    start: { row: 1, column: 1 },
    end: { row: 3, column: 3 },
  });
  expect(table.querySelectorAll("td[data-selected]")).toHaveLength(9);
  pointer(screen.getByRole("button", { name: "A1:" }), "pointerdown", 114, 49);
  pointer(scroll, "pointercancel", 114, 49);
  pointer(scroll, "pointermove", 546, 151);
  expect(book.getState().selection?.range.end).toEqual({ row: 0, column: 0 });
});
it("supports Shift-click, reversed extension and Shift-arrow with a stable anchor", () => {
  const book = empty();
  render(<Spreadsheet workbook={book} />);
  fireEvent.click(screen.getByRole("button", { name: "D4:" }));
  fireEvent.click(screen.getByRole("button", { name: "B2:" }), {
    shiftKey: true,
  });
  fireEvent.keyDown(screen.getByRole("button", { name: "B2:" }), {
    key: "ArrowRight",
    shiftKey: true,
  });
  expect(book.getState().selection).toMatchObject({
    anchor: { row: 3, column: 3 },
    focus: { row: 1, column: 2 },
    range: { start: { row: 1, column: 2 }, end: { row: 3, column: 3 } },
  });
});
it("previews drag resizing without commands and commits one undoable change; Escape cancels", () => {
  const book = empty();
  render(<Spreadsheet workbook={book} />);
  const scroll = screen.getByRole("table").parentElement;
  if (!scroll) throw Error("fixture");
  const column = screen.getByRole("separator", { name: "Resize column B" });
  pointer(column, "pointerdown", 330, 10);
  pointer(scroll, "pointermove", 416, 10);
  expect(book.getState().canUndo).toBe(false);
  expect(column.getAttribute("aria-valuenow")).toBe("230");
  pointer(scroll, "pointerup", 416, 10);
  expect(book.exportSnapshot().sheets[0]?.columnWidths?.["1"]).toBe(230);
  act(() => {
    book.undo();
  });
  expect(book.getState().canUndo).toBe(false);
  const row = screen.getByRole("separator", { name: "Resize row 2" });
  pointer(row, "pointerdown", 20, 100);
  pointer(scroll, "pointermove", 20, 140);
  fireEvent.keyDown(scroll, { key: "Escape" });
  expect(book.exportSnapshot().sheets[0]?.rowHeights).toBeUndefined();
  fireEvent.keyDown(row, { key: "ArrowDown" });
  expect(book.exportSnapshot().sheets[0]?.rowHeights?.["1"]).toBe(44);
  fireEvent.doubleClick(row);
  expect(book.exportSnapshot().sheets[0]?.rowHeights?.["1"]).toBeUndefined();
});
it("merges, edits through the anchor and unmerges via composable controls", () => {
  const book = empty({ cells: { A1: { value: "Title" } } });
  render(<Spreadsheet workbook={book} />);
  fireEvent.click(screen.getByRole("button", { name: "A1: Title" }));
  fireEvent.click(screen.getByRole("button", { name: "C2:" }), {
    shiftKey: true,
  });
  fireEvent.click(screen.getByRole("button", { name: "Merge cells" }));
  const merged = screen
    .getByRole("button", { name: "A1: Title" })
    .closest("td");
  expect(merged?.colSpan).toBe(3);
  expect(merged?.rowSpan).toBe(2);
  expect(screen.queryByRole("button", { name: "B1:" })).toBeNull();
  fireEvent.change(screen.getByRole("textbox"), {
    target: { value: "Changed" },
  });
  fireEvent.click(screen.getByRole("button", { name: "Apply" }));
  expect(screen.getByRole("button", { name: "A1: Changed" })).toBeTruthy();
  fireEvent.click(screen.getByRole("button", { name: "Unmerge cells" }));
  expect(screen.getByRole("button", { name: "B1:" })).toBeTruthy();
});
it("renders a merge anchor through a clipped window and hidden dimensions", () => {
  const book = empty({
    merges: [{ start: { row: 0, column: 0 }, end: { row: 3, column: 3 } }],
    cells: { A1: { value: "Title" } },
    hiddenRows: [0],
    hiddenColumns: [0],
  });
  render(
    <WorkbookProvider workbook={book}>
      <SheetGrid
        range={{ start: { row: 2, column: 2 }, end: { row: 3, column: 3 } }}
      />
    </WorkbookProvider>,
  );
  const cell = screen.getByRole("button", { name: "A1: Title" }).closest("td");
  expect(cell?.rowSpan).toBe(2);
  expect(cell?.colSpan).toBe(2);
});
it("honors hidden, filtered and frozen rows/columns and saved sort indicators", () => {
  const book = empty({
    frozen: { rows: 1, columns: 1 },
    hiddenColumns: [2],
    hiddenRows: [4],
    cells: {
      A1: { value: "Region" },
      A2: { value: "West" },
      A3: { value: "East" },
      A4: { value: "West" },
    },
    autoFilter: {
      range: { start: { row: 0, column: 0 }, end: { row: 5, column: 0 } },
      columns: [
        {
          column: 0,
          condition: { kind: "values", values: ["West"], includeBlank: false },
        },
      ],
    },
    sort: {
      range: { start: { row: 1, column: 0 }, end: { row: 5, column: 0 } },
      caseSensitive: false,
      conditions: [{ column: 0, direction: "ascending" }],
    },
  });
  render(<Spreadsheet workbook={book} />);
  expect(screen.queryByRole("button", { name: "A3: East" })).toBeNull();
  expect(
    screen.queryByRole("separator", { name: "Resize column C" }),
  ).toBeNull();
  const cell = screen.getByRole("button", { name: "A1: Region" }).closest("td");
  expect(cell?.style.position).toBe("sticky");
  expect(cell?.style.top).toBe("32px");
  expect(cell?.style.left).toBe("42px");
  expect(
    screen.getByRole("columnheader", { name: /A/ }).getAttribute("aria-sort"),
  ).toBe("ascending");
});
it("keeps preview tabs local, remains read-only, opens on request and subscribes to changes", () => {
  const book = workbook(),
    before = book.getState();
  let opened = 0;
  render(
    <SpreadsheetPreview
      workbook={book}
      title="report.xlsx"
      theme="dark"
      maxRows={3}
      onOpen={() => opened++}
    />,
  );
  expect(screen.queryByRole("textbox")).toBeNull();
  expect(screen.queryByRole("separator")).toBeNull();
  expect(screen.queryByRole("button", { name: "Merge cells" })).toBeNull();
  fireEvent.click(screen.getByRole("button", { name: "Summary" }));
  expect(screen.getByRole("table", { name: "Summary preview" })).toBeTruthy();
  expect(book.getState()).toBe(before);
  fireEvent.click(screen.getByRole("button", { name: "Open workbook" }));
  expect(opened).toBe(1);
  act(() => {
    unwrap(
      book.dispatch({
        type: "setCell",
        sheetId: before.activeSheetId,
        address: "B2",
        cell: { value: 100 },
      }),
    );
  });
  expect(within(screen.getByRole("table")).getByText("1550")).toBeTruthy();
});
it("resizes from both sides of body borders and highlights the whole boundary", () => {
  const book = empty();
  const { container } = render(<Spreadsheet workbook={book} />);
  const table = screen.getByRole("table"),
    scroll = table.parentElement;
  if (!scroll) throw Error("fixture");
  tableGeometry(table);
  Object.defineProperties(scroll, {
    clientWidth: { value: 762 },
    clientHeight: { value: 236 },
  });
  const a = screen.getByRole("button", { name: "A2:" }),
    b = screen.getByRole("button", { name: "B2:" });
  for (const [button, x] of [
    [a, 42],
    [b, 186],
  ] as const) {
    const cell = button.parentElement;
    if (!cell) throw Error("fixture");
    cell.getBoundingClientRect = () => new DOMRect(x, 66, 144, 34);
  }
  pointer(a, "pointermove", 182, 82);
  expect(
    container.querySelector(".nori-resize-guide")?.getAttribute("data-axis"),
  ).toBe("column");
  pointer(b, "pointermove", 190, 82);
  expect(
    container.querySelector(".nori-resize-guide")?.getAttribute("data-index"),
  ).toBe("0");
  pointer(b, "pointerdown", 190, 82);
  pointer(scroll, "pointermove", 220, 82);
  expect(book.getState().selection).toBeNull();
  expect(container.querySelector(".nori-resize-guide")).not.toBeNull();
  pointer(scroll, "pointerup", 220, 82);
  expect(book.exportSnapshot().sheets[0]?.columnWidths?.["0"]).toBe(174);
  pointer(a, "pointermove", 100, 96);
  expect(
    container.querySelector(".nori-resize-guide")?.getAttribute("data-axis"),
  ).toBe("row");
  pointer(a, "pointerdown", 100, 96);
  pointer(scroll, "pointerup", 100, 116);
  expect(book.exportSnapshot().sheets[0]?.rowHeights?.["1"]).toBe(54);
  pointer(a, "pointermove", 100, 82);
  expect(container.querySelector(".nori-resize-guide")).toBeNull();
});
it("edits a cell on double-click, commits formulas and cancels with Escape", () => {
  const book = empty({
    cells: { A1: { value: 5, style: { fontWeight: "bold" } } },
  });
  render(<Spreadsheet workbook={book} />);
  fireEvent.doubleClick(screen.getByRole("button", { name: "A1: 5" }));
  const input = screen.getByRole("textbox", { name: "Edit A1" });
  fireEvent.change(input, { target: { value: "=2+3" } });
  fireEvent.keyDown(input, { key: "Enter" });
  const id = book.getState().activeSheetId;
  expect(book.exportSnapshot().sheets[0]?.cells["A1"]).toEqual({
    value: null,
    formula: "2+3",
    style: { fontWeight: "bold" },
  });
  expect(book.getCellValue(id, "A1")).toBe(5);
  fireEvent.keyDown(screen.getByRole("button", { name: "A1: 5" }), {
    key: "F2",
  });
  const second = screen.getByRole("textbox", { name: "Edit A1" });
  fireEvent.change(second, { target: { value: "999" } });
  fireEvent.keyDown(second, { key: "Escape" });
  expect(book.getCellValue(id, "A1")).toBe(5);
  act(() => {
    expect(book.undo()).toBe(true);
  });
  expect(book.exportSnapshot().sheets[0]?.cells["A1"]?.formula).toBeUndefined();
});
it("readOnly blocks editing, resize and merge while preserving selection and tabs", () => {
  const book = empty();
  const { rerender, container } = render(<Spreadsheet workbook={book} />);
  fireEvent.doubleClick(screen.getByRole("button", { name: "A1:" }));
  fireEvent.change(screen.getByRole("textbox", { name: "Edit A1" }), {
    target: { value: "discard me" },
  });
  rerender(<Spreadsheet workbook={book} readOnly />);
  expect(screen.queryByRole("textbox", { name: "Edit A1" })).toBeNull();
  const snapshot = book.exportSnapshot();
  fireEvent.click(screen.getByRole("button", { name: "A1:" }));
  fireEvent.click(screen.getByRole("button", { name: "B2:" }), {
    shiftKey: true,
  });
  fireEvent.doubleClick(screen.getByRole("button", { name: "A1:" }));
  expect(screen.queryByRole("textbox", { name: "Edit A1" })).toBeNull();
  expect(screen.queryAllByRole("separator")).toHaveLength(0);
  expect(
    screen
      .getByRole("textbox", { name: "Cell value or formula" })
      .getAttribute("readonly"),
  ).not.toBeNull();
  const form = container.querySelector("form");
  if (!form) throw Error("fixture");
  fireEvent.submit(form);
  fireEvent.click(screen.getByRole("button", { name: "Merge cells" }));
  expect(book.exportSnapshot()).toBe(snapshot);
  expect(book.getState().selection?.range.end).toEqual({ row: 1, column: 1 });
  expect(book.getState().canUndo).toBe(false);
});
