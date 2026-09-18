import {
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type CSSProperties,
  type ReactNode,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
} from "react";
import {
  addressOf,
  mergedRangeAt,
  rangeContains,
  rangesIntersect,
  type Cell,
  type CellRange,
  type Position,
  type SheetSnapshot,
} from "@nori-internal/model";
import {
  getColumnWidth,
  getRowHeight,
  getVisibleRows,
  getVisibleColumns,
  type CommandError,
  type Workbook,
} from "@nori-internal/core";
import type { FormulaResult } from "@nori-internal/formula";
import {
  useWorkbook,
  useWorkbookState,
  useWorkbookReadOnly,
} from "./workbook-context";
import { InlineCellEditor, editedCell } from "./cell-edit";
import { cellStyleToCss, formatCellValue } from "./cell-format";

/** Custom cells receive the merge anchor's value even in a clipped viewport. */
export interface CellRenderContext {
  readonly sheet: SheetSnapshot;
  readonly address: string;
  readonly cell: Cell | undefined;
  readonly value: FormulaResult;
  readonly selected: boolean;
  readonly mergedRange: CellRange | undefined;
}
/** Render a bounded window; layout and selection remain in the headless workbook. */
export interface SheetGridProps {
  readonly className?: string;
  readonly readOnly?: boolean;
  readonly range?: CellRange;
  readonly renderCell?: (context: CellRenderContext) => ReactNode;
  readonly onError?: (error: CommandError) => void;
}
type ResizeBoundary = { axis: "column" | "row"; index: number };
type ResizePreview = { axis: "column" | "row"; index: number; size: number };
type Gesture = { pointerId: number; clientX: number; clientY: number } & (
  | { kind: "select" }
  | {
      kind: "resize";
      axis: "column" | "row";
      index: number;
      origin: number;
      originalSize: number;
      size: number;
    }
);
const minSize = (axis: "column" | "row") => (axis === "column" ? 40 : 20);
function resizedSize(
  original: number,
  delta: number,
  axis: "column" | "row",
): number {
  return delta === 0
    ? original
    : Math.max(minSize(axis), Math.min(10000, Math.round(original + delta)));
}

// Read rendered header rectangles: this respects scroll position, zoom, dimensions,
// clipped merges and host styling without putting DOM measurements into core.
function hitPosition(
  table: HTMLTableElement,
  clientX: number,
  clientY: number,
): Position | undefined {
  const columns = [
    ...table.querySelectorAll<HTMLElement>("thead [data-column-index]"),
  ];
  const rows = [
    ...table.querySelectorAll<HTMLElement>("tbody tr[data-row-index]"),
  ];
  const col =
    columns.find(
      (element) => clientX < element.getBoundingClientRect().right,
    ) ?? columns.at(-1);
  const row =
    rows.find((element) => clientY < element.getBoundingClientRect().bottom) ??
    rows.at(-1);
  return col && row
    ? {
        row: Number(row.dataset["rowIndex"]),
        column: Number(col.dataset["columnIndex"]),
      }
    : undefined;
}
// Resolve the visible cell perimeter rather than every table coordinate: internal
// boundaries of merged cells and boundaries covered by frozen panes are not handles.
function resizeBoundaryAt(
  event: ReactPointerEvent,
  rows: readonly number[],
  columns: readonly number[],
): ResizeBoundary | null {
  if (!(event.target instanceof Element)) return null;
  const handle = event.target.closest<HTMLElement>("[data-resize-axis]");
  if (handle) {
    const axis = handle.dataset["resizeAxis"];
    if (axis === "row" || axis === "column")
      return { axis, index: Number(handle.dataset["resizeIndex"]) };
  }
  const cell = event.target.closest<HTMLElement>("td, th");
  if (!cell) return null;
  const rect = cell.getBoundingClientRect();
  if (!rect.width || !rect.height) return null;
  const candidates: (ResizeBoundary & { distance: number })[] = [];
  for (const axis of ["column", "row"] as const) {
    const start = cell.dataset[axis + "Start"],
      end = cell.dataset[axis + "End"];
    if (start === undefined || end === undefined) continue;
    const visible = axis === "column" ? columns : rows;
    const previous = visible[visible.indexOf(Number(start)) - 1];
    const coordinate = axis === "column" ? event.clientX : event.clientY;
    const near = axis === "column" ? rect.left : rect.top;
    const far = axis === "column" ? rect.right : rect.bottom;
    if (previous !== undefined)
      candidates.push({
        axis,
        index: previous,
        distance: Math.abs(coordinate - near),
      });
    candidates.push({
      axis,
      index: Number(end),
      distance: Math.abs(coordinate - far),
    });
  }
  const nearest = candidates
    .filter(({ distance }) => distance <= 6)
    .sort((a, b) => a.distance - b.distance)[0];
  return nearest ? { axis: nearest.axis, index: nearest.index } : null;
}
function useGridGestures(
  workbook: Workbook,
  sheet: SheetSnapshot,
  scrollRef: RefObject<HTMLDivElement | null>,
  tableRef: RefObject<HTMLTableElement | null>,
  report: (error: CommandError) => void,
  readOnly: boolean,
) {
  const captured = useRef<Element | null>(null);
  const gesture = useRef<Gesture | null>(null),
    frame = useRef<number | null>(null);
  const [preview, setPreview] = useState<ResizePreview | null>(null);
  const stopFrame = () => {
    const view = scrollRef.current?.ownerDocument.defaultView;
    if (frame.current !== null) view?.cancelAnimationFrame(frame.current);
    frame.current = null;
  };
  const stop = (commit: boolean) => {
    const active = gesture.current;
    gesture.current = null;
    stopFrame();
    setPreview(null);
    const target = captured.current;
    captured.current = null;
    if (active && target?.hasPointerCapture?.(active.pointerId))
      target.releasePointerCapture(active.pointerId);
    if (
      !readOnly &&
      commit &&
      active?.kind === "resize" &&
      active.size !== active.originalSize
    ) {
      const result =
        active.axis === "column"
          ? workbook.dispatch({
              type: "resizeColumn",
              sheetId: sheet.id,
              column: active.index,
              width: active.size,
            })
          : workbook.dispatch({
              type: "resizeRow",
              sheetId: sheet.id,
              row: active.index,
              height: active.size,
            });
      if (!result.ok) report(result.error);
    }
  };
  useEffect(() => {
    const view = scrollRef.current?.ownerDocument.defaultView;
    const cancel = () => stop(false);
    view?.addEventListener("blur", cancel);
    return () => {
      view?.removeEventListener("blur", cancel);
      stop(false);
    };
  }, [workbook, sheet.id, readOnly]);
  const selectAtPointer = (active: Gesture) => {
    const table = tableRef.current;
    if (!table) return;
    const position = hitPosition(table, active.clientX, active.clientY);
    if (position) {
      const result = workbook.selectCell(sheet.id, position, { extend: true });
      if (!result.ok) report(result.error);
    }
  };
  const autoScroll = () => {
    const scroll = scrollRef.current,
      view = scroll?.ownerDocument.defaultView;
    if (!scroll || !view?.requestAnimationFrame) return;
    const tick = () => {
      const active = gesture.current;
      if (active?.kind !== "select") {
        frame.current = null;
        return;
      }
      const rect = scroll.getBoundingClientRect();
      const speed = (value: number, min: number, max: number) =>
        value < min + 24
          ? -Math.min(18, (min + 24 - value) / 2)
          : value > max - 24
            ? Math.min(18, (value - max + 24) / 2)
            : 0;
      const x = scroll.scrollLeft,
        y = scroll.scrollTop;
      if (rect.width > 0 && rect.height > 0) {
        scroll.scrollLeft += speed(active.clientX, rect.left, rect.right);
        scroll.scrollTop += speed(active.clientY, rect.top + 32, rect.bottom);
      }
      if (x !== scroll.scrollLeft || y !== scroll.scrollTop)
        selectAtPointer(active);
      frame.current = view.requestAnimationFrame(tick);
    };
    frame.current = view.requestAnimationFrame(tick);
  };
  const capture = (
    event: ReactPointerEvent,
    target: Element | null = scrollRef.current,
  ) => {
    captured.current = target;
    target?.setPointerCapture?.(event.pointerId);
  };
  const beginSelection = (
    event: ReactPointerEvent<HTMLButtonElement>,
    position: Position,
  ) => {
    if (event.button !== 0 || event.isPrimary === false) return;
    stop(false);
    event.preventDefault();
    event.currentTarget.focus({ preventScroll: true });
    const result = workbook.selectCell(sheet.id, position, {
      extend: event.shiftKey,
    });
    if (!result.ok) {
      report(result.error);
      return;
    }
    gesture.current = {
      kind: "select",
      pointerId: event.pointerId,
      clientX: event.clientX,
      clientY: event.clientY,
    };
    // Keep clicks targeted at the cell so the browser can produce dblclick.
    capture(event, event.currentTarget);
    autoScroll();
  };
  const beginResize = (
    event: ReactPointerEvent,
    axis: "column" | "row",
    index: number,
    size: number,
  ) => {
    if (readOnly) return;
    if (event.button !== 0 || event.isPrimary === false) return;
    event.preventDefault();
    event.stopPropagation();
    stop(false);
    gesture.current = {
      kind: "resize",
      pointerId: event.pointerId,
      clientX: event.clientX,
      clientY: event.clientY,
      axis,
      index,
      origin: axis === "column" ? event.clientX : event.clientY,
      originalSize: size,
      size,
    };
    setPreview({ axis, index, size });
    capture(event);
  };
  const move = (event: ReactPointerEvent) => {
    const active = gesture.current;
    if (!active || active.pointerId !== event.pointerId) return;
    active.clientX = event.clientX;
    active.clientY = event.clientY;
    if (active.kind === "select") selectAtPointer(active);
    else {
      active.size = resizedSize(
        active.originalSize,
        (active.axis === "column" ? event.clientX : event.clientY) -
          active.origin,
        active.axis,
      );
      setPreview({ axis: active.axis, index: active.index, size: active.size });
    }
  };
  const end = (event: ReactPointerEvent) => {
    if (gesture.current?.pointerId !== event.pointerId) return;
    move(event);
    stop(true);
  };
  return {
    preview,
    beginSelection,
    beginResize,
    move,
    end,
    cancel: () => stop(false),
    isActive: () => gesture.current !== null,
  };
}

/** Drag/Shift-click selection, Shift-arrow extension, dimension handles and merged-cell layout. */
export function SheetGrid(props: SheetGridProps): ReactNode {
  const workbook = useWorkbook(),
    state = useWorkbookState();
  const sheet = state.snapshot.sheets.find((s) => s.id === state.activeSheetId);
  return sheet ? (
    <GridSurface key={sheet.id} {...props} sheet={sheet} workbook={workbook} />
  ) : null;
}
function GridSurface({
  sheet,
  workbook,
  className,
  range,
  renderCell,
  onError,
  readOnly: readOnlyProp = false,
}: SheetGridProps & { sheet: SheetSnapshot; workbook: Workbook }): ReactNode {
  const state = useWorkbookState();
  const readOnly = useWorkbookReadOnly() || readOnlyProp;
  const [editing, setEditing] = useState<string | null>(null);
  useEffect(() => {
    if (readOnly) {
      setEditing(null);
      setHoverBoundary(null);
    }
  }, [readOnly]);
  const scrollRef = useRef<HTMLDivElement>(null),
    tableRef = useRef<HTMLTableElement>(null);
  const [hoverBoundary, setHoverBoundary] = useState<ResizeBoundary | null>(
    null,
  );
  const [guideOffset, setGuideOffset] = useState<number | null>(null);
  const [error, setError] = useState<{
    revision: number;
    message: string;
  } | null>(null);
  const report = (error: CommandError) => {
    setError({
      revision: workbook.getState().revision,
      message: error.message,
    });
    onError?.(error);
  };
  const gestures = useGridGestures(
    workbook,
    sheet,
    scrollRef,
    tableRef,
    report,
    readOnly,
  );
  const startRow = Math.max(
      0,
      Math.min(sheet.rowCount - 1, Math.floor(range?.start.row ?? 0)),
    ),
    startColumn = Math.max(
      0,
      Math.min(sheet.columnCount - 1, Math.floor(range?.start.column ?? 0)),
    );
  const read = (address: string) => {
    const value = workbook.getCellValue(sheet.id, address);
    return typeof value === "object" && value !== null && value.kind === "array"
      ? null
      : value;
  };
  const frozenRows = sheet.frozen?.rows ?? 0,
    frozenColumns = sheet.frozen?.columns ?? 0;
  const rows = [
    ...new Set([
      ...getVisibleRows(sheet, { end: frozenRows - 1, limit: 100, read }),
      ...getVisibleRows(sheet, {
        start: startRow,
        end: range?.end.row ?? sheet.rowCount - 1,
        limit: 100,
        read,
      }),
    ]),
  ]
    .sort((a, b) => a - b)
    .slice(0, 100);
  const columns = [
    ...new Set([
      ...getVisibleColumns(sheet, { end: frozenColumns - 1, limit: 26 }),
      ...getVisibleColumns(sheet, {
        start: startColumn,
        end: range?.end.column ?? sheet.columnCount - 1,
        limit: 26,
      }),
    ]),
  ]
    .sort((a, b) => a - b)
    .slice(0, 26);
  const endRow = rows.at(-1) ?? startRow,
    endColumn = columns.at(-1) ?? startColumn;
  const windowRange = {
    start: { row: rows[0] ?? startRow, column: columns[0] ?? startColumn },
    end: { row: endRow, column: endColumn },
  };
  const columnWidth = (index: number) =>
    gestures.preview?.axis === "column" && gestures.preview.index === index
      ? gestures.preview.size
      : getColumnWidth(sheet, index);
  const rowHeight = (index: number) =>
    gestures.preview?.axis === "row" && gestures.preview.index === index
      ? gestures.preview.size
      : getRowHeight(sheet, index);
  const activeBoundary =
    readOnly || editing ? null : (gestures.preview ?? hoverBoundary);
  const updateGuide = () => {
    const scroll = scrollRef.current,
      table = tableRef.current;
    if (!activeBoundary || !scroll || !table) {
      setGuideOffset(null);
      return;
    }
    const element = table.querySelector<HTMLElement>(
      activeBoundary.axis === "column"
        ? `thead [data-column-index="${activeBoundary.index}"]`
        : `tbody tr[data-row-index="${activeBoundary.index}"] > th`,
    );
    if (!element) {
      setGuideOffset(null);
      return;
    }
    const rect = element.getBoundingClientRect(),
      viewport = scroll.getBoundingClientRect();
    const offset =
      activeBoundary.axis === "column"
        ? rect.right - viewport.left
        : rect.bottom - viewport.top;
    const limit =
      activeBoundary.axis === "column"
        ? scroll.clientWidth
        : scroll.clientHeight;
    setGuideOffset(offset >= 0 && offset <= limit ? offset : null);
  };
  useLayoutEffect(updateGuide);
  const hoverAt = (event: ReactPointerEvent) => {
    if (readOnly || editing || gestures.isActive()) return;
    const next = resizeBoundaryAt(event, rows, columns);
    setHoverBoundary((previous) =>
      previous?.axis === next?.axis && previous?.index === next?.index
        ? previous
        : next,
    );
  };
  const rowTop = (row: number) =>
    32 +
    rows
      .filter((index) => index < row && index < frozenRows)
      .reduce((sum, index) => sum + rowHeight(index), 0);
  const columnLeft = (column: number) =>
    42 +
    columns
      .filter((index) => index < column && index < frozenColumns)
      .reduce((sum, index) => sum + columnWidth(index), 0);
  const selection =
    state.selection?.sheetId === sheet.id ? state.selection : null;
  const focus =
    selection &&
    rangeContains(windowRange, { start: selection.focus, end: selection.focus })
      ? selection.focus
      : (mergedRangeAt(sheet, windowRange.start)?.start ?? windowRange.start);
  const focusAddress = addressOf(focus);
  const resizeHandle = (
    axis: "column" | "row",
    index: number,
    size: number,
  ) => {
    if (readOnly) return null;
    const label =
      axis === "column"
        ? addressOf({ row: 0, column: index }).replace(/1$/, "")
        : String(index + 1);
    const commit = (value: number | null) => {
      if (readOnly) return;
      const result =
        axis === "column"
          ? workbook.dispatch({
              type: "resizeColumn",
              sheetId: sheet.id,
              column: index,
              width: value,
            })
          : workbook.dispatch({
              type: "resizeRow",
              sheetId: sheet.id,
              row: index,
              height: value,
            });
      if (!result.ok) report(result.error);
    };
    const style: CSSProperties =
      axis === "column"
        ? {
            position: "absolute",
            right: -3,
            top: 0,
            bottom: 0,
            width: 8,
            cursor: "col-resize",
            zIndex: 3,
            touchAction: "none",
          }
        : {
            position: "absolute",
            left: 0,
            right: 0,
            bottom: -3,
            height: 8,
            cursor: "row-resize",
            zIndex: 3,
            touchAction: "none",
          };
    return (
      <div
        data-resize-axis={axis}
        data-resize-index={index}
        role="separator"
        tabIndex={0}
        aria-label={`Resize ${axis} ${label}`}
        aria-orientation={axis === "column" ? "vertical" : "horizontal"}
        aria-valuemin={1}
        aria-valuemax={10000}
        aria-valuenow={size}
        className={`nori-resize-handle nori-resize-${axis}`}
        style={style}
        onPointerDown={(event) =>
          gestures.beginResize(event, axis, index, size)
        }
        onDoubleClick={() => commit(null)}
        onKeyDown={(event) => {
          const increase = axis === "column" ? "ArrowRight" : "ArrowDown",
            decrease = axis === "column" ? "ArrowLeft" : "ArrowUp";
          if (event.key === increase || event.key === decrease) {
            event.preventDefault();
            event.stopPropagation();
            commit(
              resizedSize(
                size,
                (event.key === increase ? 1 : -1) * (event.shiftKey ? 1 : 10),
                axis,
              ),
            );
          } else if (event.key === "Home") {
            event.preventDefault();
            event.stopPropagation();
            commit(null);
          }
        }}
      />
    );
  };
  return (
    <>
      {sheet.frozen ||
      sheet.autoFilter ||
      sheet.sort ||
      sheet.hiddenRows?.length ||
      sheet.hiddenColumns?.length ? (
        <div className="nori-view-status" aria-label="Worksheet view settings">
          {(frozenRows > 0 || frozenColumns > 0) && (
            <span>
              Frozen: {frozenRows} rows · {frozenColumns} columns
            </span>
          )}
          {sheet.autoFilter && <span>Filtered view</span>}
          {sheet.sort && <span>Saved sort order</span>}
          {!!sheet.hiddenRows?.length && (
            <span>{sheet.hiddenRows.length} hidden rows</span>
          )}
          {!!sheet.hiddenColumns?.length && (
            <span>{sheet.hiddenColumns.length} hidden columns</span>
          )}
        </div>
      ) : null}
      <div style={{ position: "relative", minWidth: 0 }}>
        <div
          ref={scrollRef}
          className="nori-scroll"
          style={{
            position: "relative",
            overflow: "auto",
            maxHeight: 520,
            cursor: activeBoundary
              ? activeBoundary.axis === "column"
                ? "col-resize"
                : "row-resize"
              : undefined,
          }}
          onPointerMoveCapture={hoverAt}
          onPointerLeave={() => setHoverBoundary(null)}
          onScroll={() => {
            setHoverBoundary(null);
            updateGuide();
          }}
          onPointerDownCapture={(event) => {
            if (readOnly || editing) return;
            const boundary = resizeBoundaryAt(event, rows, columns);
            if (!boundary || event.button !== 0 || event.isPrimary === false)
              return;
            setHoverBoundary(boundary);
            gestures.beginResize(
              event,
              boundary.axis,
              boundary.index,
              boundary.axis === "column"
                ? columnWidth(boundary.index)
                : rowHeight(boundary.index),
            );
          }}
          onPointerMove={gestures.move}
          onPointerUp={gestures.end}
          onPointerCancel={gestures.cancel}
          onLostPointerCapture={gestures.cancel}
          onKeyDown={(event) => {
            if (event.key === "Escape" && gestures.isActive()) {
              event.preventDefault();
              gestures.cancel();
              return;
            }
            if (
              !(event.target instanceof Element) ||
              !event.target.closest("[data-cell-address]")
            )
              return;
            const direction =
              event.key === "ArrowUp"
                ? "up"
                : event.key === "ArrowDown"
                  ? "down"
                  : event.key === "ArrowLeft"
                    ? "left"
                    : event.key === "ArrowRight"
                      ? "right"
                      : undefined;
            if (!direction) return;
            event.preventDefault();
            const result = workbook.moveSelection(direction, {
              extend: event.shiftKey,
              bounds: windowRange,
            });
            if (!result.ok) {
              report(result.error);
              return;
            }
            const next = workbook.getState().selection?.focus;
            if (next) {
              const button = tableRef.current?.querySelector<HTMLButtonElement>(
                `[data-cell-address="${addressOf(next)}"]`,
              );
              button?.focus({ preventScroll: true });
              button?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
            }
          }}
        >
          <table
            ref={tableRef}
            className={className ?? "nori-grid"}
            aria-label={`${sheet.name} spreadsheet`}
            style={{
              tableLayout: "fixed",
              borderCollapse: "collapse",
              width:
                42 +
                columns.reduce((sum, column) => sum + columnWidth(column), 0),
              userSelect: "none",
            }}
          >
            <colgroup>
              <col style={{ width: 42 }} />
              {columns.map((column) => (
                <col key={column} style={{ width: columnWidth(column) }} />
              ))}
            </colgroup>
            <thead>
              <tr style={{ height: 32 }}>
                <th
                  scope="col"
                  style={{
                    width: 42,
                    position: "sticky",
                    top: 0,
                    left: 0,
                    zIndex: 10,
                    background: "var(--nori-header-background, #f6f8f1)",
                  }}
                >
                  #
                </th>
                {columns.map((column) => (
                  <th
                    key={column}
                    scope="col"
                    data-column-index={column}
                    data-column-start={column}
                    data-column-end={column}
                    aria-sort={
                      sheet.sort?.conditions.find(
                        (condition) => condition.column === column,
                      )?.direction
                    }
                    style={{
                      position: "sticky",
                      top: 0,
                      left:
                        column < frozenColumns ? columnLeft(column) : undefined,
                      zIndex: column < frozenColumns ? 9 : 6,
                      padding: 0,
                      height: 32,
                      background: "var(--nori-header-background, #f6f8f1)",
                    }}
                  >
                    {addressOf({ row: 0, column }).replace(/1$/, "")}
                    {sheet.autoFilter?.columns.some(
                      (filter) => filter.column === column,
                    ) && (
                      <span title="Filter applied" aria-label="Filter applied">
                        {" "}
                        ⌄
                      </span>
                    )}
                    {sheet.sort?.conditions.find(
                      (condition) => condition.column === column,
                    ) && (
                      <span aria-hidden="true">
                        {" "}
                        {sheet.sort.conditions.find(
                          (condition) => condition.column === column,
                        )?.direction === "descending"
                          ? "↓"
                          : "↑"}
                      </span>
                    )}
                    {resizeHandle("column", column, columnWidth(column))}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr
                  key={row}
                  data-row-index={row}
                  style={{ height: rowHeight(row) }}
                >
                  <th
                    scope="row"
                    data-row-start={row}
                    data-row-end={row}
                    style={{
                      position: "sticky",
                      left: 0,
                      top: row < frozenRows ? rowTop(row) : undefined,
                      zIndex: row < frozenRows ? 8 : 4,
                      padding: 0,
                      height: rowHeight(row),
                      background: "var(--nori-header-background, #f6f8f1)",
                    }}
                  >
                    {row + 1}
                    {resizeHandle("row", row, rowHeight(row))}
                  </th>
                  {columns.map((column) => {
                    const position = { row, column },
                      merge = mergedRangeAt(sheet, position);
                    if (
                      merge &&
                      (row !==
                        rows.find(
                          (index) =>
                            index >= merge.start.row && index <= merge.end.row,
                        ) ||
                        column !==
                          columns.find(
                            (index) =>
                              index >= merge.start.column &&
                              index <= merge.end.column,
                          ))
                    )
                      return null;
                    const anchor = merge?.start ?? position,
                      address = addressOf(anchor),
                      cell = sheet.cells[address],
                      value = workbook.getCellValue(sheet.id, address);
                    const selected =
                      !!selection &&
                      rangesIntersect(
                        selection.range,
                        merge ?? { start: position, end: position },
                      );
                    const rowSpan = merge
                        ? rows.filter(
                            (index) => index >= row && index <= merge.end.row,
                          ).length
                        : 1,
                      colSpan = merge
                        ? columns.filter(
                            (index) =>
                              index >= column && index <= merge.end.column,
                          ).length
                        : 1;
                    const freezeRow = (merge?.end.row ?? row) < frozenRows,
                      freezeColumn =
                        (merge?.end.column ?? column) < frozenColumns;
                    return (
                      <td
                        key={column}
                        rowSpan={rowSpan}
                        colSpan={colSpan}
                        data-address={address}
                        data-column-start={column}
                        data-column-end={
                          columns[columns.indexOf(column) + colSpan - 1]
                        }
                        data-row-start={row}
                        data-row-end={rows[rows.indexOf(row) + rowSpan - 1]}
                        data-selected={selected || undefined}
                        data-merged={!!merge || undefined}
                        style={{
                          ...cellStyleToCss(cell),
                          position:
                            freezeRow || freezeColumn ? "sticky" : "relative",
                          top: freezeRow ? rowTop(row) : undefined,
                          left: freezeColumn ? columnLeft(column) : undefined,
                          zIndex:
                            freezeRow && freezeColumn
                              ? 5
                              : freezeRow || freezeColumn
                                ? 3
                                : undefined,
                          backgroundColor:
                            cell?.style?.background ??
                            (freezeRow || freezeColumn
                              ? "var(--nori-cell-background, white)"
                              : undefined),
                          padding: 0,
                          overflow: "hidden",
                        }}
                      >
                        <button
                          type="button"
                          data-cell-address={address}
                          tabIndex={address === focusAddress ? 0 : -1}
                          aria-label={`${address}: ${formatCellValue(value, cell)}`}
                          aria-pressed={selected}
                          title={formatCellValue(value, cell)}
                          style={{
                            position: "absolute",
                            cursor: "inherit",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            minWidth: 0,
                            minHeight: 0,
                            boxSizing: "border-box",
                            padding: "0 10px",
                            border: 0,
                            borderRadius: 0,
                            background: "transparent",
                            color: "inherit",
                            font: "inherit",
                            textAlign: "inherit",
                            whiteSpace: "inherit",
                            overflowWrap: "inherit",
                            overflow: "hidden",
                            touchAction: "none",
                            display: "flex",
                            flexDirection: "column",
                            justifyContent:
                              cell?.style?.verticalAlign === "top"
                                ? "flex-start"
                                : cell?.style?.verticalAlign === "bottom"
                                  ? "flex-end"
                                  : "center",
                          }}
                          onPointerDown={(event) =>
                            gestures.beginSelection(event, anchor)
                          }
                          onDoubleClick={() => {
                            if (!readOnly) {
                              gestures.cancel();
                              setHoverBoundary(null);
                              setEditing(address);
                            }
                          }}
                          onKeyDown={(event) => {
                            if (
                              !readOnly &&
                              (event.key === "Enter" || event.key === "F2")
                            ) {
                              event.preventDefault();
                              event.stopPropagation();
                              setEditing(address);
                            }
                          }}
                          onClick={(event) => {
                            if (event.detail !== 0) return;
                            const result = workbook.selectCell(
                              sheet.id,
                              anchor,
                              {
                                extend: event.shiftKey,
                              },
                            );
                            if (!result.ok) report(result.error);
                          }}
                        >
                          <span
                            style={{
                              display: "block",
                              width: "100%",
                              overflow: "hidden",
                              textOverflow: cell?.style?.wrapText
                                ? "clip"
                                : "ellipsis",
                            }}
                          >
                            {renderCell
                              ? renderCell({
                                  sheet,
                                  address,
                                  cell,
                                  value,
                                  selected,
                                  mergedRange: merge,
                                })
                              : formatCellValue(value, cell)}
                          </span>
                        </button>
                        {!readOnly && editing === address && (
                          <InlineCellEditor
                            key={address}
                            cell={cell}
                            address={address}
                            cancel={() => {
                              setEditing(null);
                              tableRef.current
                                ?.querySelector<HTMLButtonElement>(
                                  `[data-cell-address="${address}"]`,
                                )
                                ?.focus({ preventScroll: true });
                            }}
                            commit={(draft) => {
                              const result = workbook.dispatch({
                                type: "setCell",
                                sheetId: sheet.id,
                                address,
                                cell: editedCell(draft, cell),
                              });
                              if (!result.ok) {
                                report(result.error);
                                return false;
                              }
                              setEditing(null);
                              return true;
                            }}
                          />
                        )}
                        {selected && (
                          <span
                            aria-hidden="true"
                            className="nori-selection-fill"
                            style={{
                              position: "absolute",
                              inset: 0,
                              pointerEvents: "none",
                              background:
                                "var(--nori-selection-fill, rgba(107,151,71,0.13))",
                              boxShadow:
                                "inset 0 0 0 1px var(--nori-selection-border, #789a55)",
                            }}
                          />
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
          {(!rows.length || !columns.length) && (
            <p role="status">No visible cells in this view.</p>
          )}
        </div>
        {activeBoundary && guideOffset !== null && (
          <div
            aria-hidden="true"
            className="nori-resize-guide"
            data-axis={activeBoundary.axis}
            data-index={activeBoundary.index}
            style={{
              position: "absolute",
              pointerEvents: "none",
              zIndex: 20,
              background: "var(--nori-resize-border, #4b7d34)",
              ...(activeBoundary.axis === "column"
                ? { top: 0, bottom: 0, left: guideOffset - 1, width: 2 }
                : { left: 0, right: 0, top: guideOffset - 1, height: 2 }),
            }}
          />
        )}
      </div>
      {error?.revision === state.revision && (
        <div role="alert">{error.message}</div>
      )}
    </>
  );
}
