import {
  err,
  ok,
  type CellRange,
  type Result,
  type SheetSnapshot,
} from "@nori-internal/model";
import { CommandError } from "./command-error";
/** Dimensions use host-defined logical units; sparse overrides use zero-based indexes. */
export interface GridLayout {
  readonly rowHeight: number;
  readonly columnWidth: number;
  readonly rowHeights?: Readonly<Record<string, number>> | undefined;
  readonly columnWidths?: Readonly<Record<string, number>> | undefined;
}
/** Defaults shared by layout calculations and renderer adapters. */
export const defaultGridLayout: GridLayout = Object.freeze({
  rowHeight: 34,
  columnWidth: 144,
});
/** Resolve a column's document width with the default for unspecified columns. */
export function getColumnWidth(sheet: SheetSnapshot, column: number): number {
  return (
    sheet.columnWidths?.[String(column)] ??
    sheet.defaultColumnWidth ??
    defaultGridLayout.columnWidth
  );
}
/** Resolve a row's document height with the default for unspecified rows. */
export function getRowHeight(sheet: SheetSnapshot, row: number): number {
  return (
    sheet.rowHeights?.[String(row)] ??
    sheet.defaultRowHeight ??
    defaultGridLayout.rowHeight
  );
}
function axisIndex(
  offset: number,
  count: number,
  base: number,
  overrides: Readonly<Record<string, number>> | undefined,
  inclusive: boolean,
): number {
  const entries = Object.entries(overrides ?? {}).map(([index, size]) => ({
    index: Number(index),
    size,
  }));
  const start = (index: number) =>
    index * base +
    entries.reduce(
      (sum, item) => sum + (item.index < index ? item.size - base : 0),
      0,
    );
  let low = 0,
    high = count;
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    const value = start(mid);
    if (inclusive ? value <= offset : value < offset) low = mid + 1;
    else high = mid;
  }
  return Math.max(0, Math.min(count - 1, low - 1));
}
/** Map a viewport to an inclusive range using uniform or sparsely overridden dimensions. */
export function getVisibleRange(
  layout: GridLayout,
  viewport: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  },
  extent: { readonly rows: number; readonly columns: number },
): Result<CellRange, CommandError> {
  if (
    ![layout.rowHeight, layout.columnWidth, extent.rows, extent.columns].every(
      (v) => Number.isFinite(v) && v > 0,
    ) ||
    ![extent.rows, extent.columns].every(Number.isInteger) ||
    ![viewport.x, viewport.y, viewport.width, viewport.height].every(
      (v) => Number.isFinite(v) && v >= 0,
    )
  )
    return err(new CommandError("Invalid layout or viewport"));
  for (const [overrides, count] of [
    [layout.rowHeights, extent.rows],
    [layout.columnWidths, extent.columns],
  ] as const)
    for (const [index, size] of Object.entries(overrides ?? {})) {
      if (
        !/^(0|[1-9][0-9]*)$/.test(index) ||
        Number(index) >= count ||
        !Number.isFinite(size) ||
        size <= 0
      )
        return err(new CommandError("Invalid dimension override"));
    }
  const row = axisIndex(
      viewport.y,
      extent.rows,
      layout.rowHeight,
      layout.rowHeights,
      true,
    ),
    column = axisIndex(
      viewport.x,
      extent.columns,
      layout.columnWidth,
      layout.columnWidths,
      true,
    );
  return ok({
    start: { row, column },
    end: {
      row: Math.max(
        row,
        axisIndex(
          viewport.y + viewport.height,
          extent.rows,
          layout.rowHeight,
          layout.rowHeights,
          false,
        ),
      ),
      column: Math.max(
        column,
        axisIndex(
          viewport.x + viewport.width,
          extent.columns,
          layout.columnWidth,
          layout.columnWidths,
          false,
        ),
      ),
    },
  });
}
