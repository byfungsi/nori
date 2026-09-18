import type { CellRange, Position, SheetSnapshot } from "./workbook-types";

/** Normalize inclusive corners without changing the caller's input. */
export function normalizeRange(range: CellRange): CellRange {
  return {
    start: {
      row: Math.min(range.start.row, range.end.row),
      column: Math.min(range.start.column, range.end.column),
    },
    end: {
      row: Math.max(range.start.row, range.end.row),
      column: Math.max(range.start.column, range.end.column),
    },
  };
}
/** Test overlap between inclusive rectangles, accepting reversed corners. */
export function rangesIntersect(a: CellRange, b: CellRange): boolean {
  const x = normalizeRange(a),
    y = normalizeRange(b);
  return (
    x.start.row <= y.end.row &&
    x.end.row >= y.start.row &&
    x.start.column <= y.end.column &&
    x.end.column >= y.start.column
  );
}
/** Test full containment between inclusive rectangles. */
export function rangeContains(outer: CellRange, inner: CellRange): boolean {
  const a = normalizeRange(outer),
    b = normalizeRange(inner);
  return (
    a.start.row <= b.start.row &&
    a.end.row >= b.end.row &&
    a.start.column <= b.start.column &&
    a.end.column >= b.end.column
  );
}
/** Find the merged region containing a cell; undefined means an ordinary cell. */
export function mergedRangeAt(
  sheet: SheetSnapshot,
  position: Position,
): CellRange | undefined {
  return sheet.merges?.find((range) =>
    rangeContains(range, { start: position, end: position }),
  );
}
/** Expand to include every intersected merge, including merges reached by expansion. */
export function expandRangeToMerges(
  sheet: SheetSnapshot,
  range: CellRange,
): CellRange {
  let result = normalizeRange(range),
    changed = true;
  while (changed) {
    changed = false;
    for (const merge of sheet.merges ?? []) {
      if (!rangesIntersect(result, merge) || rangeContains(result, merge))
        continue;
      result = {
        start: {
          row: Math.min(result.start.row, merge.start.row),
          column: Math.min(result.start.column, merge.start.column),
        },
        end: {
          row: Math.max(result.end.row, merge.end.row),
          column: Math.max(result.end.column, merge.end.column),
        },
      };
      changed = true;
    }
  }
  return result;
}
