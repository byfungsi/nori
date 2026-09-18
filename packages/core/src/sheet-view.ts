import {
  addressOf,
  isCellError,
  type FilterComparison,
  type Scalar,
  type SheetSnapshot,
} from "@nori-internal/model";
/** Visibility checks can use calculated values supplied by the runtime or cached snapshot values. */
export type CellValueReader = (address: string) => Scalar;
/** Hidden columns remain in the model but are omitted by platform renderers. */
export function isColumnVisible(sheet: SheetSnapshot, column: number): boolean {
  return !(sheet.hiddenColumns?.includes(column) ?? false);
}
function text(value: Scalar): string {
  return isCellError(value) ? value.code : value === null ? "" : String(value);
}
function regexLiteral(value: string): string {
  return ".*+?^${}()|[]".includes(value) || value.charCodeAt(0) === 92
    ? String.fromCharCode(92) + value
    : value;
}
function matchesWildcard(value: string, pattern: string): boolean {
  let expression = "^";
  for (let i = 0; i < pattern.length; i++) {
    const char = pattern[i];
    if (char === undefined) continue;
    if (char === "~" && i + 1 < pattern.length)
      expression += regexLiteral(pattern[++i] ?? "");
    else if (char === "*") expression += ".*";
    else if (char === "?") expression += ".";
    else expression += regexLiteral(char);
  }
  return new RegExp(expression + "$", "is").test(value);
}
function compare(value: Scalar, condition: FilterComparison): boolean {
  const a = text(value),
    b = String(condition.value);
  if (condition.operator === "equal" || condition.operator === "notEqual") {
    const equal = matchesWildcard(a, b);
    return condition.operator === "equal" ? equal : !equal;
  }
  const numeric =
    a.trim() !== "" &&
    b.trim() !== "" &&
    Number.isFinite(Number(a)) &&
    Number.isFinite(Number(b));
  const left = numeric ? Number(a) : a.toLowerCase(),
    right = numeric ? Number(b) : b.toLowerCase();
  switch (condition.operator) {
    case "lessThan":
      return left < right;
    case "lessThanOrEqual":
      return left <= right;
    case "greaterThan":
      return left > right;
    case "greaterThanOrEqual":
      return left >= right;
  }
}
/** Combine saved hidden rows with supported filter criteria; the filter header remains visible. */
export function isRowVisible(
  sheet: SheetSnapshot,
  row: number,
  read: CellValueReader = (address) => sheet.cells[address]?.value ?? null,
): boolean {
  if (sheet.hiddenRows?.includes(row)) return false;
  const filter = sheet.autoFilter;
  if (!filter || row <= filter.range.start.row || row > filter.range.end.row)
    return true;
  return filter.columns.every(({ column, condition }) => {
    const value = read(addressOf({ row, column }));
    if (condition.kind === "values")
      return (
        (condition.includeBlank && (value === null || value === "")) ||
        condition.values.some(
          (item) => item.toLowerCase() === text(value).toLowerCase(),
        )
      );
    return condition.join === "and"
      ? condition.comparisons.every((c) => compare(value, c))
      : condition.comparisons.some((c) => compare(value, c));
  });
}
/** Enumerate a bounded visible slice while retaining original worksheet row identities. */
export function getVisibleRows(
  sheet: SheetSnapshot,
  options: {
    readonly start?: number;
    readonly end?: number;
    readonly limit?: number;
    readonly read?: CellValueReader;
  } = {},
): readonly number[] {
  const result: number[] = [],
    limit = Math.max(0, Math.floor(options.limit ?? 100));
  for (
    let row = Math.max(0, Math.floor(options.start ?? 0));
    row <= Math.min(sheet.rowCount - 1, options.end ?? sheet.rowCount - 1) &&
    result.length < limit;
    row++
  )
    if (isRowVisible(sheet, row, options.read)) result.push(row);
  return result;
}
/** Enumerate visible columns without renumbering worksheet addresses. */
export function getVisibleColumns(
  sheet: SheetSnapshot,
  options: {
    readonly start?: number;
    readonly end?: number;
    readonly limit?: number;
  } = {},
): readonly number[] {
  const result: number[] = [],
    limit = Math.max(0, Math.floor(options.limit ?? 26));
  for (
    let column = Math.max(0, Math.floor(options.start ?? 0));
    column <=
      Math.min(sheet.columnCount - 1, options.end ?? sheet.columnCount - 1) &&
    result.length < limit;
    column++
  )
    if (isColumnVisible(sheet, column)) result.push(column);
  return result;
}
