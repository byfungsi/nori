import type { CSSProperties } from "react";
import { formatNumber, isCellError, type Cell } from "@nori-internal/model";
import { isArrayResult, type FormulaResult } from "@nori-internal/formula";
/** Workbook format styling, intentionally separate from the surrounding application theme. */
export function cellStyleToCss(cell: Cell | undefined): CSSProperties {
  return {
    fontWeight: cell?.style?.fontWeight,
    fontStyle: cell?.style?.fontStyle,
    color: cell?.style?.color,
    backgroundColor: cell?.style?.background,
    textAlign: cell?.style?.horizontalAlign,
    verticalAlign: cell?.style?.verticalAlign,
    whiteSpace: cell?.style?.wrapText ? "pre-wrap" : "pre",
    overflowWrap: cell?.style?.wrapText ? "anywhere" : undefined,
  };
}
/** Default display supports raw scalars, errors and a small explicit numeric format subset. */
export function formatCellValue(value: FormulaResult, cell?: Cell): string {
  if (isArrayResult(value)) return "[array]";
  if (isCellError(value)) return value.code;
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number" && cell?.style?.numberFormat) {
    return formatNumber(value, cell.style.numberFormat) ?? String(value);
  }
  return String(value);
}
