import {
  useState,
  useSyncExternalStore,
  type CSSProperties,
  type ReactNode,
} from "react";
import { addressOf, mergedRangeAt, type SheetId } from "@nori-internal/model";
import {
  getColumnWidth,
  getVisibleColumns,
  getVisibleRows,
  type Workbook,
} from "@nori-internal/core";
import { cellStyleToCss, formatCellValue } from "./cell-format";
/** A compact read-only workbook card for agentic chat or attachment thumbnails. */
export interface SpreadsheetPreviewProps {
  readonly workbook: Workbook;
  readonly title?: string;
  readonly theme?: "light" | "dark";
  /** Limits bound thumbnail rendering, not the source workbook. */
  readonly maxRows?: number;
  readonly maxColumns?: number;
  readonly className?: string;
  readonly style?: CSSProperties;
  /** Optional controlled local tab. Preview navigation never mutates the runtime. */
  readonly sheetId?: SheetId;
  readonly onSheetChange?: (sheetId: SheetId) => void;
  readonly onOpen?: () => void;
}
const bounded = (value: number | undefined, fallback: number, max: number) =>
  Number.isFinite(value)
    ? Math.max(1, Math.min(max, Math.floor(value ?? fallback)))
    : fallback;
/** Responsive file card with local sheet tabs, calculated values and optional full-editor action. */
export function SpreadsheetPreview({
  workbook,
  title = "Workbook.xlsx",
  theme = "light",
  maxRows,
  maxColumns,
  className,
  style,
  sheetId,
  onSheetChange,
  onOpen,
}: SpreadsheetPreviewProps): ReactNode {
  const state = useSyncExternalStore(
    workbook.subscribe,
    workbook.getState,
    workbook.getState,
  );
  const [localSheetId, setLocalSheetId] = useState(() => state.activeSheetId);
  const sheet =
    state.snapshot.sheets.find(
      (sheet) => sheet.id === (sheetId ?? localSheetId),
    ) ?? state.snapshot.sheets[0];
  if (!sheet) return null;
  const rowLimit = bounded(maxRows, 8, 20),
    columnLimit = bounded(maxColumns, 6, 12);
  const read = (address: string) => {
    const value = workbook.getCellValue(sheet.id, address);
    return typeof value === "object" && value !== null && value.kind === "array"
      ? null
      : value;
  };
  const rows = getVisibleRows(sheet, { limit: rowLimit, read }),
    columns = getVisibleColumns(sheet, { limit: columnLimit });
  const dark = theme === "dark",
    border = dark ? "#37423e" : "#dfe5dc",
    background = dark ? "#202824" : "#ffffff",
    muted = dark ? "#a6b2ab" : "#68756b",
    foreground = dark ? "#e5ece7" : "#26372d",
    header = dark ? "#18211c" : "#f3f6ee";
  const buttonStyle: CSSProperties = {
    font: "inherit",
    fontSize: 12,
    border: 0,
    borderRadius: 5,
    padding: "6px 10px",
    cursor: "pointer",
    color: foreground,
    background: "transparent",
    whiteSpace: "nowrap",
    flexShrink: 0,
  };
  return (
    <figure
      className={className ?? "nori-preview"}
      data-theme={theme}
      style={{
        margin: 0,
        width: "100%",
        maxWidth: "100%",
        minWidth: 0,
        boxSizing: "border-box",
        overflow: "hidden",
        border: `1px solid ${border}`,
        borderRadius: 12,
        background,
        color: foreground,
        fontFamily: "inherit",
        fontSize: 13,
        ...style,
      }}
    >
      <figcaption
        style={{
          display: "flex",
          alignItems: "center",
          gap: 10,
          padding: "12px 14px",
          borderBottom: `1px solid ${border}`,
          minWidth: 0,
          flexWrap: "wrap",
        }}
      >
        <span
          aria-hidden="true"
          style={{ color: dark ? "#6cd3a3" : "#2c8057", fontSize: 19 }}
        >
          ▦
        </span>
        <strong
          title={title}
          style={{
            flex: "1 1 110px",
            overflow: "hidden",
            whiteSpace: "nowrap",
            textOverflow: "ellipsis",
            minWidth: 0,
          }}
        >
          {title}
        </strong>
        {onOpen && (
          <button
            type="button"
            style={{ ...buttonStyle, border: `1px solid ${border}` }}
            onClick={onOpen}
          >
            Open workbook
          </button>
        )}
      </figcaption>
      <div
        tabIndex={0}
        role="region"
        aria-label={`${title} sheet preview, scroll to see more columns`}
        style={{
          overflowX: "auto",
          overscrollBehaviorX: "contain",
          maxWidth: "100%",
        }}
      >
        <table
          aria-label={`${sheet.name} preview`}
          style={{
            tableLayout: "fixed",
            borderCollapse: "collapse",
            width:
              36 +
              columns.reduce(
                (sum, column) =>
                  sum +
                  Math.max(90, Math.min(180, getColumnWidth(sheet, column))),
                0,
              ),
            fontSize: 12,
          }}
        >
          <colgroup>
            <col style={{ width: 36 }} />
            {columns.map((column) => (
              <col
                key={column}
                style={{
                  width: Math.max(
                    90,
                    Math.min(180, getColumnWidth(sheet, column)),
                  ),
                }}
              />
            ))}
          </colgroup>
          <thead style={{ background: header, color: muted }}>
            <tr style={{ height: 25 }}>
              <th scope="col">#</th>
              {columns.map((column) => (
                <th
                  scope="col"
                  key={column}
                  style={{ fontWeight: 500, borderLeft: `1px solid ${border}` }}
                >
                  {addressOf({ row: 0, column }).replace(/1$/, "")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row} style={{ height: 31 }}>
                <th
                  scope="row"
                  style={{
                    fontWeight: 400,
                    color: muted,
                    background: header,
                    fontSize: 11,
                    borderTop: `1px solid ${border}`,
                  }}
                >
                  {row + 1}
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
                  const address = addressOf(merge?.start ?? position),
                    cell = sheet.cells[address],
                    value = workbook.getCellValue(sheet.id, address),
                    text = formatCellValue(value, cell);
                  return (
                    <td
                      key={column}
                      rowSpan={
                        merge
                          ? rows.filter(
                              (index) => index >= row && index <= merge.end.row,
                            ).length
                          : 1
                      }
                      colSpan={
                        merge
                          ? columns.filter(
                              (index) =>
                                index >= column && index <= merge.end.column,
                            ).length
                          : 1
                      }
                      title={text}
                      style={{
                        ...cellStyleToCss(cell),
                        position: "relative",
                        padding: 0,
                        borderLeft: `1px solid ${border}`,
                        borderTop: `1px solid ${border}`,
                      }}
                    >
                      <span
                        style={{
                          position: "absolute",
                          inset: 0,
                          display: "flex",
                          alignItems: "center",
                          padding: "0 10px",
                          minWidth: 0,
                        }}
                      >
                        <span
                          style={{
                            display: "block",
                            width: "100%",
                            overflow: "hidden",
                            whiteSpace: "nowrap",
                            textOverflow: "ellipsis",
                          }}
                        >
                          {text}
                        </span>
                      </span>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
        {(!rows.length || !columns.length) && (
          <p style={{ padding: 12, color: muted }}>No visible cells.</p>
        )}
      </div>
      <nav
        aria-label={`${title} preview sheets`}
        style={{
          display: "flex",
          gap: 4,
          padding: "7px 9px",
          overflowX: "auto",
          borderTop: `1px solid ${border}`,
          background: header,
        }}
      >
        {state.snapshot.sheets.map((item) => (
          <button
            type="button"
            key={item.id}
            aria-pressed={item.id === sheet.id}
            style={{
              ...buttonStyle,
              background:
                item.id === sheet.id
                  ? dark
                    ? "#35453b"
                    : "#dfead4"
                  : "transparent",
              fontWeight: item.id === sheet.id ? 600 : 400,
            }}
            onClick={() => {
              setLocalSheetId(item.id);
              onSheetChange?.(item.id);
            }}
          >
            {item.name}
          </button>
        ))}
      </nav>
      <div
        style={{
          padding: "6px 14px",
          fontSize: 11,
          color: muted,
          borderTop: `1px solid ${border}`,
        }}
      >
        {rows.length} rows · {columns.length} columns shown
        {sheet.autoFilter ? " · Filtered" : ""}
        {sheet.frozen?.rows || sheet.frozen?.columns
          ? " · Freeze panes saved"
          : ""}
      </div>
    </figure>
  );
}
