export { SpreadsheetPreview } from "./spreadsheet-preview";
export type { SpreadsheetPreviewProps } from "./spreadsheet-preview";
import {
  WorkbookProvider,
  useWorkbook,
  useWorkbookState,
  useWorkbookReadOnly,
} from "./workbook-context";
import { cellDraft, editedCell } from "./cell-edit";
import { SheetGrid, type SheetGridProps } from "./sheet-grid";
import { SelectionToolbar } from "./selection-toolbar";
export {
  WorkbookProvider,
  useWorkbook,
  useWorkbookState,
  useWorkbookReadOnly,
} from "./workbook-context";
export { formatCellValue, cellStyleToCss } from "./cell-format";
export { SheetGrid } from "./sheet-grid";
export type { SheetGridProps, CellRenderContext } from "./sheet-grid";
export { SelectionToolbar } from "./selection-toolbar";
import { useId, useState, type CSSProperties, type ReactNode } from "react";
import { addressOf, type Cell, type CellRange } from "@nori-internal/model";

import type {
  CommandError,
  Workbook,
  WorkbookState,
} from "@nori-internal/core";
/** Sheet switching with native buttons and an accessible current-sheet indicator. */
export function WorkbookTabs({
  className,
}: {
  readonly className?: string;
}): ReactNode {
  const workbook = useWorkbook(),
    state = useWorkbookState();
  return (
    <nav className={className ?? "nori-tabs"} aria-label="Workbook sheets">
      {state.snapshot.sheets.map((sheet) => (
        <button
          type="button"
          key={sheet.id}
          aria-pressed={state.activeSheetId === sheet.id}
          onClick={() => workbook.setActiveSheet(sheet.id)}
        >
          {sheet.name}
        </button>
      ))}
    </nav>
  );
}
/** Edit the selected cell; commands keep formulas, history and recalculation in core. */
export function FormulaBar({
  className,
  onError,
  readOnly: readOnlyProp = false,
}: {
  readonly className?: string;
  readonly onError?: (error: CommandError) => void;
  readonly readOnly?: boolean;
}): ReactNode {
  const workbook = useWorkbook(),
    state = useWorkbookState();
  const readOnly = useWorkbookReadOnly() || readOnlyProp;
  const selection = state.selection;
  if (!selection)
    return (
      <div className={className ?? "nori-formula-bar"}>
        {readOnly
          ? "Select a cell to view its value or formula."
          : "Select a cell to edit its value or formula."}
      </div>
    );
  const address = addressOf(selection.anchor),
    sheet = state.snapshot.sheets.find((s) => s.id === selection.sheetId),
    cell = sheet?.cells[address];
  return (
    <CellEditor
      key={`${selection.sheetId}/${address}/${state.revision}`}
      workbook={workbook}
      readOnly={readOnly}
      state={state}
      address={address}
      cell={cell}
      className={className}
      onError={onError}
    />
  );
}
function CellEditor({
  workbook,
  readOnly,
  state,
  address,
  cell,
  className,
  onError,
}: {
  workbook: Workbook;
  readOnly: boolean;
  state: WorkbookState;
  address: string;
  cell: Cell | undefined;
  className: string | undefined;
  onError: ((error: CommandError) => void) | undefined;
}): ReactNode {
  const [draft, setDraft] = useState(cellDraft(cell));
  const [error, setError] = useState("");
  const id = useId();
  return (
    <form
      className={className ?? "nori-formula-bar"}
      onSubmit={(event) => {
        event.preventDefault();
        if (readOnly || !state.selection) return;
        const result = workbook.dispatch({
          type: "setCell",
          sheetId: state.selection.sheetId,
          address,
          cell: editedCell(draft, cell),
        });
        if (!result.ok) {
          setError(result.error.message);
          onError?.(result.error);
        }
      }}
    >
      <label htmlFor={id}>{address}</label>
      <input
        id={id}
        aria-label="Cell value or formula"
        readOnly={readOnly}
        value={draft}
        onChange={(event) => setDraft(event.target.value)}
      />
      <button type="submit" disabled={readOnly}>
        Apply
      </button>
      {error && <span role="alert">{error}</span>}
    </form>
  );
}
/** Optional ready-to-compose view; primitives can instead be arranged independently. */
export function Spreadsheet({
  workbook,
  className,
  style,
  range,
  renderCell,
  readOnly = false,
}: {
  readonly workbook: Workbook;
  readonly readOnly?: boolean;
  readonly className?: string;
  readonly style?: CSSProperties;
  readonly range?: CellRange;
  readonly renderCell?: SheetGridProps["renderCell"];
}): ReactNode {
  return (
    <WorkbookProvider workbook={workbook} readOnly={readOnly}>
      <div className={className ?? "nori-workbook"} style={style}>
        <FormulaBar />
        <SelectionToolbar />
        <SheetGrid
          {...(range ? { range } : {})}
          {...(renderCell ? { renderCell } : {})}
        />
        <WorkbookTabs />
      </div>
    </WorkbookProvider>
  );
}
