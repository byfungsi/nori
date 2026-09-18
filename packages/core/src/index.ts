import { isColumnVisible, isRowVisible } from "./sheet-view";
export {
  isColumnVisible,
  isRowVisible,
  getVisibleRows,
  getVisibleColumns,
} from "./sheet-view";
export type { CellValueReader } from "./sheet-view";
import {
  addressOf,
  mergedRangeAt,
  normalizeRange,
  rangesIntersect,
  rangeContains,
  expandRangeToMerges,
  parseCellRange,
  cellError,
  err,
  ok,
  parseAddress,
  parseCell,
  parseWorkbookSnapshot,
  type Cell,
  type CellRange,
  type ModelError,
  type Position,
  type Result,
  type SheetId,
  type WorkbookSnapshot,
} from "@nori-internal/model";
import {
  collectReferences,
  createFunctionRegistry,
  DependencyGraph,
  evaluateFormula,
  parseFormula,
  type FormulaAst,
  type FormulaResult,
  type FunctionRegistry,
  type ReferenceAst,
} from "@nori-internal/formula";
export {
  getVisibleRange,
  getColumnWidth,
  getRowHeight,
  defaultGridLayout,
} from "./grid-layout";
export type { GridLayout } from "./grid-layout";
import { CommandError } from "./command-error";
export { CommandError } from "./command-error";
/** Atomic mutations; batches validate completely before committing. */
export type WorkbookCommand =
  | {
      readonly type: "setCell";
      readonly sheetId: SheetId;
      readonly address: string;
      readonly cell: Cell;
    }
  | {
      readonly type: "clearCell";
      readonly sheetId: SheetId;
      readonly address: string;
    }
  | {
      readonly type: "renameSheet";
      readonly sheetId: SheetId;
      readonly name: string;
    }
  | {
      readonly type: "resizeColumn";
      readonly sheetId: SheetId;
      readonly column: number;
      readonly width: number | null;
    }
  | {
      readonly type: "resizeRow";
      readonly sheetId: SheetId;
      readonly row: number;
      readonly height: number | null;
    }
  | {
      readonly type: "mergeCells";
      readonly sheetId: SheetId;
      readonly range: CellRange;
    }
  | {
      readonly type: "unmergeCells";
      readonly sheetId: SheetId;
      readonly range: CellRange;
    }
  | {
      readonly type: "batch";
      readonly commands: readonly Exclude<WorkbookCommand, { type: "batch" }>[];
    };
/** Platform-neutral selection uses sheet identity and inclusive coordinates. */
export interface Selection {
  readonly sheetId: SheetId;
  readonly range: CellRange;
}
/** Normalized selection plus stable gesture anchor and moving focus, including merged cells. */
export interface SelectionState extends Selection {
  readonly anchor: Position;
  readonly focus: Position;
}
/** Selection extension is shared by mouse, keyboard and future native adapters. */
export interface SelectionOptions {
  readonly extend?: boolean;
}
/** Directional movement can be bounded by a renderer's visible window. */
export interface SelectionMoveOptions extends SelectionOptions {
  readonly bounds?: CellRange;
}
/** Stable external-store snapshot, replaced once per successful state transition. */
export interface WorkbookState {
  readonly snapshot: WorkbookSnapshot;
  readonly activeSheetId: SheetId;
  readonly selection: SelectionState | null;
  readonly revision: number;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
}
/** Headless workbook contract used by platform renderers. */
export interface Workbook {
  readonly getState: () => WorkbookState;
  readonly subscribe: (listener: () => void) => () => void;
  readonly dispatch: (command: WorkbookCommand) => Result<void, CommandError>;
  readonly undo: () => boolean;
  readonly redo: () => boolean;
  readonly getCellValue: (sheetId: SheetId, address: string) => FormulaResult;
  readonly setActiveSheet: (sheetId: SheetId) => Result<void, CommandError>;
  readonly setSelection: (
    selection: Selection | null,
  ) => Result<void, CommandError>;
  readonly selectCell: (
    sheetId: SheetId,
    position: Position,
    options?: SelectionOptions,
  ) => Result<void, CommandError>;
  readonly moveSelection: (
    direction: "up" | "down" | "left" | "right",
    options?: SelectionMoveOptions,
  ) => Result<void, CommandError>;
  readonly exportSnapshot: () => WorkbookSnapshot;
}
/** Runtime resource bounds and pure function extension points. */
export interface WorkbookOptions {
  readonly functions?: FunctionRegistry;
  readonly historyLimit?: number;
}
const keyOf = (sheetId: SheetId, address: string) =>
  JSON.stringify([sheetId, address]);
const rangeLimit = 100_000;
function rangeAddresses(start: Position, end: Position): string[] | undefined {
  const top = Math.min(start.row, end.row),
    bottom = Math.max(start.row, end.row),
    left = Math.min(start.column, end.column),
    right = Math.max(start.column, end.column);
  if ((bottom - top + 1) * (right - left + 1) > rangeLimit) return undefined;
  const addresses: string[] = [];
  for (let row = top; row <= bottom; row++)
    for (let column = left; column <= right; column++)
      addresses.push(addressOf({ row, column }));
  return addresses;
}
/** Rehydrate a defensive immutable snapshot and initialize calculation/state ownership. */
export function createWorkbook(
  input: unknown,
  options: WorkbookOptions = {},
): Result<Workbook, ModelError | CommandError> {
  const parsed = parseWorkbookSnapshot(input);
  if (!parsed.ok) return parsed;
  const first = parsed.value.sheets[0];
  if (!first) throw new Error("Snapshot invariant: at least one sheet");
  const historyLimit = options.historyLimit ?? 100;
  if (!Number.isSafeInteger(historyLimit) || historyLimit < 0)
    return err(new CommandError("historyLimit must be a nonnegative integer"));
  let state: WorkbookState = Object.freeze({
    snapshot: parsed.value,
    activeSheetId: first.id,
    selection: null,
    revision: 0,
    canUndo: false,
    canRedo: false,
  });
  const past: WorkbookSnapshot[] = [];
  const future: WorkbookSnapshot[] = [];
  const listeners = new Set<() => void>();
  const registry = new Map(options.functions ?? createFunctionRegistry());
  let asts = new Map<string, FormulaAst>();
  let values = new Map<string, FormulaResult>();
  let graph = new DependencyGraph();
  const resolveSheet = (name: string | undefined, current: SheetId) =>
    name === undefined
      ? state.snapshot.sheets.find((s) => s.id === current)
      : state.snapshot.sheets.find(
          (s) => s.name.toLowerCase() === name.toLowerCase(),
        );
  const rebuild = () => {
    asts = new Map();
    values = new Map();
    graph = new DependencyGraph();
    for (const sheet of state.snapshot.sheets)
      for (const [address, cell] of Object.entries(sheet.cells)) {
        if (!cell.formula) continue;
        const key = keyOf(sheet.id, address);
        const parsedFormula = parseFormula(cell.formula);
        if (!parsedFormula.ok) {
          values.set(key, cellError("#ERROR!"));
          continue;
        }
        asts.set(key, parsedFormula.value);
        const dependencies: string[] = [];
        for (const ref of collectReferences(parsedFormula.value)) {
          const start = ref.kind === "range" ? ref.start : ref;
          const target = resolveSheet(start.sheet, sheet.id);
          if (!target) continue;
          const addresses =
            ref.kind === "range"
              ? rangeAddresses(ref.start.position, ref.end.position)
              : [addressOf(ref.position)];
          if (addresses) {
            for (const address of addresses)
              dependencies.push(keyOf(target.id, address));
          }
        }
        graph.set(key, dependencies);
      }
  };
  const read = (
    sheetId: SheetId,
    address: string,
    visiting: Set<string>,
  ): FormulaResult => {
    const sheet = state.snapshot.sheets.find((s) => s.id === sheetId);
    if (!sheet) return cellError("#REF!");
    const key = keyOf(sheetId, address);
    if (values.has(key)) return values.get(key) ?? null;
    if (visiting.has(key)) return cellError("#CYCLE!");
    if (visiting.size >= 256) return cellError("#NUM!");
    const cell = sheet.cells[address];
    const ast = asts.get(key);
    if (!ast) return cell?.value ?? null;
    visiting.add(key);
    const reference = (ref: ReferenceAst): FormulaResult => {
      const target = resolveSheet(ref.sheet, sheetId);
      return target
        ? read(target.id, addressOf(ref.position), visiting)
        : cellError("#REF!");
    };
    const result = evaluateFormula(
      ast,
      {
        cell: reference,
        range: (start, end) => {
          const target = resolveSheet(start.sheet, sheetId);
          if (!target) return cellError("#REF!");
          const addresses = rangeAddresses(start.position, end.position);
          if (!addresses) return cellError("#NUM!");
          const width =
            Math.abs(end.position.column - start.position.column) + 1;
          const rows: import("@nori-internal/model").Scalar[][] = [];
          for (let i = 0; i < addresses.length; i += width)
            rows.push(
              addresses.slice(i, i + width).map((a) => {
                const v = read(target.id, a, visiting);
                return typeof v === "object" && v !== null && v.kind === "array"
                  ? cellError("#VALUE!")
                  : v;
              }),
            );
          return { kind: "array", values: rows };
        },
      },
      registry,
    );
    visiting.delete(key);
    // Results are frozen so cached arrays cannot be changed through a renderer.
    if (
      typeof result === "object" &&
      result !== null &&
      result.kind === "array"
    ) {
      for (const row of result.values) Object.freeze(row);
      Object.freeze(result.values);
      Object.freeze(result);
    }
    values.set(key, result);
    return result;
  };
  const notify = () => {
    for (const listener of [...listeners]) listener();
  };
  const selectionFor = (
    snapshot: WorkbookSnapshot,
    selection: Selection,
  ): Result<SelectionState, CommandError> => {
    const sheet = snapshot.sheets.find((s) => s.id === selection.sheetId);
    if (!sheet) return err(new CommandError("Unknown sheet"));
    const checked = parseCellRange(selection.range);
    if (
      !checked.ok ||
      checked.value.end.row >= sheet.rowCount ||
      checked.value.end.column >= sheet.columnCount
    )
      return err(new CommandError("Selection outside sheet bounds"));
    const anchor =
      mergedRangeAt(sheet, selection.range.start)?.start ??
      selection.range.start;
    const focus =
      mergedRangeAt(sheet, selection.range.end)?.start ?? selection.range.end;
    const range = expandRangeToMerges(sheet, { start: anchor, end: focus });
    return ok(
      Object.freeze({
        sheetId: sheet.id,
        anchor: Object.freeze({ ...anchor }),
        focus: Object.freeze({ ...focus }),
        range: Object.freeze({
          start: Object.freeze(range.start),
          end: Object.freeze(range.end),
        }),
      }),
    );
  };
  const setSelection = (
    selection: Selection | null,
  ): Result<void, CommandError> => {
    let next: SelectionState | null = null;
    if (selection) {
      const parsed = selectionFor(state.snapshot, selection);
      if (!parsed.ok) return parsed;
      next = parsed.value;
    }
    if (JSON.stringify(next) === JSON.stringify(state.selection))
      return ok(undefined);
    state = Object.freeze({ ...state, selection: next });
    commit(state.snapshot, false);
    return ok(undefined);
  };
  const selectCell = (
    sheetId: SheetId,
    position: Position,
    options: SelectionOptions = {},
  ): Result<void, CommandError> => {
    const anchor =
      options.extend && state.selection?.sheetId === sheetId
        ? state.selection.anchor
        : position;
    return setSelection({ sheetId, range: { start: anchor, end: position } });
  };
  const commit = (snapshot: WorkbookSnapshot, recalculate: boolean) => {
    let selection = state.selection;
    if (selection) {
      const parsed = selectionFor(snapshot, {
        sheetId: selection.sheetId,
        range: { start: selection.anchor, end: selection.focus },
      });
      selection = parsed.ok ? parsed.value : null;
    }
    state = Object.freeze({
      ...state,
      snapshot,
      selection,
      revision: state.revision + 1,
      canUndo: past.length > 0,
      canRedo: future.length > 0,
    });
    if (recalculate) rebuild();
    notify();
  };
  const dispatch = (command: WorkbookCommand): Result<void, CommandError> => {
    const commands = command.type === "batch" ? command.commands : [command];
    if (!commands.length) return ok(undefined);
    let candidate = state.snapshot;
    for (const item of commands) {
      const sheet = candidate.sheets.find((s) => s.id === item.sheetId);
      if (!sheet) return err(new CommandError("Unknown sheet"));
      let updated = sheet;
      if (item.type === "renameSheet") {
        // Renaming a referenced sheet requires formula rewriting, deferred deliberately.
        if (
          candidate.sheets.some((s) =>
            Object.values(s.cells).some((c) => c.formula),
          )
        )
          return err(
            new CommandError(
              "Renaming sheets with formulas is not yet supported",
            ),
          );
        updated = { ...sheet, name: item.name };
      } else if (item.type === "resizeColumn" || item.type === "resizeRow") {
        const column = item.type === "resizeColumn";
        const index = column ? item.column : item.row;
        const size = column ? item.width : item.height;
        if (
          !Number.isInteger(index) ||
          index < 0 ||
          index >= (column ? sheet.columnCount : sheet.rowCount)
        )
          return err(new CommandError("Dimension index outside sheet bounds"));
        if (
          size !== null &&
          (!Number.isFinite(size) || size <= 0 || size > 10000)
        )
          return err(
            new CommandError(
              "Dimension size must be greater than zero and at most 10000",
            ),
          );
        const dimensions = {
          ...(column ? sheet.columnWidths : sheet.rowHeights),
        };
        if (size === null) delete dimensions[String(index)];
        else dimensions[String(index)] = size;
        updated = column
          ? { ...sheet, columnWidths: dimensions }
          : { ...sheet, rowHeights: dimensions };
      } else if (item.type === "mergeCells" || item.type === "unmergeCells") {
        const parsed = parseCellRange(item.range);
        if (
          !parsed.ok ||
          parsed.value.end.row >= sheet.rowCount ||
          parsed.value.end.column >= sheet.columnCount
        )
          return err(new CommandError("Merge range outside sheet bounds"));
        const range = parsed.value;
        const intersected = (sheet.merges ?? []).filter((merge) =>
          rangesIntersect(merge, range),
        );
        if (item.type === "unmergeCells") {
          updated = {
            ...sheet,
            merges: (sheet.merges ?? []).filter(
              (merge) => !rangesIntersect(merge, range),
            ),
          };
        } else {
          if (
            range.start.row === range.end.row &&
            range.start.column === range.end.column
          )
            return err(new CommandError("Select at least two cells to merge"));
          if (intersected.some((merge) => !rangeContains(range, merge)))
            return err(
              new CommandError(
                "Select the whole existing merged region before merging",
              ),
            );
          for (const [key, cell] of Object.entries(sheet.cells)) {
            const position = parseAddress(key);
            if (!position.ok) continue;
            if (
              key !== addressOf(range.start) &&
              rangeContains(range, {
                start: position.value,
                end: position.value,
              }) &&
              (cell.value !== null || cell.formula)
            )
              return err(
                new CommandError(
                  "Move or clear values outside the top-left cell before merging",
                ),
              );
          }
          updated = {
            ...sheet,
            merges: [
              ...(sheet.merges ?? []).filter(
                (merge) => !rangesIntersect(merge, range),
              ),
              range,
            ],
          };
        }
      } else {
        const address = parseAddress(item.address);
        if (!address.ok) return err(new CommandError(address.error.message));
        const name = addressOf(
          mergedRangeAt(sheet, address.value)?.start ?? address.value,
        );
        const cells = { ...sheet.cells };
        if (item.type === "setCell") {
          const cell = parseCell(item.cell);
          if (!cell.ok) return err(new CommandError(cell.error.message));
          cells[name] = cell.value;
        } else delete cells[name];
        updated = {
          ...sheet,
          cells,
          rowCount: Math.max(sheet.rowCount, address.value.row + 1),
          columnCount: Math.max(sheet.columnCount, address.value.column + 1),
        };
      }
      candidate = {
        ...candidate,
        sheets: candidate.sheets.map((s) =>
          s.id === updated.id ? updated : s,
        ),
      };
    }
    const checked = parseWorkbookSnapshot(candidate);
    if (!checked.ok) return err(new CommandError(checked.error.message));
    // The current milestone rebuilds the graph for every atomic edit. It is the owner
    // of dependency edges; affectedBy is available when incremental recalc is introduced.
    if (historyLimit > 0) {
      past.push(state.snapshot);
      if (past.length > historyLimit) past.shift();
    }
    future.length = 0;
    commit(checked.value, true);
    return ok(undefined);
  };
  rebuild();
  return ok({
    getState: () => state,
    subscribe: (listener) => {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    dispatch,
    undo: () => {
      const snapshot = past.pop();
      if (!snapshot) return false;
      future.push(state.snapshot);
      commit(snapshot, true);
      return true;
    },
    redo: () => {
      const snapshot = future.pop();
      if (!snapshot) return false;
      past.push(state.snapshot);
      commit(snapshot, true);
      return true;
    },
    getCellValue: (sheetId, address) => {
      const p = parseAddress(address);
      return p.ok
        ? read(sheetId, addressOf(p.value), new Set())
        : cellError("#REF!");
    },
    setActiveSheet: (sheetId) => {
      if (!state.snapshot.sheets.some((s) => s.id === sheetId))
        return err(new CommandError("Unknown sheet"));
      state = Object.freeze({
        ...state,
        activeSheetId: sheetId,
        selection: null,
      });
      commit(state.snapshot, false);
      return ok(undefined);
    },
    setSelection,
    selectCell,
    moveSelection: (direction, options = {}) => {
      const sheet = state.snapshot.sheets.find(
        (s) => s.id === state.activeSheetId,
      );
      if (!sheet) return err(new CommandError("Unknown sheet"));
      const bounds = options.bounds ?? {
        start: { row: 0, column: 0 },
        end: { row: sheet.rowCount - 1, column: sheet.columnCount - 1 },
      };
      const checked = parseCellRange(bounds);
      if (
        !checked.ok ||
        checked.value.end.row >= sheet.rowCount ||
        checked.value.end.column >= sheet.columnCount
      )
        return err(new CommandError("Movement bounds outside sheet"));
      const area = checked.value;
      const readValue = (address: string) => {
        const result = read(sheet.id, address, new Set());
        return typeof result === "object" &&
          result !== null &&
          result.kind === "array"
          ? cellError("#VALUE!")
          : result;
      };
      if (!state.selection || state.selection.sheetId !== sheet.id) {
        let row = area.start.row,
          column = area.start.column;
        while (row <= area.end.row && !isRowVisible(sheet, row, readValue))
          row++;
        while (column <= area.end.column && !isColumnVisible(sheet, column))
          column++;
        if (row > area.end.row || column > area.end.column)
          return ok(undefined);
        return selectCell(sheet.id, { row, column });
      }
      const focus = state.selection.focus,
        merge = mergedRangeAt(sheet, focus) ?? { start: focus, end: focus };
      let row = focus.row,
        column = focus.column;
      switch (direction) {
        case "up":
          row = merge.start.row - 1;
          break;
        case "down":
          row = merge.end.row + 1;
          break;
        case "left":
          column = merge.start.column - 1;
          break;
        case "right":
          column = merge.end.column + 1;
          break;
      }
      if (direction === "up" || direction === "down") {
        const step = direction === "up" ? -1 : 1;
        while (
          row >= area.start.row &&
          row <= area.end.row &&
          !isRowVisible(sheet, row, readValue)
        )
          row += step;
      } else {
        const step = direction === "left" ? -1 : 1;
        while (
          column >= area.start.column &&
          column <= area.end.column &&
          !isColumnVisible(sheet, column)
        )
          column += step;
      }
      if (
        row < area.start.row ||
        row > area.end.row ||
        column < area.start.column ||
        column > area.end.column
      )
        return ok(undefined);
      return selectCell(sheet.id, { row, column }, options);
    },
    exportSnapshot: () => state.snapshot,
  });
}
