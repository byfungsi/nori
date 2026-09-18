import { useState, type ReactNode } from "react";
import { addressOf, rangesIntersect } from "@nori-internal/model";
import type { CommandError } from "@nori-internal/core";
import {
  useWorkbook,
  useWorkbookState,
  useWorkbookReadOnly,
} from "./workbook-context";
/** Composable selection summary and safe merge/unmerge actions, backed by undoable commands. */
export function SelectionToolbar({
  className,
  onError,
  readOnly: readOnlyProp = false,
}: {
  readonly className?: string;
  readonly readOnly?: boolean;
  readonly onError?: (error: CommandError) => void;
}): ReactNode {
  const workbook = useWorkbook(),
    state = useWorkbookState();
  const readOnly = useWorkbookReadOnly() || readOnlyProp;
  const [error, setError] = useState<{
    revision: number;
    message: string;
  } | null>(null);
  const selection =
    state.selection?.sheetId === state.activeSheetId ? state.selection : null;
  const sheet = state.snapshot.sheets.find((s) => s.id === state.activeSheetId);
  const start = selection ? addressOf(selection.range.start) : "",
    end = selection ? addressOf(selection.range.end) : "";
  const intersects =
    !!selection &&
    !!sheet?.merges?.some((merge) => rangesIntersect(merge, selection.range));
  const exact =
    !!selection &&
    !!sheet?.merges?.some(
      (merge) =>
        addressOf(merge.start) === start && addressOf(merge.end) === end,
    );
  const dispatch = (type: "mergeCells" | "unmergeCells") => {
    if (readOnly || !selection) return;
    const result = workbook.dispatch({
      type,
      sheetId: selection.sheetId,
      range: selection.range,
    });
    if (!result.ok) {
      setError({ revision: state.revision, message: result.error.message });
      onError?.(result.error);
    } else setError(null);
  };
  return (
    <div className={className ?? "nori-selection-toolbar"}>
      <output aria-label="Selected range">
        {selection
          ? start === end
            ? start
            : `${start}:${end}`
          : "No selection"}
      </output>
      <button
        type="button"
        disabled={readOnly || !selection || start === end || exact}
        onClick={() => dispatch("mergeCells")}
      >
        Merge cells
      </button>
      <button
        type="button"
        disabled={readOnly || !intersects}
        onClick={() => dispatch("unmergeCells")}
      >
        Unmerge cells
      </button>
      {error?.revision === state.revision && (
        <span role="alert">{error.message}</span>
      )}
    </div>
  );
}
