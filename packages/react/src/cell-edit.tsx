import { useRef, useState, type ReactNode } from "react";
import type { Cell, Scalar } from "@nori-internal/model";
import { formatCellValue } from "./cell-format";

export function cellDraft(cell: Cell | undefined): string {
  return cell?.formula
    ? "=" + cell.formula
    : formatCellValue(cell?.value ?? null);
}
export function editedCell(draft: string, cell: Cell | undefined): Cell {
  let value: Scalar = draft;
  if (draft === "") value = null;
  else if (/^(TRUE|FALSE)$/i.test(draft))
    value = draft.toUpperCase() === "TRUE";
  else if (draft.trim() && Number.isFinite(Number(draft)))
    value = Number(draft);
  return {
    value: draft.startsWith("=") ? null : value,
    ...(draft.startsWith("=") ? { formula: draft.slice(1) } : {}),
    ...(cell?.style ? { style: cell.style } : {}),
  };
}
export function InlineCellEditor({
  cell,
  address,
  commit,
  cancel,
}: {
  cell: Cell | undefined;
  address: string;
  commit: (draft: string) => boolean;
  cancel: () => void;
}): ReactNode {
  const [draft, setDraft] = useState(() => cellDraft(cell));
  const finished = useRef(false);
  const save = () => {
    if (!finished.current) finished.current = commit(draft);
  };
  return (
    <input
      autoFocus
      aria-label={`Edit ${address}`}
      value={draft}
      onFocus={(event) => event.currentTarget.select()}
      onChange={(event) => setDraft(event.target.value)}
      onPointerDown={(event) => event.stopPropagation()}
      onDoubleClick={(event) => event.stopPropagation()}
      onBlur={save}
      onKeyDown={(event) => {
        event.stopPropagation();
        if (event.key === "Enter") {
          event.preventDefault();
          save();
        }
        if (event.key === "Escape") {
          event.preventDefault();
          finished.current = true;
          cancel();
        }
      }}
      style={{
        position: "absolute",
        inset: 0,
        width: "100%",
        height: "100%",
        minWidth: 0,
        boxSizing: "border-box",
        border: "2px solid var(--nori-selection-border, #789a55)",
        borderRadius: 0,
        padding: "0 8px",
        background: "var(--nori-editor-background, white)",
        color: "var(--nori-editor-color, #17251b)",
        font: "inherit",
        zIndex: 2,
      }}
    />
  );
}
