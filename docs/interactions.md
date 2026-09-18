# Editor interactions

- **Drag selection:** press the primary pointer in a cell and drag across the grid. Pointer capture continues the gesture outside the initial cell. At the scrollport edge, the view auto-scrolls. Releasing commits the current selection; cancellation or loss of capture stops the gesture.
- **Shift-click:** extend from the original anchor to another cell; repeated extensions retain that anchor, including when extending in reverse. Arrow keys move focus; Shift-arrows extend/contract the selection. Keyboard movement is bounded to the rendered window, skips hidden/filtered entries, and steps past a merged region as a unit.
- **Resize:** drag a column or row boundary from its header or anywhere along a cell edge. A six-unit hit area on either side of the boundary gives a stable target; hovering or dragging highlights the whole boundary across the viewport. At intersections the nearest edge wins, with columns winning exact ties. Merged cells expose only their perimeter. The drag is a temporary preview until release, creating a single undoable command. Escape or pointer cancellation restores the original size. Handles are keyboard-focusable: directional arrows change size by 10 units, Shift-arrows by one, and Home resets to the sheet default. Double-click also resets; it is not auto-fit.
- **Merge:** select a rectangle and choose Merge cells in `SelectionToolbar`. The top-left cell is the only content owner. A nonempty covered cell causes a useful error instead of data loss. Unmerge cells removes intersecting merges without losing the anchor. Both commands participate in undo/redo and snapshot serialization.

The model and core own all document/selection semantics. React owns pointer capture, DOM hit testing, animation-frame auto-scroll, keyboard events and temporary resize feedback. `Spreadsheet` is the default composition; these behaviors are also available through `SheetGrid` and the separate `SelectionToolbar`.

The full grid shows at most 100 visible rows and 26 visible columns, retaining leading frozen dimensions when given a later window. This is bounded rendering, not full virtualization. Large frozen regions can occupy the entire rendered window.

Style `.nori-resize-handle`, `.nori-selection-toolbar`, `.nori-view-status`, and the existing grid classes. The full resize guide uses `--nori-resize-border`. Selection colors can be set with `--nori-selection-fill` and `--nori-selection-border`. Frozen headers/cells use `--nori-header-background` and `--nori-cell-background` when the document has not supplied a background.

## Editing and read-only views

Double-click a cell, or press Enter/F2 on a focused cell, to edit its value or formula inline. Enter or moving focus commits through workbook commands; Escape discards the draft. There is no modal overlay.

Use `<Spreadsheet workbook={workbook} readOnly />` or `<WorkbookProvider workbook={workbook} readOnly>` for composed views. `SheetGrid`, `FormulaBar`, and `SelectionToolbar` also accept `readOnly`. A provider's read-only setting cannot be overridden by its children. Selection, navigation and sheet tabs remain available; editors, resize gestures and merge actions cannot change the document. Switching to read-only cancels an unfinished inline edit or resize.

`useWorkbookReadOnly()` lets host-owned controls disable undo/redo or other mutation actions. This is a view policy: the host can still update the headless workbook programmatically. The separate `SpreadsheetPreview` always remains read-only.
