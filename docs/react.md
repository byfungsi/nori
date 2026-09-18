# React integration

## Ready-made editor

<<< ./examples/react.tsx

The helper `makeWorkbook` is the [quick-start example](/getting-started#calculate-and-persist-a-workbook), not a library export. Create or import the runtime once, outside render or in a lazy state initializer. The renderer subscribes with `useSyncExternalStore` and also supports server rendering.

## Compose your own view

```tsx
import {
  WorkbookProvider,
  FormulaBar,
  SelectionToolbar,
  SheetGrid,
  WorkbookTabs,
} from "@byfungsi/nori/react";

// `workbook` is the live runtime created by your host.
<WorkbookProvider workbook={workbook} readOnly={false}>
  <FormulaBar />
  <SelectionToolbar />
  <SheetGrid />
  <WorkbookTabs />
</WorkbookProvider>;
```

| Primitive               | Responsibility                                                   |
| ----------------------- | ---------------------------------------------------------------- |
| `WorkbookProvider`      | Supplies the runtime and optional `readOnly` policy              |
| `useWorkbook()`         | Returns the runtime for host controls                            |
| `useWorkbookState()`    | Subscribes to stable state snapshots                             |
| `useWorkbookReadOnly()` | Reads the provider's view policy                                 |
| `FormulaBar`            | Displays/edits the selection anchor's raw value or formula       |
| `SelectionToolbar`      | Shows selection range and merge actions                          |
| `SheetGrid`             | Bounded grid, selection, inline editing, resize, imported layout |
| `WorkbookTabs`          | Switches runtime active sheet and clears selection               |
| `Spreadsheet`           | Default composition of the above                                 |
| `SpreadsheetPreview`    | Independent, always-read-only chat card                          |

`Spreadsheet` accepts `workbook`, `readOnly`, `className`, `style`, `range`, and `renderCell`. `SheetGrid` accepts `range`, `renderCell`, `className`, `readOnly`, and `onError`. `FormulaBar` and `SelectionToolbar` accept `className`, `readOnly`, and `onError`. A child primitive cannot disable its provider's read-only policy.

`range` is an inclusive zero-based rectangle. Rendering is capped at 100 visible rows and 26 visible columns, including frozen entries. It is not full virtualization.

## Read-only is a view policy

`readOnly` blocks built-in document mutations, including resizing and merging. Selection, arrows and sheet switching remain usable. A host-owned undo button must check `useWorkbookReadOnly()` itself. The runtime command API remains available for programmatic updates; this flag is not an authorization boundary.

The [chat preview](/preview) always stays read-only and keeps tab state local. Use it for messages and thumbnails; use `Spreadsheet readOnly` for a navigable full grid.

## Styling and custom cells

Nori supplies essential geometry inline and semantic CSS classes for application styling. It does not export a stylesheet or impose a design-system dependency. Style `.nori-workbook`, `.nori-grid`, `.nori-tabs`, `.nori-formula-bar`, and `.nori-selection-toolbar` in your application. The repository's demo CSS is an example, not a public package export.

Workbook styles (fill, text color, alignment and number-format metadata) belong to cells. Application chrome belongs to CSS. Customize selection through `--nori-selection-fill`, `--nori-selection-border`; resize through `--nori-resize-border`; frozen surfaces through `--nori-header-background`, `--nori-cell-background`; and inline editors through `--nori-editor-background`, `--nori-editor-color`.

`renderCell(context)` receives `sheet`, `address`, `cell`, calculated `value`, `selected`, and `mergedRange`. Render noninteractive content: it is placed inside the grid's cell button, so nested buttons/inputs are invalid. Merged cells receive their anchor's value. See [interactions](/interactions) for editing and keyboard behavior.
