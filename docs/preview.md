# Chat / thumbnail preview

`SpreadsheetPreview` is a separate read-only React component intended for agentic chat messages, file attachments and thumbnails. It does not embed the editable grid.

```tsx
import { SpreadsheetPreview } from "@byfungsi/nori/react";

<SpreadsheetPreview
  workbook={workbook}
  title="mrr-2026.xlsx"
  theme="dark"
  maxRows={8}
  maxColumns={6}
  onOpen={() => showEditor(workbook)}
/>;
```

Props:

| Prop                        | Behavior                                                                      |
| --------------------------- | ----------------------------------------------------------------------------- |
| `workbook`                  | Live runtime; subscribes to data/calculation updates but never mutates it     |
| `title`                     | Attachment filename; default `Workbook.xlsx`                                  |
| `theme`                     | `light` (default) or `dark`                                                   |
| `maxRows` / `maxColumns`    | Defaults 8/6; bounded to 20/12 for thumbnail rendering                        |
| `onOpen`                    | Optional action to open the host's full editor; no action is shown if omitted |
| `sheetId` / `onSheetChange` | Optional controlled tab; otherwise tab state is local to the card             |
| `className` / `style`       | Host styling hooks on the outer figure                                        |

Preview sheet changes do not alter the full editor's active sheet, selection or undo history. There are no cell editors, selection gestures, resize handles or merge controls. It honors hidden/filtered rows and columns, shows calculated values and merged-cell anchor content.

The card fits its parent's width, including narrow chat bubbles. The table and tab strip scroll horizontally inside the card; they never force the page wider. Headers can wrap the file/open controls. Cells use compact fixed row heights and bounded column widths, intentionally independent of the full editor's large document geometry. Long values are clipped with full text in their title attribute. The footer states the displayed row/column counts.

Workbook colors/styles are preserved separately from the card theme. Imported explicit colors may need host-specific adaptation for strong contrast in a dark chat design. Full font fidelity, annotations such as hand-drawn circles, and automatic row fitting are not part of this component.

See [Styling and theming](/styling#styling-the-chat-preview) for palette behavior, CSS precedence, sizing, and the distinction between preview and editor themes.
