# Styling and theming

Nori 0.1.0 exposes styling hooks for React applications. The full editor uses your application's CSS; the compact preview has built-in light and dark palettes. There is no theme picker in the playground, no full-editor `theme` prop, and no exported Nori stylesheet to import.

This guide describes the published API, including the limits of what can be customized without changing the renderer.

## What `className` means

`className` is a React prop that adds a CSS class to a rendered HTML element. It is not a visible label, a workbook field, or a setting in Excel. You choose the name and write its rules in your application stylesheet:

```tsx
import { Spreadsheet } from "@byfungsi/nori/react";
import "./spreadsheet.css";

// workbook is a live runtime created by createWorkbook(...).
<Spreadsheet workbook={workbook} className="my-sheet" />;
```

```css
/* spreadsheet.css in your application */
.my-sheet {
  font-family: system-ui, sans-serif;
  color: #24372c;
  background: white;
  --nori-selection-border: #286b48;
}
```

Inspect the element in browser developer tools: the editor's outer `div` now has `class="my-sheet"`. The CSS file must actually be loaded by your application. Installing Nori does not install these example styles automatically.

**A supplied `className` replaces the component's default class.** It is not appended. To retain the default root hook, use `className="nori-workbook my-sheet"`. Changing the editor's root class does not change its children's default classes, such as `.nori-grid`.

## Three separate styling layers

| Layer                           | Examples                                                                | Where it lives                                   |
| ------------------------------- | ----------------------------------------------------------------------- | ------------------------------------------------ |
| Application appearance          | Toolbar colors, tabs, borders, focus, font family                       | Your CSS and React styling props                 |
| Workbook formatting             | Explicit cell fill/text color, bold, wrapping, alignment, number format | `Cell.style` in the snapshot/runtime             |
| Layout and interaction geometry | Column widths, row heights, merged cells, frozen offsets, resize guides | Workbook layout and the renderer's inline styles |

Changing a theme does not mutate the workbook or its exported snapshot. Imported explicit cell colors take precedence over ordinary inherited application colors. A pale imported fill with no explicit text color can be hard to read in a dark application. Review representative workbooks; dark application chrome does not guarantee dark-compatible document formatting.

## Component styling props

| Component            | Props and target                                                                                     |
| -------------------- | ---------------------------------------------------------------------------------------------------- |
| `Spreadsheet`        | `className`, `style` target the outer `div`; `renderCell` customizes cell contents                   |
| `SheetGrid`          | `className` targets the **table**, not its scrolling wrapper; supports `renderCell`; no `style` prop |
| `FormulaBar`         | `className` targets its outer `div` or `form`, depending on selection; no `style` prop               |
| `SelectionToolbar`   | `className` targets the outer `div`; no `style` prop                                                 |
| `WorkbookTabs`       | `className` targets the `nav`; no `style` prop                                                       |
| `SpreadsheetPreview` | `className`, `style` target the outer `figure`; `theme` accepts `light` or `dark`                    |
| `WorkbookProvider`   | Context only; no visual wrapper or styling props                                                     |

The editor does not forward arbitrary HTML attributes such as `data-theme`. Put those attributes on a host wrapper if you need them. The preview sets its own `data-theme` from its `theme` prop.

## Complete editor CSS starter

Copy the following CSS into your application. It includes light and dark application palettes, formula controls, selection toolbar, grid borders, sheet tabs, keyboard focus, and narrow-screen wrapping. These are application-owned examples, not built-in Nori themes. Explicit workbook colors are preserved.

<<< ./examples/styling.css

Use either appearance with the same runtime:

```tsx
<Spreadsheet workbook={workbook} className="nori-workbook my-sheet" />

<Spreadsheet
  workbook={workbook}
  className="nori-workbook my-sheet my-sheet--dark"
/>
```

In a real application, render one editor and toggle the class from your existing theme state:

```tsx
<Spreadsheet
  workbook={workbook}
  className={`nori-workbook my-sheet${darkMode ? " my-sheet--dark" : ""}`}
/>
```

`darkMode` is your application's boolean, not a Nori export. No workbook recreation is needed when it changes. The `--app-*` variables in this example belong to this stylesheet; Nori itself only reads the `--nori-*` variables listed below.

## All editor CSS variables

Set these on the editor root or an ancestor so they inherit into the grid. They affect the full editor, including composed primitives, but do not theme `SpreadsheetPreview`.

| Variable                   | Built-in fallback       | Effect                                          |
| -------------------------- | ----------------------- | ----------------------------------------------- |
| `--nori-selection-fill`    | `rgba(107,151,71,0.13)` | Overlay fill on selected cells                  |
| `--nori-selection-border`  | `#789a55`               | Selection inset border and inline editor border |
| `--nori-resize-border`     | `#4b7d34`               | Full boundary guide while hovering/resizing     |
| `--nori-header-background` | `#f6f8f1`               | Row/column header backgrounds                   |
| `--nori-cell-background`   | `white`                 | Frozen cells without an explicit workbook fill  |
| `--nori-editor-background` | `white`                 | Inline editing input background                 |
| `--nori-editor-color`      | `#17251b`               | Inline editing input text                       |

`--nori-cell-background` alone does **not** paint every ordinary cell. Set the root/table or ordinary `td` background in your CSS, as in the starter. Grid lines, header text, formula-bar input styling, tab colors, and font family also need your CSS; they are not additional built-in tokens.

### Inline styles and TypeScript

`style` is useful for per-instance values. TypeScript's `CSSProperties` does not directly list custom properties; explicitly extend it:

```tsx
import type { CSSProperties } from "react";

const appearance: CSSProperties & {
  "--nori-selection-border": string;
  "--nori-selection-fill": string;
} = {
  "--nori-selection-border": "#7856d8",
  "--nori-selection-fill": "rgb(120 86 216 / 15%)",
  borderRadius: 8,
};

<Spreadsheet workbook={workbook} className="my-sheet" style={appearance} />;
```

Use `style` for the root only. It does not automatically override inline styles on descendants.

## CSS hooks and interaction states

Scope your selectors beneath your editor class to avoid changing other tables or buttons in your app.

| Hook                                     | What it selects                                                                     |
| ---------------------------------------- | ----------------------------------------------------------------------------------- |
| `.nori-workbook`                         | Default editor root                                                                 |
| `.nori-formula-bar`                      | Formula display/edit controls                                                       |
| `.nori-selection-toolbar`                | Selection summary and merge/unmerge controls                                        |
| `.nori-view-status`                      | Imported filter/freeze status, when present                                         |
| `.nori-scroll`                           | Scroll container generated by `SheetGrid`                                           |
| `.nori-grid`                             | Default grid table                                                                  |
| `.nori-tabs`                             | Sheet navigation                                                                    |
| `.nori-tabs button[aria-pressed="true"]` | Active sheet button                                                                 |
| `.nori-grid td[data-selected]`           | Selected cells; attribute is absent when unselected                                 |
| `.nori-grid td[data-merged]`             | Merged anchor cells                                                                 |
| `.nori-grid td[data-address="B2"]`       | A rendered cell by its A1 address                                                   |
| `.nori-grid button[data-cell-address]`   | Cell interaction buttons                                                            |
| `.nori-selection-fill`                   | Noninteractive selection overlay                                                    |
| `.nori-resize-handle`                    | Header resize handle; axis classes are `.nori-resize-column` and `.nori-resize-row` |
| `.nori-resize-guide`                     | Whole-border hover/drag indicator                                                   |
| `[role="alert"]`                         | Error text; scope this to your editor                                               |

Do not remove pointer-event behavior, change sticky positioning/z-index, or repurpose resize handles to draw ordinary grid borders. Keep the selection overlay noninteractive. Do not hide keyboard focus outlines without providing an equally visible replacement.

These hooks describe the current 0.1.0 DOM. Prefer public props and variables; audit descendant selectors when upgrading, especially selectors tied to internal nesting.

## Compose a custom toolbar layout

Use the primitives when the default editor arrangement does not fit your application:

```tsx
import {
  WorkbookProvider,
  FormulaBar,
  SelectionToolbar,
  SheetGrid,
  WorkbookTabs,
} from "@byfungsi/nori/react";

<div className="my-sheet my-sheet--dark">
  <WorkbookProvider workbook={workbook}>
    <WorkbookTabs className="nori-tabs project-tabs" />
    <FormulaBar className="nori-formula-bar project-formula" />
    <SelectionToolbar />
    <SheetGrid className="nori-grid project-grid" />
  </WorkbookProvider>
</div>;
```

Retaining the default class names allows the starter stylesheet to continue matching. The same styling works in a `readOnly` editor; read-only mode changes interaction policy, not appearance. Style disabled actions and read-only inputs if you want a visual distinction.

## Custom cell content

`renderCell` changes display content inside an existing cell button. It does not replace the `td`, change the stored value, or change calculations/exported formatting.

```tsx
import { Spreadsheet, formatCellValue } from "@byfungsi/nori/react";

<Spreadsheet
  workbook={workbook}
  className="my-sheet"
  renderCell={({ value, cell }) => (
    <span style={{ fontVariantNumeric: "tabular-nums" }}>
      {formatCellValue(value, cell)}
    </span>
  )}
/>;
```

The context includes `sheet`, `address`, `cell`, calculated `value`, `selected`, and `mergedRange`. `value` may be a scalar, error, or array result. `formatCellValue` handles these current result forms; avoid calling numeric methods without narrowing the value first. Merged cells use their anchor's value.

Render noninteractive content: nested buttons, links, and inputs inside the cell button interfere with semantics and editing. The built-in title and accessible cell label continue to use the default formatted value, so custom content should preserve its meaning. The preview has no `renderCell` prop in 0.1.0.

## Styling the chat preview

```tsx
import { SpreadsheetPreview } from "@byfungsi/nori/react";

<SpreadsheetPreview
  workbook={workbook}
  title="sales.xlsx"
  theme="dark"
  className="nori-preview chat-attachment"
  style={{ borderRadius: 16, maxWidth: 720 }}
  maxRows={8}
  maxColumns={6}
/>;
```

The preview is always read-only, with local sheet navigation. Its built-in palette styles the header, grid lines, tabs, footer, and text. `style` is merged last on the outer figure, so outer sizing, border, background, and font family can be customized there. The internal table uses a fixed 12px font size; changing the figure's font size does not scale all descendants.

The preview uses inline colors on internal elements. A root background override or the editor's CSS variables will not replace that full palette. Version 0.1.0 has no custom palette object, theme provider, or automatic system-theme detection. Pass `theme` from your host's theme state. Extensive internal recoloring would require overriding inline declarations; prefer the built-in palettes or a host-owned renderer for that level of control.

The card fits its container and scrolls columns horizontally. Default limits are 8 rows and 6 columns; maximum limits are 20 rows and 12 columns. Column widths are clamped to 90–180px for a compact card. Imported row heights are not reproduced; text is ellipsized. Workbook fills and text colors remain separate from its palette. See [chat preview](/preview) for the full prop reference.

## Responsive layout, sizing, and specificity

- Give flex/grid children containing the editor `min-width: 0`. For CSS grid, `minmax(0, 1fr)` prevents a wide sheet from stretching the page.
- Let `.nori-scroll` scroll horizontally. Do not squeeze all spreadsheet columns to fit a phone or override the renderer's calculated table/column widths.
- Row heights, column widths, frozen offsets, and merge geometry come from the workbook. Change them through supported runtime commands, not arbitrary cell CSS. Font-size increases do not auto-fit rows.
- `SheetGrid` renders a bounded window (up to 100 visible rows and 26 visible columns); styling does not enable virtualization or remove that limit.
- The scroll container has an inline `max-height: 520px`. There is no public height prop in 0.1.0. If necessary, a scoped `.my-sheet .nori-scroll { max-height: 65vh !important; }` overrides that one declaration. Treat it as a version-specific workaround and test frozen panes and resize guides afterward.
- Essential cell padding, positioning, button geometry, and workbook formatting are inline. Ordinary CSS rules do not override inline declarations. Avoid broad `!important` rules that erase workbook styles or break interaction geometry.
- The starter uses `background-color` on `td` so an explicit inline workbook background wins. Avoid blanket `td { color: ... }` if you intend unformatted cells to inherit root text color.

### CSS Modules and utility CSS

With CSS Modules, use `className={styles.sheet}` and `:global(.nori-grid)` inside the scoped rules to target Nori's global descendant classes. Import the module in your application; there is no module export from Nori.

Utility classes can style the root through `className`. For the many descendant hooks and custom properties, a scoped stylesheet is usually clearer. No Tailwind plugin or theme adapter ships with Nori.

## Troubleshooting and validation

| Symptom                                | Check                                                                                                                    |
| -------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| My class does nothing                  | Confirm the CSS file is loaded and the rendered element has that class. A class name does not generate styles by itself. |
| Default selector stopped matching      | A custom `className` replaces the default; pass both names if needed.                                                    |
| Grid headers remain light              | Set `--nori-header-background`; their background is assigned inline using this variable.                                 |
| Frozen cells remain white              | Set `--nori-cell-background`; also style ordinary cells separately.                                                      |
| Toolbar input remains light            | Style `.nori-formula-bar input`; editor variables only style the inline cell editor.                                     |
| Imported cell looks wrong in dark mode | Inspect its explicit workbook fill and text color; application theme changes preserve them.                              |
| My row size or padding rule is ignored | Inspect inline layout styles. Use layout commands for dimensions and avoid breaking cell geometry.                       |
| Preview ignores editor variables       | Use its `theme` prop; preview colors are a separate implementation.                                                      |
| CSS affects another workbook           | Scope rules beneath a per-instance class; avoid global `table`, `button`, or `td` rules.                                 |

Before shipping your stylesheet, check keyboard focus, selected ranges, active tabs, disabled/read-only controls, formula input, inline editing, error text, hover/drag resize borders, frozen rows/columns, merged cells, long text, and imported explicit colors. Check narrow screens and horizontal scrolling as well as a desktop viewport. Maintain readable contrast; do not use color alone to communicate errors or selection.

There is currently no unified theme object, full-editor light/dark preset, theme persistence, exported CSS bundle, or automatic contrast correction for imported formatting. The [React integration guide](/react), [interaction guide](/interactions), and [Excel support matrix](/support) describe the surrounding capabilities.
