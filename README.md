# Nori

`@byfungsi/nori` is a TypeScript spreadsheet library with a canonical JSON model, a platform-independent runtime, an XLSX adapter, and composable React views. This is an initial working milestone, not full Excel compatibility. Nothing has been published to npm by this repository setup.

## Run locally

Node 22.12+ (Node 22 LTS recommended) and npm workspaces:

```sh
npm ci
npm run verify
npm run dev
```

Open the local URL printed by Vite. `npm run build` builds the public package and demo. Build once before starting the demo because it imports the actual public package output. After changing library source, rebuild it with `npm run build -w @byfungsi/nori`.

The demo includes two editable sheets, cross-sheet formulas, history controls, drag/Shift selection, dimension handles, merge/unmerge, XLSX upload, a semantic pivot panel, and a separate dark chat preview. The layout sample demonstrates imported widths, wraps, merged titles, frozen panes, hidden dimensions, and saved filters/sort. Download its sample XLSX and open it again to exercise the complete adapter/runtime/rendering path. `tests/fixtures/basic.xlsx` is also used in the integration suite; `node scripts/create-fixture.mjs` regenerates both copies.

## One public package

After a release, consumers install only `@byfungsi/nori`; React consumers also supply React 18.3 or 19. The facade bundles private workspace packages, so they are never required as separately published dependencies.

| Import                   | Responsibility                                                   |
| ------------------------ | ---------------------------------------------------------------- |
| `@byfungsi/nori`         | `parseXlsx`, `createWorkbook`, snapshot parsing and common types |
| `@byfungsi/nori/model`   | Canonical snapshots, cells, scalar errors, address parsing       |
| `@byfungsi/nori/xlsx`    | XLSX byte adapter and import diagnostics                         |
| `@byfungsi/nori/core`    | Commands, history, subscriptions, selection, layout              |
| `@byfungsi/nori/formula` | AST parser, evaluator, registry, dependency graph                |
| `@byfungsi/nori/pivot`   | Semantic grouping and aggregation                                |
| `@byfungsi/nori/react`   | Provider, hooks, grid, tabs, formula bar, convenience view       |

All entrypoints are ESM, include declarations, and declare no import-time side effects. React is an optional peer and appears only in the React entrypoint. Narrow subpaths let bundlers avoid loading XLSX code when only a runtime or model is needed. CommonJS is not supplied in this milestone.

```ts
import { parseXlsx, createWorkbook } from "@byfungsi/nori";

// bytes: Uint8Array, obtained by the host from a file, network, or native adapter.
const imported = parseXlsx(bytes);
if (!imported.ok) {
  console.error(imported.error.reason, imported.error.message);
} else {
  console.log(imported.value.warnings);
  const created = createWorkbook(imported.value.snapshot);
  if (created.ok) {
    const workbook = created.value;
    const sheetId = workbook.getState().activeSheetId;
    workbook.dispatch({
      type: "setCell",
      sheetId,
      address: "A1",
      cell: { value: 42 },
    });
    console.log(workbook.getCellValue(sheetId, "A1"));
    const json = JSON.stringify(workbook.exportSnapshot());
    // createWorkbook(JSON.parse(json)) validates rehydrated data again.
  }
}
```

Commands and parsers return `{ ok: true, value } | { ok: false, error }`. Callers should handle command rejection. Spreadsheet calculation errors are serializable cell values such as `{ kind: 'error', code: '#DIV/0!' }`.

## Composable React

```tsx
import {
  WorkbookProvider,
  WorkbookTabs,
  SheetGrid,
  FormulaBar,
} from "@byfungsi/nori/react";
import type { Workbook } from "@byfungsi/nori/core";

function Editor({ workbook }: { workbook: Workbook }) {
  return (
    <WorkbookProvider workbook={workbook}>
      <WorkbookTabs className="my-tabs" />
      <FormulaBar onError={(error) => console.error(error.message)} />
      <SheetGrid
        renderCell={({ value, address }) => (
          <span title={address}>
            {typeof value === "object" ? "" : String(value ?? "")}
          </span>
        )}
      />
    </WorkbookProvider>
  );
}
```

Use `Spreadsheet` for the default composition, or `useWorkbook` and `useWorkbookState` for your own UI. Create the runtime once outside render or in a lazy state initializer. A provider does not dispose or take ownership of its runtime; subscriptions clean themselves up when components unmount.

The library imports no global CSS. Style `.nori-grid`, `.nori-tabs`, `.nori-formula-bar`, `.nori-scroll`, and `[data-selected]`, pass class names, or replace cell content. Workbook `CellStyle` describes document formatting and is mapped separately to inline cell styles. Theme CSS should style the surrounding application rather than overwrite the workbook's data formatting. The demo stylesheet is an example, not a required theme.

`SheetGrid` renders at most 100 visible rows × 26 visible columns, excluding hidden/filtered entries. It respects document widths and heights and scrolls horizontally instead of compressing columns. Frozen leading rows and columns remain visible while scrolling. A host can pass a `range` for a moving window; full virtualization is not implemented.

Drag to select, Shift-click to extend from a stable anchor, or use arrows and Shift-arrows. Selection expands over merged cells. Drag row/column header boundaries to resize; a gesture creates one undo step. Escape cancels a resize; double-click a handle (or press Home while focused) to restore the document default. `SelectionToolbar` adds safe merge/unmerge actions; `Spreadsheet` includes it by default. See [editor interaction details](docs/interactions.md).

## Read-only chat preview

```tsx
import { SpreadsheetPreview } from "@byfungsi/nori/react";

<SpreadsheetPreview
  workbook={workbook}
  title="report.xlsx"
  theme="dark"
  maxRows={8}
  maxColumns={6}
  onOpen={() => openFullEditor()}
/>;
```

`SpreadsheetPreview` is a separate, responsive attachment card for agentic chat. It has no editors, resize handles or mutation controls. Its sheet tabs are local to the preview; they never change the full editor's active sheet or selection. It subscribes to calculated values, supports light/dark themes, and scrolls inside its parent on narrow screens. `onOpen` is optional. See the [preview API](docs/preview.md).

See [architecture](docs/architecture.md), [API](docs/api.md), [support matrix](docs/support.md), and [roadmap](docs/roadmap.md).

## Verification

`npm run verify` performs package-boundary checks, strict TypeScript checks (including headless compilation without DOM or Node ambient types), behavioral/property tests, production builds, and an isolated packed-package consumer test. The latter checks public declaration types, headless imports with no React installation, all subpaths, bundler isolation, and React SSR. CI runs this command on Node 22.

For a navigable view without editing, use `<Spreadsheet workbook={workbook} readOnly />`. Composed views can set `readOnly` on `WorkbookProvider`; custom controls can read `useWorkbookReadOnly()`. Editable grids support double-click or Enter/F2 to edit inline, Enter to save, and Escape to cancel. See [editor interactions](docs/interactions.md).

## Documentation app

The VitePress app in `docs/` provides guides, API references, local search and agent-oriented exports. Run `npm run docs:dev` or build with `npm run docs:build` and serve with `npm run docs:preview` (port 4179). The build generates `llms.txt`, `llms-full.txt`, per-page Markdown, public declarations and a versioned content manifest. See [the agent guide](docs/agents.md) and [contributing instructions](docs/contributing.md).

Source: [byfungsi/nori](https://github.com/byfungsi/nori) · Documentation: [GitHub Pages](https://byfungsi.github.io/nori/) · [Agent index](https://byfungsi.github.io/nori/llms.txt)

[Try the live spreadsheet playground](https://byfungsi.github.io/nori/playground.html) — edit cells, recalculate formulas, switch sheets, and open a local XLSX file without installing anything.

CSV import: `import { parseCsv } from "@byfungsi/nori/csv"`. Quoted and multiline fields are supported; text is preserved by default, with optional number/boolean inference. [CSV guide](https://byfungsi.github.io/nori/csv.html). The live playground accepts CSV and XLSX files.
