import { useState, type ReactNode } from "react";
import { createRoot } from "react-dom/client";
import {
  createWorkbook,
  parseXlsx,
  parseCsv,
  type Workbook,
} from "@byfungsi/nori";
import {
  WorkbookProvider,
  WorkbookTabs,
  SheetGrid,
  FormulaBar,
  SelectionToolbar,
  SpreadsheetPreview,
  useWorkbook,
  useWorkbookState,
  useWorkbookReadOnly,
} from "@byfungsi/nori/react";
import { createPivot } from "@byfungsi/nori/pivot";
import "./style.css";
const initial = createWorkbook({
  version: 1,
  dateSystem: "1900",
  sheets: [
    {
      id: "sales",
      name: "Sales",
      rowCount: 14,
      columnCount: 6,
      merges: [{ start: { row: 8, column: 0 }, end: { row: 8, column: 3 } }],
      cells: {
        A9: {
          value: "Drag to select · Shift-click to extend",
          style: { fontWeight: "bold", background: "#EDF4E6" },
        },
        A1: { value: "Region", style: { fontWeight: "bold" } },
        B1: { value: "Revenue", style: { fontWeight: "bold" } },
        A2: { value: "West" },
        B2: { value: 1200 },
        A3: { value: "East" },
        B3: { value: 850 },
        A4: { value: "West" },
        B4: { value: 600 },
        A6: { value: "Total" },
        B6: {
          value: null,
          formula: "SUM(B2:B4)",
          style: { fontWeight: "bold", background: "#DDF4D8" },
        },
      },
    },
    {
      id: "summary",
      name: "Summary",
      rowCount: 14,
      columnCount: 6,
      cells: {
        A1: { value: "Sales total", style: { fontWeight: "bold" } },
        B1: { value: null, formula: "Sales!B6" },
        A2: { value: "Average order" },
        B2: { value: null, formula: "AVERAGE(Sales!B2:B4)" },
        A4: { value: "Try editing Sales B2" },
      },
    },
  ],
});
if (!initial.ok) throw initial.error;
const startingSnapshot = initial.value.exportSnapshot();
function Toolbar(): ReactNode {
  const readOnly = useWorkbookReadOnly();
  const workbook = useWorkbook(),
    state = useWorkbookState();
  return (
    <div className="toolbar">
      <span>
        {state.snapshot.sheets.length} sheets · {state.snapshot.dateSystem} date
        system
      </span>
      <div>
        <button
          disabled={readOnly || !state.canUndo}
          onClick={() => workbook.undo()}
        >
          Undo
        </button>
        <button
          disabled={readOnly || !state.canRedo}
          onClick={() => workbook.redo()}
        >
          Redo
        </button>
      </div>
    </div>
  );
}
function PivotPreview(): ReactNode {
  const workbook = useWorkbook();
  const state = useWorkbookState();
  const sales = state.snapshot.sheets.find((s) => s.name === "Sales");
  if (!sales) return null;
  const records = [2, 3, 4].map((row) => {
    const region = workbook.getCellValue(sales.id, `A${row}`),
      revenue = workbook.getCellValue(sales.id, `B${row}`);
    return {
      region: typeof region === "string" ? region : "",
      revenue: typeof revenue === "number" ? revenue : null,
    };
  });
  const pivot = createPivot(records, {
    rows: ["region"],
    values: [{ id: "revenue", field: "revenue", aggregate: "sum" }],
  });
  return pivot.ok ? (
    <aside>
      <h2>Revenue by region</h2>
      <p>A headless pivot, rendered independently.</p>
      {pivot.value.groups.map((group) => (
        <div className="pivot-row" key={String(group.key[0])}>
          <span>{String(group.key[0])}</span>
          <strong>{group.values["revenue"]}</strong>
        </div>
      ))}
    </aside>
  ) : null;
}
async function importFile(file: File) {
  if (file.name.toLowerCase().endsWith(".csv")) {
    const parsed = parseCsv(await file.text(), { valueMode: "infer" });
    return parsed.ok
      ? { ...parsed, value: { ...parsed.value, warnings: [] } }
      : parsed;
  }
  return parseXlsx(new Uint8Array(await file.arrayBuffer()));
}
function App({ initialWorkbook }: { initialWorkbook: Workbook }): ReactNode {
  const [workbook, setWorkbook] = useState(initialWorkbook),
    [filename, setFilename] = useState("sales-workbook.xlsx"),
    [message, setMessage] = useState(""),
    [busy, setBusy] = useState(false),
    [readOnly, setReadOnly] = useState(false);
  return (
    <main
      className={
        new URLSearchParams(location.search).has("embed")
          ? "embedded"
          : undefined
      }
    >
      <header>
        <div className="eyebrow">BY FUNGSI / NORI</div>
        <h1>
          A little space
          <br />
          for big ideas<span>.</span>
        </h1>
        <p>Open a workbook. Explore your sheets. Make the numbers yours.</p>
        <label className="upload">
          {busy ? "Opening…" : "Open .xlsx or .csv"}
          <input
            type="file"
            accept=".xlsx,.csv,text/csv"
            disabled={busy}
            onChange={async (event) => {
              const file = event.target.files?.[0];
              if (!file) return;
              setBusy(true);
              try {
                const imported = await importFile(file);
                if (!imported.ok) {
                  setMessage(imported.error.message);
                  return;
                }
                const runtime = createWorkbook(imported.value.snapshot);
                if (!runtime.ok) {
                  setMessage(runtime.error.message);
                  return;
                }
                setWorkbook(runtime.value);
                setFilename(file.name);
                const warnings = imported.value.warnings;
                setMessage(
                  warnings.length
                    ? `Opened ${file.name}. Import notes: ${[...new Set(warnings.map((w) => w.feature))].join(", ")}.`
                    : `Opened ${file.name}.`,
                );
              } catch {
                setMessage("The file could not be read.");
              } finally {
                setBusy(false);
                event.target.value = "";
              }
            }}
          />
        </label>
        <a className="sample" href="./sample.xlsx" download>
          Download sample workbook
        </a>
        <a className="sample" href="./layout.xlsx" download>
          Download layout sample
        </a>
        <a className="sample" href="./sample.csv" download>
          Download CSV sample
        </a>
      </header>
      <button
        type="button"
        className="reset-example"
        onClick={() => {
          const reset = createWorkbook(startingSnapshot);
          if (!reset.ok) {
            setMessage(reset.error.message);
            return;
          }
          setWorkbook(reset.value);
          setFilename("sales-workbook.xlsx");
          setMessage("");
          setReadOnly(false);
        }}
      >
        Reset example
      </button>
      {message && <p role="status">{message}</p>}
      <label style={{ display: "block", marginBottom: 12 }}>
        <input
          type="checkbox"
          checked={readOnly}
          onChange={(event) => setReadOnly(event.target.checked)}
        />{" "}
        Read only
      </label>
      <WorkbookProvider workbook={workbook} readOnly={readOnly}>
        <div className="workspace">
          <section className="sheet">
            <Toolbar />
            <FormulaBar />
            <SelectionToolbar />
            <SheetGrid />
            <WorkbookTabs />
          </section>
          <PivotPreview />
        </div>
      </WorkbookProvider>
      <footer>
        Drag to select. Shift-click or Shift + arrow keys extend a selection.
        Drag header edges to resize. Double-click an edge to reset. The first
        100 visible rows and 26 visible columns are shown.
      </footer>
      <section className="chat-example" aria-label="Chat preview example">
        <div className="eyebrow">IN A CONVERSATION</div>
        <h2>Your workbook, inline.</h2>
        <div className="chat-message">
          <p>
            Here’s the workbook. Explore the sheets below, or open it to make
            changes.
          </p>
          <SpreadsheetPreview
            workbook={workbook}
            title={filename}
            theme="dark"
            onOpen={() =>
              document
                .querySelector(".sheet")
                ?.scrollIntoView({ behavior: "smooth", block: "start" })
            }
          />
        </div>
      </section>
    </main>
  );
}
const element = document.getElementById("root");
if (element)
  createRoot(element).render(<App initialWorkbook={initial.value} />);
