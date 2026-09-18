import { execFileSync } from "node:child_process";
import { mkdirSync } from "node:fs";
import { resolve } from "node:path";
import assert from "node:assert/strict";
const session = `nori-check-${process.pid}`;
const cli = (...args) =>
  execFileSync("agent-browser", ["--session", session, ...args], {
    encoding: "utf8",
    timeout: 30000,
  }).trim();
const evaluate = (expression) =>
  JSON.parse(JSON.parse(cli("eval", `JSON.stringify(${expression})`)));
const box = (selector) =>
  evaluate(
    `(()=>{const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2,width:r.width,height:r.height,left:r.left,top:r.top}})()`,
  );
const drag = (from, to) => {
  cli("mouse", "move", String(Math.round(from.x)), String(Math.round(from.y)));
  cli("mouse", "down");
  cli("mouse", "move", String(Math.round(to.x)), String(Math.round(to.y)));
  cli("mouse", "up");
};
const selected = () =>
  evaluate(
    'document.querySelector(".nori-selection-toolbar output").textContent',
  );
const screenshot = (name) => cli("screenshot", resolve("artifacts", name));
mkdirSync("artifacts", { recursive: true });
try {
  const executable = process.env["NORI_BROWSER_EXECUTABLE"];
  cli(
    ...(executable ? ["--executable-path", executable] : []),
    "open",
    process.env["NORI_DEMO_URL"] ?? "http://127.0.0.1:4178/",
  );
  cli("wait", "--load", "networkidle");
  cli("snapshot", "-i");
  evaluate(
    'document.querySelector(".sheet").scrollIntoView({block:"start"})??true',
  );
  cli("dblclick", '.nori-grid [data-cell-address="B2"]');
  screenshot("inline-editor.png");
  cli("fill", '[aria-label="Edit B2"]', "=100+25");
  cli("press", "Enter");
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid [data-cell-address=B2]").textContent',
    ),
    "125",
  );
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid [data-cell-address=B6]").textContent',
    ),
    "1575",
  );
  cli("click", ".toolbar button:first-child");
  cli("dblclick", '.nori-grid [data-cell-address="B2"]');
  cli("fill", '[aria-label="Edit B2"]', "9999");
  cli("press", "Escape");
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid [data-cell-address=B2]").textContent',
    ),
    "1200",
  );
  cli("check", 'input[type="checkbox"]');
  cli("dblclick", '.nori-grid [data-cell-address="B2"]');
  assert.equal(evaluate('!!document.querySelector(".nori-grid input")'), false);
  assert.equal(
    evaluate('document.querySelectorAll(".nori-resize-handle").length'),
    0,
  );
  assert.equal(
    evaluate('document.querySelector(".nori-formula-bar input").readOnly'),
    true,
  );
  cli("uncheck", 'input[type="checkbox"]');
  evaluate(
    'document.querySelector(".sheet").scrollIntoView({block:"start"})??true',
  );
  drag(
    box('.nori-grid [data-cell-address="C2"]'),
    box('.nori-grid [data-cell-address="E4"]'),
  );
  assert.equal(selected(), "C2:E4");
  assert.equal(
    evaluate(
      'document.querySelectorAll(".nori-grid td[data-selected]").length',
    ),
    9,
  );
  // A real keyboard chord extends from the saved drag focus.
  cli("press", "Shift+ArrowRight");
  assert.equal(selected(), "C2:F4");
  evaluate(
    'document.querySelector(".nori-grid [data-cell-address=\\"D3\\"]").dispatchEvent(new MouseEvent("click",{bubbles:true,shiftKey:true,detail:0}))',
  );
  assert.equal(selected(), "C2:D3");
  cli("click", ".nori-selection-toolbar button:first-of-type");
  assert.deepEqual(
    evaluate(
      '(()=>{const c=document.querySelector(".nori-grid td[data-address=\\"C2\\"]");return [c.rowSpan,c.colSpan]})()',
    ),
    [2, 2],
  );
  cli("click", ".nori-selection-toolbar button:nth-of-type(2)");
  assert.equal(
    evaluate(
      '!!document.querySelector(".nori-grid [data-cell-address=\\"D3\\"]")',
    ),
    true,
  );
  const column = '[aria-label="Resize column D"]',
    widthBefore = box('.nori-grid th[data-column-index="3"]').width;
  const columnPoint = box(column);
  drag(columnPoint, { x: columnPoint.x + 70, y: columnPoint.y });
  assert.equal(
    Math.round(box('.nori-grid th[data-column-index="3"]').width - widthBefore),
    70,
  );
  cli("click", ".toolbar button:first-child");
  assert.equal(box('.nori-grid th[data-column-index="3"]').width, widthBefore);
  const row = '[aria-label="Resize row 3"]',
    heightBefore = box('.nori-grid tr[data-row-index="2"]').height,
    rowPoint = box(row);
  drag(rowPoint, { x: rowPoint.x, y: rowPoint.y + 26 });
  assert.equal(
    Math.round(box('.nori-grid tr[data-row-index="2"]').height - heightBefore),
    26,
  );
  // Body edges use the same resize gesture and full-viewport guide as headers.
  const body = box('.nori-grid td[data-address="D4"]');
  const boundary = { x: body.left + body.width, y: body.y };
  for (const delta of [-4, 4]) {
    cli(
      "mouse",
      "move",
      String(Math.round(boundary.x + delta)),
      String(Math.round(boundary.y)),
    );
    assert.equal(
      evaluate('document.querySelector(".nori-resize-guide").dataset.index'),
      "3",
    );
    assert.equal(
      evaluate('document.querySelector(".nori-resize-guide").dataset.axis'),
      "column",
    );
  }
  assert.ok(
    Math.abs(box(".nori-resize-guide").height - box(".nori-scroll").height) < 2,
  );
  screenshot("resize-border.png");
  drag(
    { x: boundary.x + 4, y: boundary.y },
    { x: boundary.x + 44, y: boundary.y },
  );
  assert.equal(
    Math.round(box('.nori-grid th[data-column-index="3"]').width - widthBefore),
    40,
  );
  cli("click", ".toolbar button:first-child");
  const bodyRow = box('.nori-grid td[data-address="D4"]');
  const rowBefore = box('.nori-grid tr[data-row-index="3"]').height;
  drag(
    { x: bodyRow.x, y: bodyRow.top + bodyRow.height - 3 },
    { x: bodyRow.x, y: bodyRow.top + bodyRow.height + 17 },
  );
  assert.equal(
    Math.round(box('.nori-grid tr[data-row-index="3"]').height - rowBefore),
    20,
  );
  cli("click", ".toolbar button:first-child");
  screenshot("editor.png");
  // Check drag auto-scroll while the pointer remains at the scrollport edge.
  const start = box('.nori-grid [data-cell-address="E4"]'),
    scroll = box(".nori-scroll");
  cli(
    "mouse",
    "move",
    String(Math.round(start.x)),
    String(Math.round(start.y)),
  );
  cli("mouse", "down");
  cli(
    "mouse",
    "move",
    String(Math.round(scroll.left + scroll.width - 40)),
    String(Math.round(scroll.top + scroll.height - 3)),
  );
  cli("wait", "350");
  cli("mouse", "up");
  assert.ok(evaluate('document.querySelector(".nori-scroll").scrollTop') > 0);
  cli("upload", "input[type=file]", resolve("tests/fixtures/layout.xlsx"));
  cli("wait", ".nori-view-status");
  evaluate(
    'document.querySelector(".sheet").scrollIntoView({block:"start"})??true',
  );
  assert.equal(
    evaluate(
      '!!document.querySelector(".nori-grid [data-column-index=\\"2\\"]")',
    ),
    false,
  );
  assert.equal(
    evaluate(
      '!!document.querySelector(".nori-grid tr[data-row-index=\\"5\\"]")',
    ),
    false,
  );
  assert.equal(
    evaluate(
      '!!document.querySelector(".nori-grid tr[data-row-index=\\"7\\"]")',
    ),
    false,
  );
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid th[data-column-index=\\"1\\"]").getAttribute("aria-sort")',
    ),
    "descending",
  );
  assert.equal(
    Math.round(box('.nori-grid th[data-column-index="3"]').width),
    336,
  );
  const frozenBefore = box('.nori-grid td[data-address="A4"]');
  evaluate(
    '(()=>{const s=document.querySelector(".nori-scroll");s.scrollLeft=400;s.scrollTop=160;return true})()',
  );
  const frozenAfter = box('.nori-grid td[data-address="A4"]');
  assert.ok(Math.abs(frozenAfter.left - frozenBefore.left) < 2);
  assert.ok(Math.abs(frozenAfter.top - frozenBefore.top) < 2);
  screenshot("imported-layout.png");
  // Preview has its own read-only tabs and fits a narrow chat column.
  cli("set", "viewport", "390", "844");
  evaluate(
    'document.querySelector(".chat-example").scrollIntoView({block:"start"})??true',
  );
  assert.equal(
    evaluate("document.documentElement.scrollWidth<=innerWidth"),
    true,
  );
  assert.equal(
    evaluate(
      'document.querySelectorAll(".nori-preview input, .nori-preview [role=separator], .nori-preview [data-cell-address]").length',
    ),
    0,
  );
  assert.equal(
    evaluate(
      'document.querySelector(".nori-preview").getBoundingClientRect().width<=innerWidth',
    ),
    true,
  );
  screenshot("chat-preview-mobile.png");
  const errors = cli("errors");
  assert.equal(errors, "");
  assert.equal(
    evaluate('!!document.querySelector("vite-error-overlay")'),
    false,
  );
  console.log(
    "Browser passed: double-click editing + formulas + Escape + readOnly, pointer drag + auto-scroll, Shift selection, merge/unmerge, row/column resize + undo, imported geometry/filter/hidden/freeze/sort, responsive read-only preview.",
  );
} catch (error) {
  try {
    screenshot("failure.png");
    console.error(cli("errors"));
  } catch {}
  throw error;
} finally {
  cli("close");
}
