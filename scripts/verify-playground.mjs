import { execFileSync } from "node:child_process";
import { mkdirSync, writeFileSync, mkdtempSync, rmSync } from "node:fs";
import { resolve, join } from "node:path";
import { tmpdir } from "node:os";
import assert from "node:assert/strict";
const site = new URL(
  process.env.NORI_DOCS_URL ?? "http://127.0.0.1:4180/nori/",
);
const session = `nori-playground-${process.pid}`;
const temp = mkdtempSync(join(tmpdir(), "nori-sample-"));
const cli = (...args) =>
  execFileSync("agent-browser", ["--session", session, ...args], {
    encoding: "utf8",
    timeout: 30000,
  }).trim();
const evaluate = (expression) =>
  JSON.parse(JSON.parse(cli("eval", `JSON.stringify(${expression})`)));
try {
  const executable = process.env.NORI_BROWSER_EXECUTABLE;
  cli(
    ...(executable ? ["--executable-path", executable] : []),
    "open",
    new URL("playground.html", site).href,
  );
  cli("wait", "--load", "networkidle");
  cli("scrollintoview", "iframe");
  cli("wait", "1000");
  assert.equal(
    evaluate(
      'document.querySelector("iframe").contentDocument.querySelector("[data-cell-address=B2]").textContent',
    ),
    "1200",
  );
  assert.equal(
    evaluate(
      'document.querySelector("iframe").contentDocument.querySelector(".embedded") !== null',
    ),
    true,
  );
  // Select in the embedded document and verify the visible formula bar responds.
  evaluate(
    'document.querySelector("iframe").contentDocument.querySelector("[data-cell-address=B2]").click() ?? true',
  );
  assert.equal(
    evaluate(
      'document.querySelector("iframe").contentDocument.querySelector(".nori-formula-bar input").value',
    ),
    "1200",
  );
  mkdirSync("artifacts", { recursive: true });
  cli("screenshot", resolve("artifacts/live-playground.png"));
  cli("open", new URL("demo/", site).href);
  cli("wait", "--load", "networkidle");
  cli("dblclick", ".nori-grid [data-cell-address=B2]");
  cli("fill", '[aria-label="Edit B2"]', "100");
  cli("press", "Enter");
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid [data-cell-address=B6]").textContent',
    ),
    "1550",
  );
  cli("click", ".nori-tabs button:nth-child(2)");
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid [data-cell-address=B1]").textContent',
    ),
    "1550",
  );
  cli("click", ".toolbar button:first-child");
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid [data-cell-address=B1]").textContent',
    ),
    "2650",
  );
  cli("check", "input[type=checkbox]");
  assert.equal(
    evaluate('document.querySelectorAll(".nori-resize-handle").length'),
    0,
  );
  cli("click", ".reset-example");
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid [data-cell-address=B2]").textContent',
    ),
    "1200",
  );
  assert.equal(
    evaluate('document.querySelector("input[type=checkbox]").checked'),
    false,
  );
  const sample = await fetch(new URL("demo/sample.xlsx", site));
  assert.equal(sample.status, 200);
  const samplePath = join(temp, "sample.xlsx");
  writeFileSync(samplePath, Buffer.from(await sample.arrayBuffer()));
  cli("upload", "input[type=file]", samplePath);
  cli("wait", "1000");
  assert.ok(evaluate("document.body.innerText").includes("Opened sample.xlsx"));
  const csvResponse = await fetch(new URL("demo/sample.csv", site));
  assert.equal(csvResponse.status, 200);
  const csvPath = join(temp, "sample.csv");
  writeFileSync(csvPath, await csvResponse.text());
  cli("upload", "input[type=file]", csvPath);
  cli("wait", "1000");
  assert.ok(evaluate("document.body.innerText").includes("Opened sample.csv"));
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid [data-cell-address=C2]").textContent',
    ),
    "Quoted, comma",
  );
  cli("dblclick", ".nori-grid [data-cell-address=B1]");
  cli("fill", '[aria-label="Edit B1"]', "=SUM(B2:B4)");
  cli("press", "Enter");
  assert.equal(
    evaluate(
      'document.querySelector(".nori-grid [data-cell-address=B1]").textContent',
    ),
    "2650",
  );
  cli("set", "viewport", "390", "844");
  assert.ok(evaluate("document.documentElement.scrollWidth <= innerWidth"));
  cli("open", new URL("playground.html", site).href);
  cli("wait", "--load", "networkidle");
  assert.ok(evaluate("document.documentElement.scrollWidth <= innerWidth"));
  assert.equal(cli("errors"), "");
  console.log(
    "Live playground passed: embedded selection, editing, calculation, cross-sheet totals, undo, read-only, reset, hosted XLSX/CSV imports and mobile layout.",
  );
} finally {
  cli("close");
  rmSync(temp, { recursive: true, force: true });
}
