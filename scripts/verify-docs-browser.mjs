import { execFileSync } from "node:child_process";
import { mkdirSync } from "node:fs";
import { resolve } from "node:path";
import assert from "node:assert/strict";
const site = new URL(process.env.NORI_DOCS_URL ?? "http://127.0.0.1:4179/");
const session = `nori-docs-${process.pid}`;
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
    site.href,
  );
  for (const path of [
    "/llms.txt",
    "/llms-full.txt",
    "/markdown/react.md",
    "/types/react.d.ts",
    "/docs-manifest.json",
  ]) {
    const response = await fetch(new URL(path.slice(1), site));
    assert.equal(response.status, 200);
    const content = await response.text();
    assert.ok(content.length > 100);
    assert.ok(
      !content.includes("<!DOCTYPE html>"),
      "Agent downloads must serve source, not an HTML fallback",
    );
  }
  cli("wait", "--load", "networkidle");
  assert.ok(
    evaluate("document.body.innerText").includes(
      "Spreadsheets, on your terms.",
    ),
  );
  cli("click", 'a[href*="getting-started"].VPButton');
  cli("wait", "h1");
  assert.ok(
    evaluate('document.querySelector("h1").textContent').includes(
      "Quick start",
    ),
  );
  cli("click", "button.DocSearch");
  cli("fill", "#localsearch-input", "readOnly");
  cli("wait", "1000");
  assert.ok(evaluate("document.body.innerText").includes("React integration"));
  cli("press", "Escape");
  cli("open", new URL("agents.html", site).href);
  cli("wait", "--load", "networkidle");
  assert.ok(
    evaluate('document.querySelector("main").textContent').includes(
      "Integration rules",
    ),
  );
  mkdirSync("artifacts", { recursive: true });
  cli("screenshot", resolve("artifacts/docs-desktop.png"));
  cli("set", "viewport", "390", "844");
  assert.ok(evaluate("document.documentElement.scrollWidth <= innerWidth"));
  cli("click", "button.VPNavBarHamburger");
  assert.ok(evaluate("document.body.innerText").includes("For agents"));
  cli("wait", "400");
  cli("screenshot", resolve("artifacts/docs-mobile.png"));
  assert.equal(cli("errors"), "");
  console.log(
    "Docs browser passed: home, navigation, local search, agent guide, mobile layout/menu, no application errors.",
  );
} finally {
  cli("close");
}
