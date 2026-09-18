import assert from "node:assert/strict";
import { readFileSync, existsSync, readdirSync } from "node:fs";
import { resolve } from "node:path";
import { createHash } from "node:crypto";
const root = resolve("docs/.vitepress/dist");
const read = (path) => readFileSync(resolve(root, path), "utf8");
const manifest = JSON.parse(read("docs-manifest.json"));
assert.equal(
  manifest.version,
  JSON.parse(readFileSync("packages/nori/package.json", "utf8")).version,
);
assert.ok(manifest.pages.length >= 15);
const full = read("llms-full.txt"),
  index = read("llms.txt");
for (const page of manifest.pages) {
  assert.ok(page.markdown.startsWith(manifest.base));
  const content = read(page.markdown.slice(manifest.base.length));
  assert.equal(createHash("sha256").update(content).digest("hex"), page.sha256);
  assert.ok(full.includes(content));
  assert.ok(index.includes(page.markdown));
  assert.ok(!/^<<< /m.test(content), "Code includes must be expanded");
  assert.ok(
    existsSync(
      resolve(
        root,
        page.route === manifest.base
          ? "index.html"
          : page.route
              .slice(manifest.base.length)
              .replace(/(?:\.html)?$/, ".html"),
      ),
    ),
  );
}
for (const entry of [
  "index",
  "react",
  "core",
  "model",
  "formula",
  "pivot",
  "xlsx",
  "csv",
])
  assert.ok(read("types/" + entry + ".d.ts").length > 20);
for (const name of readdirSync(resolve(root, "types"))) {
  for (const match of read("types/" + name).matchAll(
    /from ['"]\.\/([^'"]+)['"]/g,
  ))
    assert.ok(
      existsSync(resolve(root, "types", match[1].replace(/\.js$/, ".d.ts"))),
      `Missing declaration dependency ${match[1]}`,
    );
}
assert.ok(read("react.html").includes("readOnly"));
assert.ok(read("agents.html").includes(`${manifest.base}llms-full.txt`));
assert.ok(read("index.html").includes(`${manifest.base}assets/`));
console.log(
  `Documentation passed: ${manifest.pages.length} HTML/Markdown pages, content hashes, complete text and declaration dependency closure.`,
);

assert.ok(read("demo/index.html").includes("./assets/"));
assert.ok(read("playground.html").includes(`${manifest.base}demo/?embed=1`));
for (const fixture of ["sample.xlsx", "layout.xlsx", "sample.csv"])
  assert.ok(existsSync(resolve(root, "demo", fixture)));
