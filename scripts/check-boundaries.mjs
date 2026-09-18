import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import ts from "typescript";
const allowed = {
  model: [],
  formula: ["model"],
  pivot: ["model"],
  core: ["model", "formula"],
  xlsx: ["model"],
  react: ["model", "core", "formula"],
  nori: ["model", "core", "formula", "xlsx", "react", "pivot"],
};
function files(path) {
  return readdirSync(path, { withFileTypes: true }).flatMap((entry) =>
    entry.isDirectory()
      ? files(join(path, entry.name))
      : [join(path, entry.name)],
  );
}
const errors = [];
for (const [name, dependencies] of Object.entries(allowed)) {
  const manifest = JSON.parse(
    readFileSync(`packages/${name}/package.json`, "utf8"),
  );
  for (const path of files(`packages/${name}/src`)) {
    const source = ts.createSourceFile(
      path,
      readFileSync(path, "utf8"),
      ts.ScriptTarget.Latest,
      true,
      path.endsWith("tsx") ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
    );
    function visit(node) {
      if (
        (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) &&
        node.moduleSpecifier &&
        ts.isStringLiteral(node.moduleSpecifier)
      ) {
        const specifier = node.moduleSpecifier.text;
        if (
          specifier.startsWith("@nori-internal/") &&
          !dependencies.includes(specifier.slice("@nori-internal/".length))
        )
          errors.push(`${path}: forbidden dependency ${specifier}`);
        if (name !== "nori" && specifier.startsWith("@byfungsi/nori"))
          errors.push(`${path}: internal package depends on facade`);
        if (specifier.startsWith("../") && specifier.includes("/src"))
          errors.push(`${path}: relative cross-package dependency`);
        const packageName = specifier.startsWith("@")
          ? specifier.split("/").slice(0, 2).join("/")
          : specifier.split("/")[0];
        if (
          !specifier.startsWith(".") &&
          name !== "nori" &&
          !manifest.dependencies?.[packageName] &&
          !manifest.peerDependencies?.[packageName]
        )
          errors.push(`${path}: undeclared dependency ${packageName}`);
      }
      if (
        !["react", "nori"].includes(name) &&
        ts.isIdentifier(node) &&
        [
          "window",
          "document",
          "FileReader",
          "WebWorker",
          "Worker",
          "HTMLElement",
          "localStorage",
        ].includes(node.text)
      )
        errors.push(`${path}: platform global ${node.text}`);
      ts.forEachChild(node, visit);
    }
    visit(source);
  }
}
if (errors.length) throw Error(errors.join("\n"));
console.log("Package boundaries and headless platform checks passed.");
