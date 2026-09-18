import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

const manifest = JSON.parse(readFileSync("packages/nori/package.json", "utf8"));
assert.equal(manifest.name, "@byfungsi/nori");
assert.equal(manifest.publishConfig.access, "public");
assert.ok(!manifest.private);
if (process.env.RELEASE_VERSION) {
  assert.equal(
    manifest.version,
    process.env.RELEASE_VERSION,
    "Release version must match package.json",
  );
}
if (process.argv.includes("--require-license")) {
  assert.ok(
    manifest.license,
    "Choose a license before publishing and add it to packages/nori/package.json",
  );
  if (manifest.license !== "UNLICENSED") {
    assert.ok(
      readFileSync("packages/nori/LICENSE", "utf8").trim(),
      "Include the chosen license in the package",
    );
  }
}
const run = (args) => execFileSync("npm", args, { stdio: "inherit" });
run(["run", "verify"]);
const destination = resolve("artifacts/release");
mkdirSync(destination, { recursive: true });
const [packed] = JSON.parse(
  execFileSync(
    "npm",
    [
      "pack",
      "--workspace",
      manifest.name,
      "--pack-destination",
      destination,
      "--json",
    ],
    { encoding: "utf8" },
  ),
);
assert.equal(packed.name, manifest.name);
assert.equal(packed.version, manifest.version);
const paths = new Set(packed.files.map((file) => file.path));
for (const entry of Object.values(manifest.exports)) {
  for (const path of Object.values(entry))
    assert.ok(paths.has(path.replace(/^\.\//, "")), `Missing export: ${path}`);
}
for (const path of paths)
  assert.ok(
    /^(dist\/|package\.json$|README\.md$|LICENSE$)/.test(path),
    `Unexpected package file: ${path}`,
  );
const archive = resolve(destination, packed.filename);
run([
  "publish",
  archive,
  "--dry-run",
  "--access",
  "public",
  "--registry",
  "https://registry.npmjs.org/",
]);
writeFileSync(
  resolve(destination, "manifest.json"),
  JSON.stringify(packed, null, 2) + "\n",
);
console.log(
  `Release archive ready: ${archive}\nIntegrity: ${packed.integrity}\nNothing was published.`,
);
if (!manifest.license)
  console.warn(
    "License decision pending: publishing workflow will refuse publication until it is set.",
  );
