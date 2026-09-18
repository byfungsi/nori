import { execFileSync } from "node:child_process";
import { cpSync, rmSync } from "node:fs";
// Relative asset URLs let the same demo run at /demo/ or /nori/demo/.
execFileSync(
  "npm",
  ["run", "build", "--workspace", "@nori/demo", "--", "--base", "./"],
  { stdio: "inherit" },
);
rmSync("docs/public/demo", { recursive: true, force: true });
cpSync("examples/react-demo/dist", "docs/public/demo", {
  recursive: true,
});
