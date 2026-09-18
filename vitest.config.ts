import { defineConfig } from "vitest/config";
import { resolve } from "node:path";
const names = ["model", "formula", "pivot", "core", "xlsx", "csv", "react"];
export default defineConfig({
  resolve: {
    alias: Object.fromEntries([
      ...names.map((name) => [
        "@nori-internal/" + name,
        resolve(
          "packages/" +
            name +
            "/src/index." +
            (name === "react" ? "tsx" : "ts"),
        ),
      ]),
      ...names.map((name) => [
        "@byfungsi/nori/" + name,
        resolve("packages/nori/src/" + name + ".ts"),
      ]),
      ["@byfungsi/nori", resolve("packages/nori/src/index.ts")],
    ]),
  },
  test: { include: ["tests/**/*.test.{ts,tsx}"] },
});
