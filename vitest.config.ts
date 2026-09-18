import { defineConfig } from "vitest/config";
import { resolve } from "node:path";
const names = ["model", "formula", "pivot", "core", "xlsx", "react"];
export default defineConfig({
  resolve: {
    alias: Object.fromEntries(
      names.map((name) => [
        "@nori-internal/" + name,
        resolve(
          "packages/" +
            name +
            "/src/index." +
            (name === "react" ? "tsx" : "ts"),
        ),
      ]),
    ),
  },
  test: { include: ["tests/**/*.test.{ts,tsx}"] },
});
