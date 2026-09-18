import { defineConfig } from "tsup";
export default defineConfig({
  entry: [
    "src/index.ts",
    "src/model.ts",
    "src/xlsx.ts",
    "src/csv.ts",
    "src/core.ts",
    "src/formula.ts",
    "src/pivot.ts",
    "src/react.ts",
  ],
  format: ["esm"],
  target: "es2022",
  dts: true,
  splitting: true,
  sourcemap: true,
  clean: true,
  treeshake: true,
  noExternal: [/^@nori-internal\//],
  external: ["react", "react/jsx-runtime"],
  tsconfig: "../../tsconfig.json",
});
