import { readFileSync } from "node:fs";
import { createWorkbook, type Workbook } from "@nori-internal/core";
import { parseXlsx } from "@nori-internal/xlsx";
import type { Result } from "@nori-internal/model";
/** Test boundary fails with the original typed error. */
export function unwrap<T, E extends Error>(result: Result<T, E>): T {
  if (!result.ok) throw result.error;
  return result.value;
}
/** A real two-sheet ZIP/OOXML document, also downloadable from the demo. */
export const basicBytes = new Uint8Array(
  readFileSync("tests/fixtures/basic.xlsx"),
);
/** Rehydrate the same adapter/runtime flow used in the demo. */
export function workbook(): Workbook {
  return unwrap(createWorkbook(unwrap(parseXlsx(basicBytes)).snapshot));
}
