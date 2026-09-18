import { describe, it, expect } from "vitest";
import fc from "fast-check";
import {
  addressOf,
  createEmptySnapshot,
  parseAddress,
  parseWorkbookSnapshot,
} from "@nori-internal/model";
import { unwrap } from "./helpers";
describe("canonical model", () => {
  it("roundtrips arbitrary Excel coordinates", () =>
    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: 1048575 }),
        fc.integer({ min: 0, max: 16383 }),
        (row, column) => {
          expect(unwrap(parseAddress(addressOf({ row, column })))).toEqual({
            row,
            column,
          });
        },
      ),
    ));
  it.each(["A0", "XFE1", "A1048577", "A-1", "1A", "A1junk", "A$$1"])(
    "rejects invalid address %s",
    (address) => expect(parseAddress(address).ok).toBe(false),
  );
  it("normalizes absolute addresses", () =>
    expect(unwrap(parseAddress("$a$12"))).toEqual({ row: 11, column: 0 }));
  it("is serializable and immutable", () => {
    const snapshot = createEmptySnapshot();
    expect(
      unwrap(parseWorkbookSnapshot(JSON.parse(JSON.stringify(snapshot)))),
    ).toEqual(snapshot);
    expect(Object.isFrozen(snapshot.sheets[0]?.cells)).toBe(true);
  });
  it("rejects duplicate identity and sheet names ignoring case", () => {
    const s = createEmptySnapshot();
    const sheet = s.sheets[0];
    expect(parseWorkbookSnapshot({ ...s, sheets: [sheet, sheet] }).ok).toBe(
      false,
    );
    expect(
      parseWorkbookSnapshot({
        ...s,
        sheets: [sheet, { ...sheet, id: "other", name: "sheet1" }],
      }).ok,
    ).toBe(false);
  });
  it("rejects out-of-extent cells, invalid versions and NaN", () => {
    const s = createEmptySnapshot();
    expect(parseWorkbookSnapshot({ ...s, version: 2 }).ok).toBe(false);
    expect(
      parseWorkbookSnapshot({
        ...s,
        sheets: [{ ...s.sheets[0], cells: { A101: { value: 1 } } }],
      }).ok,
    ).toBe(false);
    expect(
      parseWorkbookSnapshot({
        ...s,
        sheets: [{ ...s.sheets[0], cells: { A1: { value: NaN } } }],
      }).ok,
    ).toBe(false);
  });
});
