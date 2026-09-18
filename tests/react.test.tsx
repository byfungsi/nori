// @vitest-environment jsdom
import { afterEach, expect, it } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import {
  Spreadsheet,
  WorkbookProvider,
  WorkbookTabs,
  SheetGrid,
  FormulaBar,
} from "@nori-internal/react";
import { workbook } from "./helpers";
afterEach(cleanup);
it("renders imported sheets, switches tabs, edits and recalculates", () => {
  const book = workbook();
  render(<Spreadsheet workbook={book} />);
  expect(screen.getByRole("table", { name: "Sales spreadsheet" })).toBeTruthy();
  fireEvent.click(screen.getByRole("button", { name: "B2: 1200" }));
  fireEvent.change(
    screen.getByRole("textbox", { name: "Cell value or formula" }),
    { target: { value: "100" } },
  );
  fireEvent.click(screen.getByRole("button", { name: "Apply" }));
  expect(screen.getByRole("button", { name: "B6: 1550.00" })).toBeTruthy();
  fireEvent.click(screen.getByRole("button", { name: "Summary" }));
  expect(
    screen.getByRole("table", { name: "Summary spreadsheet" }),
  ).toBeTruthy();
  expect(screen.getByRole("button", { name: "B1: 1550" })).toBeTruthy();
});
it("composes primitives and custom content while preserving workbook style", () => {
  const book = workbook();
  render(
    <WorkbookProvider workbook={book}>
      <WorkbookTabs className="custom-tabs" />
      <FormulaBar />
      <SheetGrid
        renderCell={({ address, value }) => (
          <span>
            {address === "B2" ? "Custom revenue" : String(value ?? "")}
          </span>
        )}
      />
    </WorkbookProvider>,
  );
  expect(screen.getByText("Custom revenue")).toBeTruthy();
  const total = screen.getByRole("button", { name: "B6: 2650.00" });
  expect(total.closest("td")?.style.fontWeight).toBe("bold");
  expect(screen.getByRole("navigation").className).toBe("custom-tabs");
});
it("reports rejected edits and preserves the previous cell", () => {
  const book = workbook();
  render(<Spreadsheet workbook={book} />);
  fireEvent.click(screen.getByRole("button", { name: "B2: 1200" }));
  fireEvent.change(screen.getByRole("textbox"), { target: { value: "=" } });
  fireEvent.click(screen.getByRole("button", { name: "Apply" }));
  expect(screen.getByRole("alert")).toBeTruthy();
  expect(book.getCellValue(book.getState().activeSheetId, "B2")).toBe(1200);
});
