# Try Nori live

Edit a real workbook below—no installation or account needed. This example uses the same public `@byfungsi/nori` package and React components described in the guides.

## Try these actions

1. Double-click **Sales B2**, change `1200` to `100`, and press Enter. The total in **B6** becomes **1550**; the pivot updates too.
2. Switch to **Summary** to see the cross-sheet total, then choose **Undo** to restore the original values.
3. Drag across cells, Shift-click to extend selection, or drag a row/column boundary to resize it.
4. Turn on **Read only** to try a navigable sheet without editing.
5. Download a sample, then open it with **Open .xlsx**. The layout sample includes merged cells, frozen panes, hidden columns and saved filters.

<LivePlayground />

## Your files stay in this browser

The example parses selected files locally. It does not upload workbook contents, save edits to a server, or persist them between page reloads. Choose **Reset example** to restore the starting workbook. The sample files contain synthetic data.

This is a bounded example, not complete Excel compatibility: the editor renders at most 100 visible rows and 26 visible columns, and XLSX export is not available. Import notes identify supported warnings. Read the [support matrix](/support) before relying on document fidelity.

## Build your own

Start with the [React integration guide](/react) or the [headless quick start](/getting-started). The complete runnable example lives in [examples/react-demo](https://github.com/byfungsi/nori/tree/main/examples/react-demo). Its chat preview, sheet tabs and pivot share one live workbook runtime.
