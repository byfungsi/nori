# Roadmap

1. **Compatibility corpus:** check in independently generated Excel/LibreOffice fixtures; expand malformed-file cases, number/date formats, formula coercion and error precedence. Report import losses more comprehensively.
2. **Runtime scale:** structural sharing, inverse-command history, compact range dependency indexing, incremental recalculation and benchmarks on sparse large workbooks. Scheduling/cancellation must remain host-supplied.
3. **Formula breadth:** SUMIF/COUNTIF, lookup and date functions, named ranges, structured references, reference rewriting for rename/copy/fill, spill ownership and collision errors, precise Excel semantics.
4. **Document editing:** sheet lifecycle, row/column insertion/deletion, richer styles, round-trip XLSX writing with explicit preservation guarantees.
5. **React editor:** expanded keyboard shortcuts, clipboard, virtualized rows/columns, auto-fit, interactive sort/filter editors, accessible editing refinements and native selection adapters.
6. **Pivots:** column axes, filtering, subtotals, sorting, richer measure semantics and XLSX pivot definition/cache adapters.
7. **React Native:** native renderer using the existing store, native file/byte adapter, platform text measurement, Metro/Hermes checks, touch selection and scheduling integration. No DOM primitives should be moved into core to enable this work.
8. **Release preparation:** choose licensing, add release/version automation, API stability policy, compatibility guarantees, generated API reference, and package-size budgets before a public stable release.
