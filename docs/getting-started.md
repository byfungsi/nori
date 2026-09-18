# Quick start

## Install a local package

This repository has not published Nori to npm. From the repository root:

```sh
npm ci
npm run build
npm pack --workspace @byfungsi/nori
```

In your application's directory, install the generated archive (adjust its path):

```sh
npm install /path/to/nori/byfungsi-nori-0.1.0.tgz
```

For the React renderer, also install compatible React peers:

```sh
npm install react@^19 react-dom@^19
```

Headless use does not require React. The package is ESM; use an ESM application or a compatible bundler. Repository development requires Node 22.12 or later. Native bindings are future work.

## Calculate and persist a workbook

This is a complete, typechecked example. Expected failures are Result values; this short program chooses to throw after checking them. An application should display or return these errors at its boundary.

<<< ./examples/quick-start.ts

`WorkbookSnapshot` is immutable JSON-compatible document data. `Workbook` is a live runtime with commands, subscriptions, history and calculated values. Persist the snapshot, never the runtime object. Exported cell values may be old formula caches; read calculated results through `getCellValue`.

## Choose an import

| Import                   | Purpose                                                           |
| ------------------------ | ----------------------------------------------------------------- |
| `@byfungsi/nori`         | Convenient model constructors, XLSX parser and runtime            |
| `@byfungsi/nori/model`   | Canonical types, validation, addresses, ranges and Result helpers |
| `@byfungsi/nori/core`    | Runtime, commands, selection and layout                           |
| `@byfungsi/nori/csv`     | Text CSV to a single-sheet snapshot                               |
| `@byfungsi/nori/xlsx`    | Byte-to-snapshot adapter                                          |
| `@byfungsi/nori/formula` | AST parser, evaluator, registry and dependency graph              |
| `@byfungsi/nori/pivot`   | Semantic grouping and aggregation                                 |
| `@byfungsi/nori/react`   | React DOM components and hooks                                    |

Only `/react` imports React. Prefer `/core` when you do not need XLSX dependencies. Internal workspace names are implementation details and must not appear in consumer code.

Continue with [React](/react), [XLSX import](/importing), or [headless recipes](/recipes).
