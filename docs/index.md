---
layout: home
hero:
  name: Nori
  text: Spreadsheets, on your terms.
  tagline: Import XLSX and CSV, calculate formulas, build pivots and compose React views. One public package. A platform-independent core.
  actions:
    - theme: brand
      text: Start building
      link: /getting-started
    - theme: alt
      text: Try it live
      link: /playground
    - theme: alt
      text: Build with an AI agent
      link: /agents
features:
  - title: A model you can keep
    details: Sparse, serializable workbook snapshots are independent of files, UI and live runtime state.
  - title: A runtime you can compose
    details: Commands, history, subscriptions, formulas and semantic pivots work without React or the DOM.
  - title: Views that fit your product
    details: Compose an editor, a read-only sheet or a compact chat preview. Keep document styles separate from your theme.
---

## Before you integrate

Nori is an initial 0.1.0 milestone, not a full Excel replacement. This repository has not published the package to npm. Use a local packed build today; the stable public import path is `@byfungsi/nori`.

Read the [support matrix](/support) before promising Excel fidelity. XLSX export, dynamic-array spilling, advanced filters, charts, and a React Native renderer are not implemented.

## Pick your entry point

| Task                                      | Start here                                    |
| ----------------------------------------- | --------------------------------------------- |
| Create a workbook and calculate a formula | [Quick start](/getting-started)               |
| Import an Excel attachment                | [XLSX import](/importing)                     |
| Embed an editable or read-only sheet      | [React integration](/react)                   |
| Render an agent's spreadsheet attachment  | [Chat preview](/preview)                      |
| Generate integration code with an agent   | [Agent guide](/agents), [llms.txt](/llms.txt) |

[Download all documentation as plain text](/llms-full.txt) · [Machine-readable page manifest](/docs-manifest.json)
