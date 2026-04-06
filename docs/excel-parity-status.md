# Excel Parity Status

This document tracks the current Excel parity state of `officekit` relative to OfficeCLI.

## Summary

| Area | Status | Notes |
| --- | --- | --- |
| Core workbook/sheet/cell flows | Done | Create, add, set, remove, get, query, view, raw, import, and metadata-free OOXML mutation are in place. |
| Advanced Excel object paths | Mostly done | Named ranges, validations, conditional formatting, comments, tables, sparklines, charts, pivots, shapes, and pictures all have working adapter paths. |
| Deep parity engines | Partial | Formula evaluation, chart property coverage, and style management are improved but not yet fully equivalent to OfficeCLI. |
| Verification | Strong | End-to-end CLI coverage is in place and fixture-backed tests are passing. |

## Detailed Status

| Module | Current status | Already implemented | Remaining gaps | Priority |
| --- | --- | --- | --- | --- |
| Workbook and sheet model | Mostly done | Workbook settings, sheet creation, rename/update flows, freeze panes, zoom, headings, gridlines, tab color, print/header/footer, row/column breaks, protection basics | Additional long-tail workbook/sheet flags and more exact OfficeCLI property semantics | Medium |
| Cell and range operations | Mostly done | Cell/range set/get/remove, number/string/boolean/date handling, authored formulas, import type inference, style id preservation | More exact Excel coercion semantics, more precise writeback for complex mixed-type edits | Medium |
| Raw OOXML access | Mostly done | `/workbook`, `/styles`, `/sharedstrings`, sheet raw, drawing raw, chart raw, filtered raw sheet output | More niche raw part coverage and more exact filtered raw parity behavior | Medium |
| Query and view surface | Mostly done | `outline`, `text`, `annotated`, `stats`, `issues`, `html`, `json`; object queries for many Excel node families | More OfficeCLI-like selector/filter semantics and deeper query detail for some object families | Medium |
| Named ranges | Done | Add/get/set/remove/query/path support | More edge-case behavior around workbook mutations that shift references | Low |
| Validations | Mostly done | Add/get/set/remove/query, XML round-trip, prompt/error/basic rule fields | Broader validation rule surface and richer parameter parity | Medium |
| Conditional formatting | Partial | Add/get/set/remove/query for `databar`, `colorscale`, `iconset`, `formula`, `topn`, `aboveaverage`, `uniquevalues`, `duplicatevalues`, `containstext`, `dateoccurring`; basic differential formats | More complete CF rule options, richer dxf support, closer OOXML parity for advanced rules | High |
| Comments | Mostly done | Add/get/set/remove/query, authors, OOXML part creation | More exact legacy comment shape/anchor behavior if OfficeCLI depends on it | Medium |
| Tables | Mostly done | Add/get/set/remove/query, table rels, style name, totals/header flags, OOXML table creation | More table property depth and better column metadata parity | Medium |
| Sparklines | Mostly done | Add/get/set/remove/query/path support, type/location/source range updates | More sparkline group styling and long-tail sparkline properties | Medium |
| Charts object model | Partial | Add/get/set/remove/query/raw, title, legend, data labels, axis titles, axis min/max, axis units, axis number format, chart/plot fills, series name, series colors, style id | More chart types, more internal chart XML shapes, more exact series/axis behavior across real-world files | High |
| Chart property parity (`ChartHelper.SetChartProperties`) | Partial | High-frequency properties are present and tested | Missing broader matrix: title/legend/label fonts, line width/dash, marker style/size, gridline style, transparency, gradient, effects, more series-level controls, preset/theme behavior | Very high |
| Pivot tables | Partial | Add/get/set/remove/query for path-level objects, basic pivot metadata like name and selected booleans | Row/column/data/filter field semantics, deeper pivot definition properties, cache-related behavior | High |
| Shapes and pictures | Partial | Add/get/set/remove/query/raw, drawing creation, anchor positioning, basic name/text/alt updates | Richer text styling, fills, line style, rotation, sizing nuances, broader mixed-drawing compatibility | High |
| Formula evaluator (`FormulaEvaluator`) | Partial | Display-time evaluation for selected high-frequency functions and some cross-sheet references; current support includes `SUM`, `AVERAGE`, `MIN`, `MAX`, `IF`, `COUNTA`, `SUMPRODUCT` | Large remaining function matrix: `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`, `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, more text/date/time/logical/statistical functions, more nested-expression support, richer Excel-style coercion/error propagation, more robust circular/lookup semantics | Very high |
| Style manager (`ExcelStyleManager`) | Partial | Generated styles for common `font.*`, `fill`, `numFmt`, and `alignment.*`; style reuse for equivalent style props; styles.xml persisted and reused | Broader style surface: border families, protection flags, more alignment fields, more font/fill combinations, better inheritance/merge against existing styles, stronger dedupe and style preservation semantics | Very high |
| Metadata-free OOXML compatibility | Mostly done | Real external OOXML files can be mutated without losing many existing workbook/style/formula features | More pressure-testing across complex third-party workbooks mixing chart + pivot + drawing + CF + style + formula behavior | High |
| End-to-end regression coverage | Strong | Fixture-backed CLI tests cover a wide Excel surface and all tests are currently green | More real-world workbook fixture diversity and more cross-feature stress tests | Medium |

## Highest-Value Remaining Work

| Rank | Focus area | Why it matters most | Concrete next targets |
| --- | --- | --- | --- |
| 1 | Formula evaluator | Directly affects `get`, `view`, and user trust in Excel read behavior | `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`, `SUMIF`, `COUNTIF`, `AVERAGEIF`, plus more nested cross-sheet evaluation |
| 2 | Chart property parity | Largest remaining gap in advanced Excel editing parity | Marker/line/grid/title/legend font/effects/transparency/gradient/preset/theme and more per-series controls |
| 3 | Style manager depth | Needed for reliable formatting parity, especially on real-world workbooks | Borders, protection, more alignment/font/fill variants, inheritance and merge logic against existing styles |
| 4 | Pivot deeper semantics | Important for real business workbook parity | Row/column/data field behavior, summary/layout toggles, broader pivot definition read/write coverage |
| 5 | Mixed OOXML compatibility | Final confidence layer before claiming near-complete parity | More fixtures combining CF + chart + pivot + drawing + style + formulas in the same workbook |

## Suggested Interpretation

| Label | Meaning |
| --- | --- |
| Done | Feature family is broadly present and not the main blocker to OfficeCLI parity. |
| Mostly done | Main workflows are present, but there are still important edge cases or long-tail fields missing. |
| Partial | Significant parity progress exists, but the module is still one of the main reasons Excel is not fully replicated yet. |
