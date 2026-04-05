# Current implementation status

This report is the lane-4 parity snapshot for the current `officekit` checkpoint. It is intentionally conservative: it records what is **explicitly evidenced today** by package manifests, fixture inventory, and runnable tests, and separates that from work that is still pending.

## Status legend

- `scaffolded` — source lineage, package contracts, and parity expectations are encoded in code/tests/docs.
- `implemented` — executable document behavior exists for that slice.
- `verified` — runnable checks prove the implemented behavior against fixtures or differential expectations.

## Word / Excel / PowerPoint slice status

| Format | Current status | Fixture-backed evidence | Explicitly supported today | Remaining gaps |
| --- | --- | --- | --- | --- |
| Word | scaffolded | `word-formulas-script`, `word-tables-script`, `word-textbox-script`, `word-complex-formulas-output`, `word-complex-tables-output` | Manifested public surface for add/get/query/set/remove/move/swap/raw/batch + preview modes `html/forms`; canonical paths and parity risks are under test | Document traversal, mutation behavior, section/style fidelity, and HTML/forms rendering still need real implementation and verification |
| Excel | scaffolded | `excel-beautiful-charts-script`, `excel-charts-demo-script`, `excel-sales-report-output`, `excel-charts-demo-output`, `excel-beautiful-charts-output` | Manifested workbook/chart/import/formula/style/pivot scope with canonical workbook/chart paths under test | Workbook editing, formula execution, style fidelity, chart/pivot implementation, filtered raw-sheet views, and preview output remain to be built and verified |
| PowerPoint | scaffolded | `ppt-beautiful-script`, `ppt-animations-script`, `ppt-video-script`, `ppt-3d-script`, `ppt-beautiful-output`, `ppt-data-output`, `ppt-animations-output`, referenced `ppt-3d-model-asset` | Manifested slide/layout/placeholder/theme/media/animation scope with preview modes `html/svg` and overflow-check contract under test | Slide mutation/layout resolution, theme/text fidelity, preview rendering fidelity, media/3D behavior, and fixture replay verification remain pending |

## What is verified by lane-4 tests right now

1. The harvested fixture manifest still contains canonical Word, Excel, and PowerPoint scenarios.
2. Each format package exposes parity-critical manifest/contract metadata that matches the fixture corpus.
3. Documentation keeps the OfficeCLI lineage explicit and calls out scaffolded vs future implemented/verified states.

## Remaining gaps

- The format packages are still metadata-first scaffolds, not document-manipulation implementations.
- Fixture-backed evidence currently proves **coverage/reporting alignment**, not end-user Office document mutation fidelity.
- Differential document output checks, preview rendering comparisons, and live watch/browser flows still need format-aware executable implementations.
