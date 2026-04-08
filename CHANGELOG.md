# Changelog

All notable changes to officekit will be documented in this file.

## [Unreleased]

### Changed

- **Capability status alignment**: Updated `packages/core/src/parity.ts` to mark format-specific capabilities (create, add, set, get, query, remove, view, check, watch) as `implemented` instead of `scaffolded`, reflecting that Word, Excel, and PowerPoint format packages have reached implementation status per `docs/parity/implementation-status.md`.

- **Router implementation status**: Updated `packages/core/src/router.ts` to use `implementationStatus: "implemented"` for document commands routed to format packages. Added `"implemented"` as a valid status option in the `ExecutionPlan` interface.

- **CLI help text**: Updated `renderHelpText()` in router.ts to remove outdated "scaffold" terminology.

- **Parity test status**: Updated `packages/parity-tests/src/index.js` to report format status as `implemented` instead of `scaffolded`, matching the actual implementation state documented in `docs/parity/implementation-status.md`.

- **Package READMEs**: Updated `packages/word/README.md`, `packages/excel/README.md`, and `packages/ppt/README.md` to reflect that these are "Parity-first adapters" rather than "Lane-3 parity-first scaffolds".

### Added

- **Implemented count tracking**: Added `implementedCount` to the `summarizeParity()` function in `packages/core/src/parity.ts` to track the number of capabilities with `implemented` status.

## [0.1.0] - 2026-04-08

### Added

- **@officekit/cli** (0.1.0): Initial CLI package with command routing, parity-aware execution planning, and help text generation.

- **@officekit/core** (0.1.0): Core package with capability families, parity ledger, format detection, and error handling.

### Changed

- **Format packages** (0.0.0): Word, Excel, and PowerPoint adapter packages are available with fixture-backed implementations and extensive test coverage.

## Version Notes

- `@officekit/cli` and `@officekit/core` are at version 0.1.0 indicating initial release.
- Format packages (`@officekit/word`, `@officekit/excel`, `@officekit/ppt`) and utility packages remain at 0.0.0 as they continue development toward v1 parity.
- MCP support is explicitly excluded from v1 scope.
