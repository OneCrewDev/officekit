# officekit

> **officekit is the Bun + Node.js migration of OfficeCLI.**

`officekit` is rebuilding the OfficeCLI product surface for a Bun-packaged Node.js runtime while preserving **v1 capability/detail parity** for Word, Excel, and PowerPoint workflows, except for MCP support which is intentionally excluded.

## Lineage

This version is migrated from OfficeCLI. This repository is **migrated from OfficeCLI**. The goal is parity in document capability and operator detail, not command-line compatibility. The Node/Bun version is free to adopt a cleaner package layout and command surface as long as the OfficeCLI feature families remain covered and verifiable.

## Migration principles

- **Parity before polish** — feature/detail coverage comes before surface elegance.
- **Node-native packaging** — packages are split by capability instead of mirroring the .NET layout line-for-line.
- **Agent-first and developer-friendly** — preview/watch, docs/help, install/update, and skill installation are first-class migration lanes.
- **Fixture-backed verification** — lane work should produce runnable evidence, not only scaffolding.

## Current package lanes

- `packages/preview` — preview HTML shell, SSE relay server, and file-watch session glue.
- `packages/skills` — agent skill registry, bundle inventory, and install helpers.
- `packages/install` — platform-aware install/update/config planning helpers.
- `packages/docs` — markdown-backed docs/help loader for top-level product surfaces.

## Scope notes

- Command compatibility with `officecli` is **not** a goal.
- **Command compatibility is not required.**
- **MCP is excluded for v1.**
- **README lineage must stay explicit** so users understand this codebase is the OfficeCLI migration.

## Lane 4 migration surfaces ported here

These files/packages are the Node/Bun destination for OfficeCLI lane-4 concerns:

- preview/watch/html rendering
- skills packaging/install flow
- install/update/config flow
- docs/help/README lineage

## Verification intent

Lane-4 verification should show:

1. preview server + watch flow working end-to-end
2. skill bundle installation into detected agent directories
3. install/config/update helper behavior matching expected platform decisions
4. docs/help loader resolving migrated command documentation

## Source lineage references

The migration work in this repository is grounded by:

- `../OfficeCLI/README.md`
- `../OfficeCLI/SKILL.md`
- `../OfficeCLI/install.sh`
- `../OfficeCLI/install.ps1`
- `../OfficeCLI/src/officecli/Core/{Installer,SkillInstaller,UpdateChecker,WatchServer,WatchNotifier,HtmlPreviewHelper}.cs`
- `../OfficeCLI/src/officecli/{Program,HelpCommands,WikiHelpLoader,CommandBuilder.View,CommandBuilder.Watch}.cs`
