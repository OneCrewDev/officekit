import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync, copyFileSync, statSync, existsSync, readdirSync } from 'node:fs';
import { dirname, extname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const scriptDir = dirname(fileURLToPath(import.meta.url));
const officekitRoot = resolve(scriptDir, '..');
const sourceRoot = resolve(officekitRoot, '..', 'OfficeCLI');
const fixtureRoot = resolve(officekitRoot, 'fixtures', 'officecli-source');
const docsRoot = resolve(officekitRoot, 'docs', 'migration');

if (!existsSync(sourceRoot)) {
  throw new Error(`Expected OfficeCLI source at ${sourceRoot}`);
}

const ensureDir = (dir) => mkdirSync(dir, { recursive: true });
const sha256 = (filePath) => createHash('sha256').update(readFileSync(filePath)).digest('hex');
const toPosix = (value) => value.split('\\').join('/');

const walk = (dir) => {
  const entries = [];
  for (const name of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, name.name);
    if (name.isDirectory()) entries.push(...walk(full));
    else entries.push(full);
  }
  return entries;
};

const extraFixtureSpecs = [
  { path: 'README.md', kind: 'readme', rationale: 'Top-level capability and workflow promises' },
  { path: 'SKILL.md', kind: 'skill', rationale: 'Agent-install surface and usage guidance' },
  { path: '.github/workflows/build.yml', kind: 'ci-workflow', rationale: 'Release packaging matrix plus smoke-test flow' },
  { path: 'src/officecli/Resources/preview.css', kind: 'preview-asset', rationale: 'Preview HTML styling fixture' },
  { path: 'src/officecli/Resources/preview.js', kind: 'preview-asset', rationale: 'Preview auto-refresh/browser behavior fixture' }
];

const shouldIncludeExample = (relativePath, size) => {
  const normalized = toPosix(relativePath);
  if (normalized.includes('/outputs/') || normalized.includes('/models/')) return false;
  const ext = extname(normalized).toLowerCase();
  if (['.md', '.sh', '.py', '.json', '.csv', '.tsv'].includes(ext)) return true;
  if (['.pptx', '.docx', '.xlsx'].includes(ext) && size <= 256 * 1024) return true;
  return false;
};

const exampleFixtures = walk(resolve(sourceRoot, 'examples')).flatMap((filePath) => {
  const relativePath = relative(sourceRoot, filePath);
  const size = statSync(filePath).size;
  if (!shouldIncludeExample(relativePath, size)) return [];
  return [{
    path: toPosix(relativePath),
    kind: 'example-fixture',
    rationale: normalizedRationale(relativePath),
  }];
});

function normalizedRationale(relativePath) {
  const normalized = toPosix(relativePath);
  if (normalized.startsWith('examples/word/')) return 'Word parity example or generator script';
  if (normalized.startsWith('examples/excel/')) return 'Excel parity example or generator script';
  if (normalized.startsWith('examples/ppt/')) return 'PowerPoint parity example or generator script';
  if (normalized.endsWith('.pptx') || normalized.endsWith('.docx') || normalized.endsWith('.xlsx')) return 'Small binary sample document from OfficeCLI examples';
  return 'Top-level example documentation or supporting fixture';
}

const fixtureSpecs = [...extraFixtureSpecs, ...exampleFixtures].sort((a, b) => a.path.localeCompare(b.path));

const ledger = [
  {
    sourceCluster: 'Program.cs update/install/skills/config entrypoints',
    evidence: 'OfficeCLI/src/officecli/Program.cs',
    officekitTarget: 'packages/cli/src/index.ts + packages/install + packages/skills + packages/core/config',
    ownerLane: 'lane-2 cli/core shell + lane-4 install/skills/docs',
    verification: 'unit + integration + e2e install/help flows',
    status: 'inventory',
    notes: 'MCP subcommands are explicitly excluded from officekit v1; install/skills/config/update remain required.'
  },
  {
    sourceCluster: 'Root command, --json envelope, resident open/close plumbing',
    evidence: 'OfficeCLI/src/officecli/CommandBuilder.cs',
    officekitTarget: 'packages/cli/src/root-command.ts + packages/core/src/resident/* + packages/core/src/output/*',
    ownerLane: 'lane-2 cli/core shell',
    verification: 'unit + integration',
    status: 'inventory',
    notes: 'Officekit should preserve agent-friendly JSON envelopes and resident-mode ergonomics even if syntax changes.'
  },
  {
    sourceCluster: 'Add/remove/move/swap DOM mutation commands',
    evidence: 'OfficeCLI/src/officecli/CommandBuilder.Add.cs',
    officekitTarget: 'packages/cli/src/commands/dom-mutations.ts + format packages',
    ownerLane: 'lane-2 cli/core shell + lane-3 format adapters',
    verification: 'integration + differential parity',
    status: 'inventory',
    notes: 'Covers insertion positioning, copy-from, force-write, and reordering semantics.'
  },
  {
    sourceCluster: 'Set command and property mutation grammar',
    evidence: 'OfficeCLI/src/officecli/CommandBuilder.Set.cs',
    officekitTarget: 'packages/cli/src/commands/set.ts + packages/core/src/property-parser.ts',
    ownerLane: 'lane-2 cli/core shell + lane-3 format adapters',
    verification: 'unit + integration',
    status: 'inventory',
    notes: 'Property normalization and error handling must align across formats.'
  },
  {
    sourceCluster: 'Get/query read commands',
    evidence: 'OfficeCLI/src/officecli/CommandBuilder.GetQuery.cs',
    officekitTarget: 'packages/cli/src/commands/get-query.ts + packages/core/src/selectors/*',
    ownerLane: 'lane-2 cli/core shell',
    verification: 'unit + integration + differential parity',
    status: 'inventory',
    notes: 'Selector grammar and JSON/text result shaping are parity critical for agents.'
  },
  {
    sourceCluster: 'Raw/raw-set/add-part XML fallback commands',
    evidence: 'OfficeCLI/src/officecli/CommandBuilder.Raw.cs',
    officekitTarget: 'packages/cli/src/commands/raw.ts + packages/core/src/raw-xml/*',
    ownerLane: 'lane-2 cli/core shell + lane-3 format adapters',
    verification: 'integration + differential parity',
    status: 'inventory',
    notes: 'Universal fallback for long-tail OpenXML operations; add-part required for chart/header/footer creation.'
  },
  {
    sourceCluster: 'Validate/check quality commands',
    evidence: 'OfficeCLI/src/officecli/CommandBuilder.Check.cs',
    officekitTarget: 'packages/cli/src/commands/validate-check.ts + packages/core/src/validation/*',
    ownerLane: 'lane-2 cli/core shell + lane-4 preview/docs',
    verification: 'integration + fixture-backed issue snapshots',
    status: 'inventory',
    notes: 'Covers OpenXML validation plus higher-level layout/overflow issue reporting.'
  },
  {
    sourceCluster: 'Batch command / one-open-save-cycle execution',
    evidence: 'OfficeCLI/src/officecli/CommandBuilder.Batch.cs + Core/BatchTypes.cs',
    officekitTarget: 'packages/cli/src/commands/batch.ts + packages/core/src/batch/*',
    ownerLane: 'lane-2 cli/core shell',
    verification: 'unit + integration + e2e',
    status: 'inventory',
    notes: 'Needs stdin, file, inline JSON array, and force-continue semantics.'
  },
  {
    sourceCluster: 'Create/import/merge commands',
    evidence: 'OfficeCLI/src/officecli/CommandBuilder.Import.cs + Core/TemplateMerger.cs',
    officekitTarget: 'packages/cli/src/commands/create-import-merge.ts + packages/core/src/templates/*',
    ownerLane: 'lane-2 cli/core shell + lane-3 format adapters',
    verification: 'integration + fixture-backed template scenarios',
    status: 'inventory',
    notes: 'Includes CSV/TSV import, blank doc creation, and JSON template merge across docx/xlsx/pptx.'
  },
  {
    sourceCluster: 'View/watch/unwatch/preview command surface',
    evidence: 'OfficeCLI/src/officecli/CommandBuilder.View.cs + CommandBuilder.Watch.cs + Core/HtmlPreviewHelper.cs + Core/Watch*.cs',
    officekitTarget: 'packages/preview/src/* + packages/cli/src/commands/view-watch.ts',
    ownerLane: 'lane-4 preview/skills/install/docs',
    verification: 'integration + rendered snapshot + manual browser smoke',
    status: 'inventory',
    notes: 'HTML preview, outline/text/stats/issues/svg/forms modes, and live watch refresh are explicit parity targets.'
  },
  {
    sourceCluster: 'Built-in help and wiki-backed help surfaces',
    evidence: 'OfficeCLI/src/officecli/HelpCommands.cs + WikiHelpLoader.cs',
    officekitTarget: 'packages/docs/src/help/* + packages/cli/src/help-command.ts',
    ownerLane: 'lane-4 preview/skills/install/docs',
    verification: 'docs acceptance + CLI golden output tests',
    status: 'inventory',
    notes: 'Format-prefixed deep help must survive the migration even if rendered from new source files.'
  },
  {
    sourceCluster: 'Word handler domain boundary',
    evidence: 'OfficeCLI/src/officecli/Handlers/WordHandler.cs',
    officekitTarget: 'packages/word/src/word-handler.ts',
    ownerLane: 'lane-3 format adapters',
    verification: 'format integration + differential parity',
    status: 'inventory',
    notes: 'Owns Word DOM/raw/query/set/add/remove/batch semantics.'
  },
  {
    sourceCluster: 'Excel handler domain boundary',
    evidence: 'OfficeCLI/src/officecli/Handlers/ExcelHandler.cs',
    officekitTarget: 'packages/excel/src/excel-handler.ts',
    ownerLane: 'lane-3 format adapters',
    verification: 'format integration + differential parity',
    status: 'inventory',
    notes: 'Owns worksheets, ranges, formulas, charts, pivot tables, import, raw filtering, and validation.'
  },
  {
    sourceCluster: 'PowerPoint handler domain boundary',
    evidence: 'OfficeCLI/src/officecli/Handlers/PowerPointHandler.cs',
    officekitTarget: 'packages/ppt/src/ppt-handler.ts',
    ownerLane: 'lane-3 format adapters',
    verification: 'format integration + differential parity + preview snapshots',
    status: 'inventory',
    notes: 'Owns slides, shapes, connectors, media, charts, notes, placeholders, and preview-sensitive rendering flows.'
  },
  {
    sourceCluster: 'Document abstraction + factory + shared node model',
    evidence: 'OfficeCLI/src/officecli/Core/DocumentHandlerFactory.cs + Core/IDocumentHandler.cs + Core/DocumentNode.cs',
    officekitTarget: 'packages/core/src/document-model/*',
    ownerLane: 'lane-2 cli/core shell',
    verification: 'unit + integration',
    status: 'inventory',
    notes: 'Critical seam for thin CLI shell and reusable format adapters.'
  },
  {
    sourceCluster: 'Selector/query/path semantics',
    evidence: 'OfficeCLI/src/officecli/Core/PathAliases.cs + Core/AttributeFilter.cs + Core/GenericXmlQuery.cs',
    officekitTarget: 'packages/core/src/selectors/*',
    ownerLane: 'lane-2 cli/core shell',
    verification: 'unit + differential parity',
    status: 'inventory',
    notes: 'Cross-format semantics gate; must be validated before broad handler migration.'
  },
  {
    sourceCluster: 'Units, EMU, spacing, colors, parse helpers, theme resolution',
    evidence: 'OfficeCLI/src/officecli/Core/Units.cs + Core/EmuConverter.cs + Core/SpacingConverter.cs + Core/ColorMath.cs + Core/ParseHelpers.cs + Core/ThemeColorResolver.cs',
    officekitTarget: 'packages/core/src/values/*',
    ownerLane: 'lane-2 cli/core shell',
    verification: 'unit',
    status: 'inventory',
    notes: 'Foundational parsing/formatting semantics reused by all document packages.'
  },
  {
    sourceCluster: 'Output envelopes, CLI exceptions, issue modeling',
    evidence: 'OfficeCLI/src/officecli/Core/OutputFormatter.cs + Core/CliException.cs + Core/DocumentIssue.cs',
    officekitTarget: 'packages/core/src/output/*',
    ownerLane: 'lane-2 cli/core shell',
    verification: 'unit + snapshot',
    status: 'inventory',
    notes: 'Agent-facing result and error shapes need regression coverage.'
  },
  {
    sourceCluster: 'Raw XML execution + extended/theme/document properties helpers',
    evidence: 'OfficeCLI/src/officecli/Core/RawXmlHelper.cs + Core/ExtendedPropertiesHandler.cs + Core/ThemeHandler.cs',
    officekitTarget: 'packages/core/src/raw-xml/* + format packages',
    ownerLane: 'lane-2 cli/core shell + lane-3 format adapters',
    verification: 'integration + differential parity',
    status: 'inventory',
    notes: 'Supports long-tail fallback and metadata/theme/property editing across formats.'
  },
  {
    sourceCluster: 'Charts, formulas, styles, pivot helpers',
    evidence: 'OfficeCLI/src/officecli/Core/Chart*.cs + Core/Formula*.cs + Core/ExcelStyleManager.cs + Core/PivotTableHelper.cs',
    officekitTarget: 'packages/excel/src/* + packages/ppt/src/chart/* + packages/word/src/chart/*',
    ownerLane: 'lane-3 format adapters',
    verification: 'format integration + differential parity',
    status: 'inventory',
    notes: 'Large helper cluster covering chart creation/inspection/SVG rendering, formula evaluation, style management, and pivots.'
  },
  {
    sourceCluster: 'Resident client/server + CLI logging',
    evidence: 'OfficeCLI/src/officecli/Core/ResidentClient.cs + Core/ResidentServer.cs + Core/CliLogger.cs',
    officekitTarget: 'packages/core/src/resident/* + packages/core/src/logging/*',
    ownerLane: 'lane-2 cli/core shell',
    verification: 'integration + e2e',
    status: 'inventory',
    notes: 'Resident mode is a headline README feature and must be testable as a local workflow.'
  },
  {
    sourceCluster: 'Installer / skills / updater surfaces',
    evidence: 'OfficeCLI/src/officecli/Core/Installer.cs + Core/SkillInstaller.cs + Core/UpdateChecker.cs',
    officekitTarget: 'packages/install/src/* + packages/skills/src/* + packages/core/src/config/update.ts',
    ownerLane: 'lane-4 preview/skills/install/docs',
    verification: 'integration + docs acceptance',
    status: 'inventory',
    notes: 'Auto-detect agent tooling, install skill assets, manage PATH/config, and background update checks.'
  },
  {
    sourceCluster: 'MCP server and installer',
    evidence: 'OfficeCLI/src/officecli/Core/McpServer.cs + Core/McpInstaller.cs + Program.cs mcp branch',
    officekitTarget: 'excluded',
    ownerLane: 'n/a',
    verification: 'n/a',
    status: 'excluded',
    notes: 'PRD explicitly excludes MCP from officekit v1 parity.'
  }
];

const capabilities = [
  { family: 'CLI shell', capability: 'create blank .docx/.xlsx/.pptx', sourceEvidence: 'CommandBuilder.Import.cs#create + README quick start', targetPackage: 'packages/cli + format packages', verification: 'integration + smoke', status: 'inventory', notes: 'Type may be inferred from extension or overridden.' },
  { family: 'CLI shell', capability: 'open/close resident mode', sourceEvidence: 'CommandBuilder.cs + README resident mode section', targetPackage: 'packages/cli + packages/core', verification: 'integration + e2e', status: 'inventory', notes: 'Resident mode is a top-level workflow for low-latency repeated operations.' },
  { family: 'CLI shell', capability: 'get/query structured inspection', sourceEvidence: 'CommandBuilder.GetQuery.cs + README command reference', targetPackage: 'packages/cli + packages/core', verification: 'unit + differential parity', status: 'inventory', notes: 'JSON and text outputs both matter.' },
  { family: 'CLI shell', capability: 'set/add/remove/move/swap mutations', sourceEvidence: 'CommandBuilder.Add.cs + CommandBuilder.Set.cs + README patterns', targetPackage: 'packages/cli + format packages', verification: 'integration + differential parity', status: 'inventory', notes: 'Covers copy-from, insert positions, and property setters.' },
  { family: 'CLI shell', capability: 'batch multi-command execution', sourceEvidence: 'CommandBuilder.Batch.cs + README resident/batch section', targetPackage: 'packages/cli + packages/core', verification: 'integration + e2e', status: 'inventory', notes: 'stdin/input file/inline JSON modes required.' },
  { family: 'CLI shell', capability: 'raw/raw-set/add-part fallback', sourceEvidence: 'CommandBuilder.Raw.cs + README three-layer architecture', targetPackage: 'packages/cli + packages/core', verification: 'integration + differential parity', status: 'inventory', notes: 'Long-tail parity escape hatch.' },
  { family: 'CLI shell', capability: 'validate/check quality scans', sourceEvidence: 'CommandBuilder.Check.cs + README command reference', targetPackage: 'packages/cli + packages/core', verification: 'integration + snapshot', status: 'inventory', notes: 'Schema validation and higher-level issue reporting are distinct.' },
  { family: 'CLI shell', capability: 'view modes: text/annotated/outline/stats/issues/html/svg/forms', sourceEvidence: 'CommandBuilder.View.cs + README live preview section', targetPackage: 'packages/preview + packages/cli', verification: 'integration + rendered snapshots', status: 'inventory', notes: 'Preview family is developer-facing, not optional polish.' },
  { family: 'CLI shell', capability: 'watch/unwatch live preview server', sourceEvidence: 'CommandBuilder.Watch.cs + README developer preview quick start', targetPackage: 'packages/preview + packages/cli', verification: 'manual browser smoke + integration', status: 'inventory', notes: 'Auto-refresh semantics need fixture-backed coverage.' },
  { family: 'CLI shell', capability: 'import CSV/TSV into Excel sheets', sourceEvidence: 'CommandBuilder.Import.cs + README common patterns', targetPackage: 'packages/excel + packages/cli', verification: 'integration', status: 'inventory', notes: 'Supports file or stdin sources, header and start-cell options.' },
  { family: 'CLI shell', capability: 'template merge for docx/xlsx/pptx', sourceEvidence: 'CommandBuilder.Import.cs#merge + Core/TemplateMerger.cs + README workflow examples', targetPackage: 'packages/core + format packages', verification: 'integration + fixture-backed parity', status: 'inventory', notes: 'Includes unresolved placeholder reporting.' },
  { family: 'CLI shell', capability: 'help / format-prefixed deep help', sourceEvidence: 'HelpCommands.cs + README built-in help section', targetPackage: 'packages/docs + packages/cli', verification: 'snapshot + docs acceptance', status: 'inventory', notes: 'Command compatibility can change, but deep discoverability must remain.' },
  { family: 'CLI shell', capability: 'install/skills/config/update', sourceEvidence: 'Program.cs + Installer.cs + SkillInstaller.cs + UpdateChecker.cs + README installation/update sections', targetPackage: 'packages/install + packages/skills + packages/core', verification: 'integration + docs acceptance', status: 'inventory', notes: 'Officekit naming and lineage wording required.' },
  { family: 'CLI shell', capability: 'MCP server/install', sourceEvidence: 'Program.cs mcp branch + README AI integration section', targetPackage: 'excluded', verification: 'n/a', status: 'excluded', notes: 'Explicitly out of scope for officekit v1.' },
  { family: 'Shared core', capability: 'selector/path alias grammar', sourceEvidence: 'Core/PathAliases.cs + Core/AttributeFilter.cs + Core/GenericXmlQuery.cs', targetPackage: 'packages/core', verification: 'unit + differential parity', status: 'inventory', notes: 'Phase-gate semantic surface.' },
  { family: 'Shared core', capability: 'units/colors/EMU/theme parsing', sourceEvidence: 'Core/Units.cs + Core/EmuConverter.cs + Core/ColorMath.cs + Core/ParseHelpers.cs + Core/ThemeColorResolver.cs', targetPackage: 'packages/core', verification: 'unit', status: 'inventory', notes: 'Cross-format numeric and color behavior.' },
  { family: 'Shared core', capability: 'JSON/text output envelopes', sourceEvidence: 'Core/OutputFormatter.cs + CommandBuilder.cs', targetPackage: 'packages/core', verification: 'snapshot + unit', status: 'inventory', notes: 'AI-facing result shape must be stable.' },
  { family: 'Shared core', capability: 'resident IPC and batch request schemas', sourceEvidence: 'Core/ResidentClient.cs + Core/ResidentServer.cs + Core/BatchTypes.cs', targetPackage: 'packages/core', verification: 'integration + e2e', status: 'inventory', notes: 'Needed for low-latency loops and atomic updates.' },
  { family: 'Shared core', capability: 'raw XML helper / metadata / theme helpers', sourceEvidence: 'Core/RawXmlHelper.cs + Core/ExtendedPropertiesHandler.cs + Core/ThemeHandler.cs', targetPackage: 'packages/core + format packages', verification: 'integration', status: 'inventory', notes: 'Backs long-tail escape hatch plus document/theme metadata.' },
  { family: 'Word', capability: 'paragraphs, runs, tables, styles', sourceEvidence: 'README key features Word list + Handlers/WordHandler.cs', targetPackage: 'packages/word', verification: 'format integration + differential parity', status: 'inventory', notes: 'Core document editing path.' },
  { family: 'Word', capability: 'headers/footers, images, equations, hyperlinks', sourceEvidence: 'README key features Word list + WordHandler raw layer', targetPackage: 'packages/word', verification: 'format integration + fixture-backed parity', status: 'inventory', notes: 'Long-tail but promised in source docs.' },
  { family: 'Word', capability: 'comments, footnotes, bookmarks, TOC, sections', sourceEvidence: 'README key features Word list', targetPackage: 'packages/word', verification: 'format integration', status: 'inventory', notes: 'Needs explicit status tracking even before implementation.' },
  { family: 'Word', capability: 'watermarks, form fields, content controls, document properties', sourceEvidence: 'README key features Word list + Core/ExtendedPropertiesHandler.cs', targetPackage: 'packages/word + packages/core', verification: 'format integration', status: 'inventory', notes: 'Properties cross-cut with shared metadata helper.' },
  { family: 'Excel', capability: 'workbook/sheet/cell/range operations', sourceEvidence: 'README key features Excel list + Handlers/ExcelHandler.cs', targetPackage: 'packages/excel', verification: 'format integration + differential parity', status: 'inventory', notes: 'Includes $Sheet:A1 addressing.' },
  { family: 'Excel', capability: 'formulas and evaluator', sourceEvidence: 'README key features Excel list + Core/FormulaEvaluator*.cs', targetPackage: 'packages/excel', verification: 'unit + integration', status: 'inventory', notes: '150+ built-in function claim requires evidence-backed migration.' },
  { family: 'Excel', capability: 'styles, conditional formatting, data validation, autofilter', sourceEvidence: 'README key features Excel list + Core/ExcelStyleManager.cs', targetPackage: 'packages/excel', verification: 'format integration', status: 'inventory', notes: 'Style semantics are a major parity risk.' },
  { family: 'Excel', capability: 'charts and SVG/preview rendering', sourceEvidence: 'README key features Excel list + Core/Chart*.cs', targetPackage: 'packages/excel + packages/preview', verification: 'format integration + rendered snapshots', status: 'inventory', notes: 'Shared chart helper cluster also supports PPT/Word embeddings.' },
  { family: 'Excel', capability: 'pivot tables, named ranges, sparklines, comments, shapes', sourceEvidence: 'README key features Excel list + Core/PivotTableHelper.cs', targetPackage: 'packages/excel', verification: 'format integration + differential parity', status: 'inventory', notes: 'Long-tail parity families called out explicitly in PRD/test spec.' },
  { family: 'PowerPoint', capability: 'slides, shapes, tables, images', sourceEvidence: 'README key features PowerPoint list + Handlers/PowerPointHandler.cs', targetPackage: 'packages/ppt', verification: 'format integration + differential parity', status: 'inventory', notes: 'Baseline presentation editing path.' },
  { family: 'PowerPoint', capability: 'charts, themes, placeholders, notes, groups, connectors', sourceEvidence: 'README key features PowerPoint list + Core/Chart*.cs + PowerPointHandler.cs', targetPackage: 'packages/ppt', verification: 'format integration + preview snapshots', status: 'inventory', notes: 'Connector/group/media handling is preview-sensitive.' },
  { family: 'PowerPoint', capability: 'animations, morph transitions, slide zoom, 3D models', sourceEvidence: 'README key features PowerPoint list + examples/ppt/*', targetPackage: 'packages/ppt', verification: 'fixture-backed parity + manual inspection', status: 'inventory', notes: 'Examples provide real fixtures for these long-tail promises.' },
  { family: 'PowerPoint', capability: 'video/audio/media embedding', sourceEvidence: 'README key features PowerPoint list + PowerPointHandler.cs media cleanup paths', targetPackage: 'packages/ppt', verification: 'format integration', status: 'inventory', notes: 'Media relationship handling needs dedicated fixtures.' },
  { family: 'Preview/docs', capability: 'developer quick-start live preview workflow', sourceEvidence: 'README developer 30-second flow + build workflow smoke test', targetPackage: 'packages/preview + README', verification: 'manual browser smoke + docs acceptance', status: 'inventory', notes: 'Must work with officekit branding and lineage statement.' },
  { family: 'Skills/install', capability: 'agent skill discovery and installation', sourceEvidence: 'SKILL.md + Program.cs skills branch + Installer.cs', targetPackage: 'packages/skills + packages/install', verification: 'integration + docs acceptance', status: 'inventory', notes: 'OfficeCLI auto-installs across agent tools; officekit needs equivalent story.' },
  { family: 'Docs/examples', capability: 'README quick starts and end-to-end examples', sourceEvidence: 'README.md + examples/**/* + build workflow smoke section', targetPackage: 'packages/docs + fixtures/officecli-source', verification: 'docs acceptance + fixture manifest', status: 'inventory', notes: 'Lane 1 harvests these as migration fixtures.' }
];

const smokeCommands = [
  'create test_smoke.docx',
  'add test_smoke.docx /body --type paragraph --prop text="Hello from CI"',
  'get test_smoke.docx /body/p[1]'
];

const writeJson = (path, data) => {
  ensureDir(dirname(path));
  writeFileSync(path, JSON.stringify(data, null, 2) + '\n');
};

const writeText = (path, text) => {
  ensureDir(dirname(path));
  writeFileSync(path, text);
};

for (const spec of fixtureSpecs) {
  const sourcePath = resolve(sourceRoot, spec.path);
  const targetPath = resolve(fixtureRoot, spec.path);
  ensureDir(dirname(targetPath));
  copyFileSync(sourcePath, targetPath);
}

const fixtureManifest = fixtureSpecs.map((spec) => {
  const sourcePath = resolve(sourceRoot, spec.path);
  const stats = statSync(sourcePath);
  return {
    path: spec.path,
    kind: spec.kind,
    rationale: spec.rationale,
    bytes: stats.size,
    sha256: sha256(sourcePath)
  };
});

const skippedExamples = walk(resolve(sourceRoot, 'examples')).flatMap((filePath) => {
  const relativePath = toPosix(relative(sourceRoot, filePath));
  const size = statSync(filePath).size;
  return shouldIncludeExample(relativePath, size)
    ? []
    : [{ path: relativePath, bytes: size, reason: size > 256 * 1024 ? 'excluded-large-binary-or-generated-output' : 'excluded-non-curated-extension-or-output-model' }];
}).sort((a, b) => a.path.localeCompare(b.path));

writeJson(resolve(docsRoot, 'source-to-target-ledger.json'), ledger);
writeJson(resolve(docsRoot, 'capability-matrix.json'), capabilities);
writeJson(resolve(fixtureRoot, 'manifest.json'), {
  generatedAt: new Date().toISOString(),
  sourceRoot: toPosix(relative(officekitRoot, sourceRoot)),
  smokeCommands,
  included: fixtureManifest,
  skippedExamples
});
writeText(resolve(fixtureRoot, 'ci-smoke-flow.txt'), smokeCommands.map((command) => `officecli ${command}`).join('\n') + '\n');

const table = (rows, columns) => {
  const header = `| ${columns.join(' | ')} |`;
  const divider = `| ${columns.map(() => '---').join(' | ')} |`;
  const body = rows.map((row) => `| ${columns.map((key) => String(row[key] ?? '').replace(/\|/g, '\\|')).join(' | ')} |`);
  return [header, divider, ...body].join('\n');
};

writeText(resolve(docsRoot, 'source-to-target-ledger.md'), `# Source-to-target ledger\n\nThis ledger inventories the OfficeCLI source clusters assigned to lane 1 and maps them to the planned officekit package/module ownership from the approved PRD. Status reflects inventory-only progress unless otherwise stated.\n\n${table(ledger.map((row) => ({
  'Source cluster': row.sourceCluster,
  Evidence: row.evidence,
  'Officekit target': row.officekitTarget,
  'Owner lane': row.ownerLane,
  Verification: row.verification,
  Status: row.status,
  Notes: row.notes
})), ['Source cluster', 'Evidence', 'Officekit target', 'Owner lane', 'Verification', 'Status', 'Notes'])}\n`);

writeText(resolve(docsRoot, 'capability-matrix.md'), `# Capability matrix\n\nThis matrix captures capability/detail families harvested from OfficeCLI source + README/examples/CI evidence. MCP is marked excluded per the approved migration plan.\n\n${table(capabilities.map((row) => ({
  Family: row.family,
  Capability: row.capability,
  'Source evidence': row.sourceEvidence,
  'Target package': row.targetPackage,
  Verification: row.verification,
  Status: row.status,
  Notes: row.notes
})), ['Family', 'Capability', 'Source evidence', 'Target package', 'Verification', 'Status', 'Notes'])}\n`);

const includedSummary = fixtureManifest.reduce((acc, item) => {
  acc.count += 1;
  acc.bytes += item.bytes;
  acc.kinds[item.kind] = (acc.kinds[item.kind] || 0) + 1;
  return acc;
}, { count: 0, bytes: 0, kinds: {} });

const skippedLarge = skippedExamples.filter((item) => item.reason === 'excluded-large-binary-or-generated-output');

writeText(resolve(docsRoot, 'fixture-harvest-report.md'), `# Fixture harvest report\n\n## Included harvest\n- Included files: ${includedSummary.count}\n- Included bytes: ${includedSummary.bytes}\n- Included kinds: ${Object.entries(includedSummary.kinds).map(([kind, count]) => `${kind}=${count}`).join(', ')}\n\n## CI smoke flow harvested\n${smokeCommands.map((command) => `- \`officecli ${command}\``).join('\n')}\n\n## Notable exclusions\n- Large generated outputs and model assets were intentionally not copied into the curated fixture set to keep lane-1 commits reviewable.\n- MCP-specific assets are tracked as excluded by design in the capability matrix/ledger.\n\n## Largest skipped example assets\n${skippedLarge.slice(-10).map((item) => `- \`${item.path}\` (${item.bytes} bytes) — ${item.reason}`).join('\n') || '- none'}\n`);

console.log(`Generated lane-1 migration artifacts in ${relative(process.cwd(), officekitRoot) || '.'}`);
console.log(`Harvested fixtures: ${fixtureManifest.length}`);
console.log(`Skipped examples: ${skippedExamples.length}`);
