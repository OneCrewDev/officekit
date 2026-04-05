import { OfficekitError, UnsupportedCapabilityError, UsageError } from "./errors.js";
import { assertFormat, type SupportedFormat } from "./formats.js";
import {
  capabilityFamilies,
  officeCliLineageStatement,
  summarizeParity,
  type CapabilityFamily,
  type PackageLane,
} from "./parity.js";

export interface ParsedCliInput {
  argv: string[];
  json: boolean;
  plan: boolean;
  help: boolean;
  version: boolean;
}

export interface ExecutionPlan {
  command: CapabilityFamily;
  filePath?: string;
  format?: SupportedFormat;
  targetPackage: PackageLane;
  implementationStatus: "scaffold_only" | "planned";
  verificationHint: string;
  lineage: string;
  summary: string;
  rawArgs: string[];
}

const documentCommands = new Set<CapabilityFamily>([
  "create",
  "add",
  "set",
  "get",
  "query",
  "remove",
  "view",
  "raw",
  "watch",
  "check",
  "import",
]);

const formatToPackage: Record<SupportedFormat, Extract<PackageLane, "packages/word" | "packages/excel" | "packages/ppt">> = {
  word: "packages/word",
  excel: "packages/excel",
  powerpoint: "packages/ppt",
};

export function parseCliInput(argv: string[]): ParsedCliInput {
  const flags = new Set(argv.filter((token) => token.startsWith("--")));
  return {
    argv: argv.filter((token) => token !== "--plan" && token !== "--help" && token !== "--version"),
    json: flags.has("--json"),
    plan: flags.has("--plan"),
    help: flags.has("--help"),
    version: flags.has("--version"),
  };
}

export function buildExecutionPlan(argv: string[]): ExecutionPlan {
  if (argv.length === 0) {
    throw new UsageError("No command provided.", "Run 'officekit help' to inspect the scaffolded parity-aware CLI surface.");
  }

  const [command, ...rest] = argv;

  if (command === "mcp") {
    throw new UnsupportedCapabilityError(command);
  }

  if (!capabilityFamilies.includes(command as CapabilityFamily)) {
    throw new UsageError(`Unknown command '${command}'.`, `Supported commands: ${capabilityFamilies.join(", ")}, about, contracts.`);
  }

  const capability = command as CapabilityFamily;

  if (documentCommands.has(capability)) {
    const filePath = rest[0];
    if (!filePath) {
      throw new UsageError(`Command '${capability}' requires a target Office document path.`, `Example: officekit ${capability} demo.docx --plan --json`);
    }
    const format = assertFormat(filePath);
    return {
      command: capability,
      filePath,
      format,
      targetPackage: formatToPackage[format],
      implementationStatus: capability === "raw" || capability === "import" ? "planned" : "scaffold_only",
      verificationHint: `Use --plan to validate routing now; full behavior arrives in ${formatToPackage[format]}.`,
      lineage: officeCliLineageStatement,
      summary: `${capability} routes ${format} documents through ${formatToPackage[format]} once format handlers land.`,
      rawArgs: rest,
    };
  }

  const targetPackageByCommand: Record<Exclude<CapabilityFamily, typeof documentCommands extends Set<infer T> ? T : never>, PackageLane> = {
    batch: "packages/parity-tests",
    install: "packages/install",
    skills: "packages/skills",
    config: "packages/install",
    help: "packages/docs",
  };

  return {
    command: capability,
    targetPackage: targetPackageByCommand[capability as keyof typeof targetPackageByCommand],
    implementationStatus: capability === "batch" ? "planned" : "scaffold_only",
    verificationHint: `CLI contract is scaffolded; downstream work continues in ${targetPackageByCommand[capability as keyof typeof targetPackageByCommand]}.`,
    lineage: officeCliLineageStatement,
    summary: `${capability} is reserved as a first-class officekit product surface.`,
    rawArgs: rest,
  };
}

export function renderHelpText(): string {
  const parity = summarizeParity();
  return [
    "officekit CLI scaffold",
    officeCliLineageStatement,
    "",
    "Supported scaffolded commands:",
    `  ${capabilityFamilies.join(", ")}`,
    "",
    "Important scope note:",
    `  Excluded by design: ${parity.excluded.join(", ")}`,
    "",
    "Early vertical-slice verification examples:",
    "  officekit create demo.docx --plan --json",
    "  officekit create demo.xlsx --plan --json",
    "  officekit create demo.pptx --plan --json",
    "  officekit contracts --json",
  ].join("\n");
}

export function renderPlanResult(plan: ExecutionPlan, asJson: boolean): string {
  if (asJson) {
    return JSON.stringify(plan, null, 2);
  }

  return [
    `${plan.command}: ${plan.summary}`,
    `target package: ${plan.targetPackage}`,
    `status: ${plan.implementationStatus}`,
    `verification hint: ${plan.verificationHint}`,
    `lineage: ${plan.lineage}`,
  ].join("\n");
}

export function normalizeError(error: unknown): OfficekitError {
  if (error instanceof OfficekitError) {
    return error;
  }
  if (error instanceof Error) {
    return new OfficekitError(error.message);
  }
  return new OfficekitError("Unknown officekit failure.");
}
