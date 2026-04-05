import {
  buildExecutionPlan,
  type CommandResult,
  executeCommand,
  normalizeError,
  parseCliInput,
  renderHelpText,
  renderPlanResult,
  summarizeParity,
} from "@officekit/core";

const VERSION = "0.1.0";

export async function runCli(rawArgv: string[]): Promise<CommandResult> {
  const input = parseCliInput(rawArgv);
  const [command] = input.argv;

  if (input.version) {
    return { exitCode: 0, stdout: VERSION };
  }

  if (input.help || !command || command === "help") {
    return { exitCode: 0, stdout: renderHelpText() };
  }

  if (command === "about") {
    return {
      exitCode: 0,
      stdout: summarizeParity().lineage,
    };
  }

  if (command === "contracts") {
    return {
      exitCode: 0,
      stdout: JSON.stringify(summarizeParity(), null, input.json ? 2 : 0),
    };
  }

  try {
    if (input.plan) {
      const plan = buildExecutionPlan(input.argv);
      return {
        exitCode: 0,
        stdout: renderPlanResult(plan, input.json),
      };
    }

    if (command === "install") {
      // @ts-expect-error workspace JS runtime module
      const install = await import("../../install/src/index.js");
      const plan = install.buildInstallPlan();
      return {
        exitCode: 0,
        stdout: input.json ? JSON.stringify(plan, null, 2) : [`asset: ${plan.assetName}`, `install dir: ${plan.installDir}`, `path: ${plan.pathInstruction}`].join("\n"),
      };
    }

    if (command === "skills") {
      // @ts-expect-error workspace JS runtime module
      const skills = await import("../../skills/src/index.js");
      const subcommand = input.argv[1] ?? "list";
      if (subcommand === "list") {
        const bundles = skills.listSkillBundles();
        return {
          exitCode: 0,
          stdout: input.json ? JSON.stringify(bundles, null, 2) : bundles.map((bundle: { name: string; description: string }) => `${bundle.name}: ${bundle.description}`).join("\n"),
        };
      }

      if (subcommand === "install") {
        const bundleNames = input.argv.slice(2);
        const installed = await skills.installSkillBundles({ bundleNames: bundleNames.length > 0 ? bundleNames : ["officekit"] });
        return {
          exitCode: 0,
          stdout: input.json ? JSON.stringify(installed, null, 2) : installed.map((item: { agent: string; bundle: string }) => `${item.agent}: installed ${item.bundle}`).join("\n"),
        };
      }
    }

    if (command === "config") {
      // @ts-expect-error workspace JS runtime module
      const install = await import("../../install/src/index.js");
      const action = input.argv[1] ?? "list";
      if (action === "list") {
        const config = await install.readConfig();
        return { exitCode: 0, stdout: JSON.stringify(config, null, 2) };
      }
      if (action === "set") {
        const [key, value] = input.argv.slice(2);
        if (!key) {
          return { exitCode: 1, stderr: "config set requires <key> <value>" };
        }
        const current = await install.readConfig();
        const nextValue: unknown =
          value === "true" ? true : value === "false" ? false : value ?? null;
        await install.writeConfig({ ...current, [key]: nextValue });
        return { exitCode: 0, stdout: JSON.stringify({ ok: true, key, value: nextValue }, null, 2) };
      }
    }

    if (command === "help" && input.argv[1]) {
      // @ts-expect-error workspace JS runtime module
      const docs = await import("../../docs/src/index.js");
      const content = await docs.resolveCommandDoc(input.argv[1]).catch(() => null);
      if (content) {
        return { exitCode: 0, stdout: content };
      }
    }

    return await executeCommand(input.argv);
  } catch (error) {
    const normalized = normalizeError(error);
    const body = input.json
      ? JSON.stringify(
          {
            error: normalized.message,
            code: normalized.code,
            suggestion: normalized.suggestion,
          },
          null,
          2,
        )
      : [normalized.message, normalized.suggestion].filter(Boolean).join("\n");

    return {
      exitCode: 1,
      stderr: body,
    };
  }
}
