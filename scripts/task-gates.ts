import { spawn, spawnSync } from "node:child_process";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";

import { z } from "zod";

export const taskGateSchema = z.enum([
  "check",
  "check:heavy",
  "check:full",
  "fix",
]);

export type TaskGate = z.infer<typeof taskGateSchema>;

const scriptRoleSchema = z.enum([
  "quality:light",
  "quality:heavy",
  "quality:conditional",
  "quality:entry",
  "work:development",
  "work:repair",
  "work:release",
  "work:preparation",
  "work:lifecycle",
  "work:on-demand",
]);

type ScriptRole = z.infer<typeof scriptRoleSchema>;

// Turbo returns null for script/command on commandless tasks and includes extra metadata fields in query/dry-run JSON.
const turboQueryTaskItemSchema = z.object({
  name: z.string().trim().min(1),
  fullName: z.string().trim().min(1),
  tags: z.array(z.string()),
  script: z.string().nullish(),
  command: z.string().nullish(),
});

const turboQueryPackageItemSchema = z.object({
  name: z.string().trim().min(1),
  tasks: z.object({
    items: z.array(turboQueryTaskItemSchema),
  }),
});

const turboTaskDirectoryResponseSchema = z.object({
  data: z.object({
    packages: z.object({
      items: z.array(turboQueryPackageItemSchema),
    }),
  }),
});

const turboDryRunResponseSchema = z.object({
  tasks: z.array(
    z.object({
      taskId: z.string().trim().min(1),
    }),
  ),
});

type TurboTask = {
  readonly name: string;
  readonly fullName: string;
  readonly packageName: string;
  readonly tags: readonly string[];
  readonly executable: boolean;
  readonly role?: ScriptRole;
};

type GateLayer = {
  readonly role: "quality:light" | "quality:heavy" | "work:repair";
  readonly label: string;
};

type CommandResult = {
  readonly status: number;
  readonly stdout: string;
  readonly stderr: string;
  readonly error: Error | undefined;
};

const turboTaskQuery =
  "{ packages { items { name tasks { items { name fullName tags script command } } } } }";

const defaultEntryCommand = "node --conditions=source scripts/task-gates.ts";
const maxTurboOutputBytes = 16 * 1024 * 1024;

function hasCommand(value: string | null | undefined): boolean {
  return typeof value === "string" && value.trim().length > 0;
}

function formatIssuePath(path: readonly PropertyKey[]): string {
  if (path.length === 0) {
    return "<root>";
  }
  let formatted = "";
  for (const segment of path) {
    if (typeof segment === "number") {
      formatted += `[${String(segment)}]`;
    } else {
      const key = String(segment);
      formatted += formatted === "" ? key : `.${key}`;
    }
  }
  return formatted;
}

function formatSchemaIssues(
  label: string,
  issues: readonly z.core.$ZodIssue[],
): readonly string[] {
  return issues.map(
    (issue) =>
      `${label} invalid at ${formatIssuePath(issue.path)}: ${issue.message}`,
  );
}

function report(message: string): void {
  process.stderr.write(`${message}\n`);
}

type TurboCommand = readonly [string, ...string[]];

function runCaptured(
  cwd: string,
  args: readonly string[],
  command: TurboCommand = ["pnpm", "exec", "turbo"],
): CommandResult {
  const result = spawnSync(command[0], [...command.slice(1), ...args], {
    cwd,
    encoding: "utf8",
    maxBuffer: maxTurboOutputBytes,
  });

  return {
    status: result.status ?? 1,
    stdout: result.stdout ?? "",
    stderr: result.stderr ?? "",
    error: result.error,
  };
}

function runVisible(cwd: string, args: readonly string[]): Promise<number> {
  return new Promise((resolveStatus) => {
    const command =
      process.platform === "win32"
        ? "pnpm"
        : resolve(cwd, "node_modules/.bin/turbo");
    const commandArgs =
      process.platform === "win32" ? ["exec", "turbo", ...args] : args;
    const child = spawn(command, commandArgs, {
      cwd,
      detached: process.platform !== "win32",
      stdio: "inherit",
    });
    let interruptedBy: NodeJS.Signals | undefined;
    const forward = (signal: NodeJS.Signals): void => {
      if (child.exitCode === null && child.signalCode === null) {
        interruptedBy = signal;
        const turboSignal = signal === "SIGTERM" ? "SIGINT" : signal;
        if (process.platform !== "win32" && child.pid !== undefined) {
          try {
            process.kill(-child.pid, turboSignal);
          } catch {
            child.kill(turboSignal);
          }
        } else {
          child.kill(turboSignal);
        }
      }
    };
    const forwardInterrupt = (): void => forward("SIGINT");
    const forwardTerminate = (): void => forward("SIGTERM");
    const removeSignalHandlers = (): void => {
      process.off("SIGINT", forwardInterrupt);
      process.off("SIGTERM", forwardTerminate);
    };

    process.on("SIGINT", forwardInterrupt);
    process.on("SIGTERM", forwardTerminate);
    child.once("error", (error) => {
      removeSignalHandlers();
      report(`Failed to spawn Turbo for ${args.join(" ")}: ${error.message}`);
      resolveStatus(1);
    });
    child.once("close", (code, signal) => {
      removeSignalHandlers();
      if (interruptedBy !== undefined) {
        resolveStatus(interruptedBy === "SIGINT" ? 130 : 143);
      } else if (code !== null) {
        resolveStatus(code);
      } else {
        resolveStatus(
          signal === "SIGINT" ? 130 : signal === "SIGTERM" ? 143 : 1,
        );
      }
    });
  });
}

function commandFailure(label: string, result: CommandResult): string {
  const details = result.stderr.trim();
  if (result.error !== undefined) {
    return `Failed to run Turbo ${label}: ${result.error.message}`;
  }
  return `Turbo ${label} failed (exit code ${String(result.status)})${details === "" ? "" : `:\n${details}`}`;
}

function roleTags(tags: readonly string[]): readonly string[] {
  return tags.filter(
    (tag) => tag.startsWith("quality:") || tag.startsWith("work:"),
  );
}

function publicGateTask(task: {
  readonly name: string;
  readonly packageName: string;
}): boolean {
  return (
    task.packageName === "//" && taskGateSchema.safeParse(task.name).success
  );
}

function protectedVerificationTask(task: TurboTask): boolean {
  return (
    task.name === "check" ||
    task.name.startsWith("check:") ||
    task.name === "test" ||
    task.name.startsWith("test:") ||
    task.name === "lint" ||
    task.name === "typecheck" ||
    task.name === "format:check"
  );
}

function assignRole(
  task: Omit<TurboTask, "role">,
  diagnostics: string[],
): ScriptRole | undefined {
  const roles = roleTags(task.tags);
  const isPublicGate = publicGateTask(task);
  if (roles.length === 0) {
    diagnostics.push(
      `${task.fullName} is missing a unique role tag; please declare a quality scope or work purpose in Turbo task tags`,
    );
    return undefined;
  }
  if (roles.length > 1) {
    diagnostics.push(
      `${task.fullName} has multiple role tags: ${roles.join(", ")}; each executable task may declare only one role`,
    );
    return undefined;
  }

  const role = roles[0]!;
  const parsedRole = scriptRoleSchema.safeParse(role);
  if (!parsedRole.success) {
    diagnostics.push(`${task.fullName} uses unknown role tag ${role}`);
    return undefined;
  }
  const scriptRole = parsedRole.data;
  const publicRole = task.name === "fix" ? "work:repair" : "quality:entry";
  if (isPublicGate && scriptRole !== publicRole) {
    diagnostics.push(
      `${task.fullName} is a public task entrypoint; role must be ${publicRole}, but received ${scriptRole}`,
    );
    return undefined;
  }
  if (!isPublicGate && scriptRole === "quality:entry") {
    diagnostics.push(
      `${task.fullName} cannot use quality:entry; this role is only permitted for root check, check:heavy, and check:full`,
    );
    return undefined;
  }
  if (protectedVerificationTask(task) && !scriptRole.startsWith("quality:")) {
    diagnostics.push(
      `${task.fullName} is a protected verification task and cannot be assigned to non-verification role ${scriptRole}`,
    );
    return undefined;
  }

  return scriptRole;
}

function readTaskDirectory(
  cwd: string,
  entryCommand: string,
  turboCommand?: TurboCommand,
): {
  readonly tasks: readonly TurboTask[];
  readonly diagnostics: readonly string[];
} {
  const result = runCaptured(cwd, ["query", turboTaskQuery], turboCommand);
  if (result.error !== undefined || result.status !== 0) {
    return {
      diagnostics: [commandFailure("native task directory query", result)],
      tasks: [],
    };
  }

  let rawResponse: unknown;
  try {
    rawResponse = JSON.parse(result.stdout) as unknown;
  } catch (error) {
    return {
      diagnostics: [
        `Turbo native task directory is not valid JSON: ${error instanceof Error ? error.message : String(error)}`,
      ],
      tasks: [],
    };
  }

  const parsedResponse =
    turboTaskDirectoryResponseSchema.safeParse(rawResponse);
  if (!parsedResponse.success) {
    return {
      diagnostics: formatSchemaIssues(
        "Turbo native task directory",
        parsedResponse.error.issues,
      ),
      tasks: [],
    };
  }

  const diagnostics: string[] = [];
  const tasks: TurboTask[] = [];
  for (const packageItem of parsedResponse.data.data.packages.items) {
    for (const taskItem of packageItem.tasks.items) {
      const { name, fullName, tags, script, command } = taskItem;
      const executable = hasCommand(script) || hasCommand(command);
      const taskWithoutRole = {
        name,
        fullName,
        packageName: packageItem.name,
        tags,
        executable,
      };
      if (publicGateTask(taskWithoutRole)) {
        const expectedCommand = `${entryCommand} ${name}`;
        if (script !== expectedCommand || command !== expectedCommand) {
          diagnostics.push(
            `${fullName} public entrypoint wiring error; package.json script and effective command must be ${expectedCommand}`,
          );
        }
      }
      if (!executable && roleTags(tags).length === 0) {
        tasks.push(taskWithoutRole);
        continue;
      }

      const role = assignRole(taskWithoutRole, diagnostics);
      tasks.push(
        role === undefined ? taskWithoutRole : { ...taskWithoutRole, role },
      );
    }
  }

  for (const name of taskGateSchema.options) {
    if (
      !tasks.some((task) => task.packageName === "//" && task.name === name)
    ) {
      diagnostics.push(`Missing root public task entrypoint ${name}`);
    }
  }

  return { diagnostics, tasks };
}

function parseDryRun(
  cwd: string,
  tasks: readonly TurboTask[],
  turboCommand?: TurboCommand,
): {
  readonly taskIds: readonly string[];
  readonly diagnostics: readonly string[];
} {
  if (tasks.length === 0) {
    return { diagnostics: [], taskIds: [] };
  }
  const result = runCaptured(
    cwd,
    ["run", ...tasks.map((task) => task.fullName), "--dry=json"],
    turboCommand,
  );
  if (result.error !== undefined || result.status !== 0) {
    return {
      diagnostics: [commandFailure("task graph dry run", result)],
      taskIds: [],
    };
  }

  let rawResponse: unknown;
  try {
    rawResponse = JSON.parse(result.stdout) as unknown;
  } catch (error) {
    return {
      diagnostics: [
        `Turbo task graph dry run is not valid JSON: ${error instanceof Error ? error.message : String(error)}`,
      ],
      taskIds: [],
    };
  }

  const parsedResponse = turboDryRunResponseSchema.safeParse(rawResponse);
  if (!parsedResponse.success) {
    return {
      diagnostics: formatSchemaIssues(
        "Turbo task graph dry run",
        parsedResponse.error.issues,
      ),
      taskIds: [],
    };
  }

  return {
    diagnostics: [],
    taskIds: parsedResponse.data.tasks.map((task) => task.taskId),
  };
}

function validateLayerClosure(
  layer: GateLayer,
  taskIds: readonly string[],
  tasksByName: ReadonlyMap<string, TurboTask>,
): readonly string[] {
  const diagnostics: string[] = [];
  for (const taskId of taskIds) {
    const dependency = tasksByName.get(taskId);
    if (dependency === undefined) {
      diagnostics.push(
        `${layer.label} task graph contains task not present in directory: ${taskId}`,
      );
      continue;
    }
    if (publicGateTask(dependency)) {
      diagnostics.push(
        `${layer.label} task graph contains ${dependency.fullName}: leaf task cannot depend on public aggregate entrypoint`,
      );
    } else if (
      dependency.role !== undefined &&
      dependency.role !== layer.role &&
      dependency.role !== "work:preparation" &&
      !(layer.role !== "quality:light" && dependency.role === "quality:light")
    ) {
      const scope =
        layer.role === "quality:light" && dependency.role === "quality:heavy"
          ? "light check cannot indirectly trigger heavy verification"
          : dependency.role === "quality:conditional"
            ? "standard verification cannot indirectly trigger conditional verification"
            : `indirectly triggering role ${dependency.role} is not permitted`;
      diagnostics.push(
        `${layer.label} task graph contains ${dependency.fullName}: ${scope}`,
      );
    }
  }

  return diagnostics;
}

function layerFor(role: GateLayer["role"]): GateLayer {
  return {
    role,
    label:
      role === "quality:light"
        ? "light check"
        : role === "quality:heavy"
          ? "heavy verification"
          : "repair",
  };
}

function layersFor(gate: TaskGate): readonly GateLayer[] {
  if (gate === "check") {
    return [layerFor("quality:light")];
  }
  if (gate === "check:heavy") {
    return [layerFor("quality:heavy")];
  }
  if (gate === "fix") {
    return [layerFor("work:repair")];
  }

  return [layerFor("quality:light"), layerFor("quality:heavy")];
}

export function inspectTaskGate(
  gate: TaskGate,
  cwd = process.cwd(),
  entryCommand = defaultEntryCommand,
  turboCommand?: TurboCommand,
): {
  readonly diagnostics: readonly string[];
  readonly plans: readonly {
    readonly layer: GateLayer;
    readonly tasks: readonly TurboTask[];
  }[];
  readonly conditionalTasks: readonly TurboTask[];
} {
  if (!taskGateSchema.safeParse(gate).success) {
    return {
      diagnostics: [`Unknown public task entrypoint: ${gate}`],
      plans: [],
      conditionalTasks: [],
    };
  }

  const directory = readTaskDirectory(cwd, entryCommand, turboCommand);
  if (directory.diagnostics.length > 0) {
    return {
      diagnostics: directory.diagnostics,
      plans: [],
      conditionalTasks: [],
    };
  }
  const tasksByName = new Map(
    directory.tasks.map((task) => [task.fullName, task]),
  );
  const layers = layersFor(gate);
  const plans: {
    readonly layer: GateLayer;
    readonly tasks: readonly TurboTask[];
  }[] = [];
  const diagnostics: string[] = [];

  for (const task of directory.tasks) {
    if (task.executable || task.role === undefined) {
      continue;
    }
    const dryRun = parseDryRun(cwd, [task], turboCommand);
    diagnostics.push(...dryRun.diagnostics);
    if (
      dryRun.diagnostics.length === 0 &&
      !dryRun.taskIds.some(
        (taskId) => tasksByName.get(taskId)?.executable === true,
      )
    ) {
      diagnostics.push(
        `${task.fullName} declares dangling role tag ${task.role}: native execution graph of commandless node has no executable work`,
      );
    }
  }

  for (const layer of layers) {
    const selected = directory.tasks.filter(
      (task) => task.role === layer.role && !publicGateTask(task),
    );
    const dryRun = parseDryRun(cwd, selected, turboCommand);
    diagnostics.push(...dryRun.diagnostics);
    diagnostics.push(
      ...validateLayerClosure(layer, dryRun.taskIds, tasksByName),
    );
    if (
      layer.role === "quality:light" &&
      !dryRun.taskIds.some(
        (taskId) => tasksByName.get(taskId)?.executable === true,
      )
    ) {
      diagnostics.push(
        "Light check selected no executable tasks; declare quality:light for actual verification tasks",
      );
    }
    plans.push({ layer, tasks: selected });
  }

  const conditionalTasks = directory.tasks.filter(
    (task) => task.role === "quality:conditional",
  );
  return { diagnostics, plans, conditionalTasks };
}

export async function runTaskGate(
  gate: TaskGate,
  cwd = process.cwd(),
  entryCommand = defaultEntryCommand,
): Promise<number> {
  const { diagnostics, plans, conditionalTasks } = inspectTaskGate(
    gate,
    cwd,
    entryCommand,
  );
  if (diagnostics.length > 0) {
    report(diagnostics.join("\n"));
    return 1;
  }

  if (gate !== "fix" && conditionalTasks.length > 0) {
    report(
      `Conditional verification was not executed: ${conditionalTasks.map((task) => task.fullName).join(", ")}`,
    );
  }

  for (const plan of plans) {
    if (plan.tasks.length === 0) {
      report(`${plan.layer.label} selected no tasks, skipping execution`);
      continue;
    }
    const status = await runVisible(cwd, [
      "run",
      ...plan.tasks.map((task) => task.fullName),
      // Hosted runners share limited CPU across the rendering and CLI test processes.
      ...(process.env.GITHUB_ACTIONS === "true" &&
      plan.layer.role === "quality:heavy"
        ? ["--concurrency=1"]
        : []),
      "--continue=dependencies-successful",
      gate === "fix" ? "--output-logs=full" : "--output-logs=errors-only",
      "--log-order=grouped",
      "--log-prefix=task",
    ]);
    if (status !== 0) {
      return status;
    }
  }

  return 0;
}

export async function runTaskGateCli(
  entryCommand = defaultEntryCommand,
): Promise<void> {
  const [gate, ...extraArguments] = process.argv.slice(2);
  const parsedGate = taskGateSchema.safeParse(gate);
  if (!parsedGate.success || extraArguments.length > 0) {
    report(`Usage: ${entryCommand} <check|check:heavy|check:full|fix>`);
    process.exitCode = 2;
    return;
  }

  try {
    process.exitCode = await runTaskGate(
      parsedGate.data,
      process.cwd(),
      entryCommand,
    );
  } catch (error) {
    report(
      `Public task entrypoint failed: ${error instanceof Error ? error.message : String(error)}`,
    );
    process.exitCode = 1;
  }
}

const mainFile = process.argv[1];
if (
  mainFile !== undefined &&
  pathToFileURL(resolve(mainFile)).href === import.meta.url
) {
  void runTaskGateCli();
}
