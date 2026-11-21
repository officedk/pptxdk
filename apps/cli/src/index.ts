import { randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import { mkdir, rename, rm, stat, writeFile } from "node:fs/promises";
import { basename, dirname, extname, resolve } from "node:path";

import {
  CliWriteError,
  type DataOutcome,
  defineCli,
  executeCli,
  type FailureOutcome,
  helpCapability,
  parseCliInvocation,
  schemaCapability,
  text,
  versionCapability,
  type WriteCliOutput,
} from "@ykdz/cli-contract";
import { isNodeBrokenPipe } from "@ykdz/cli-contract/node";
import { BuildError, targets, type Target } from "pptxdk";
import { buildFromSource, SourceLoadError } from "pptxdk/node";
import { z } from "zod";

/**
 * Runtime seam for programmatic CLI execution.
 *
 * `argv` uses Node's `process.argv` layout (user arguments start at `argv[2]`).
 * `cwd` defaults to `process.cwd()` when omitted.
 */
export type CliRuntime = {
  readonly argv: readonly string[];
  readonly cwd?: string;
  readonly signal?: AbortSignal;
  readonly write: WriteCliOutput;
};

type CliDependencies = {
  readonly cwd: string;
  readonly signal?: AbortSignal;
};

// Read package.json at runtime rather than statically importing it because tsconfig.build.json sets rootDir to src.
const { version: cliVersion } = JSON.parse(
  readFileSync(new URL("../package.json", import.meta.url), "utf8"),
) as { readonly version: string };

const targetChoices = [
  targets.transitional.name,
  ...Object.values(targets)
    .map((target) => target.name)
    .filter((name) => name !== targets.transitional.name),
] as const;

const diagnosticItemSchema = z.object({
  path: z.string(),
  code: z.string(),
  message: z.string(),
});

const buildInputSchema = z.object({
  entry: z.string().min(1),
  out: z.string().min(1).optional(),
  export: z.string().min(1).optional(),
  target: z.enum(targetChoices).optional(),
  embedFonts: z.enum(["full", "subset"]).optional(),
});

const builtSchema = z.object({
  output: z.string(),
  slides: z.number().int().nonnegative(),
  warnings: z.array(diagnosticItemSchema),
});

const invalidOutputSchema = z.object({
  entry: z.string(),
  output: z.string(),
});

const sourceLoadFailedSchema = z.object({
  message: z.string(),
});

const buildFailedSchema = z.object({
  diagnostics: z.array(diagnosticItemSchema),
});

const writeFailedSchema = z.object({
  output: z.string(),
  message: z.string(),
});

const cancelledSchema = z.object({});

function targetByName(name: string): Target | undefined {
  return Object.values(targets).find((target) => target.name === name);
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/**
 * Normalize dynamic text into the non-empty, single-line, NUL-free tuple
 * required by `@ykdz/cli-contract`'s `text.lines(...)` invariant.
 */
function toContractTextLines(lines: readonly string[]): [string, ...string[]] {
  const [first = "Unknown error", ...rest] = lines.flatMap((line) =>
    line.replaceAll("\0", "\\0").split(/\r\n|\r|\n/u),
  );
  return [
    first === "" && rest.every((line) => line === "") ? "Unknown error" : first,
    ...rest,
  ];
}

async function writeArtifact(path: string, bytes: Uint8Array): Promise<void> {
  const directory = dirname(path);
  const temporary = resolve(
    directory,
    `.${basename(path)}.${randomUUID()}.tmp`,
  );
  try {
    await mkdir(directory, { recursive: true });
    await writeFile(temporary, bytes, { flag: "wx" });
    await rename(temporary, path);
  } finally {
    await rm(temporary, { force: true });
  }
}

async function isSameExistingFile(
  left: string,
  right: string,
): Promise<boolean> {
  if (left === right) return true;
  try {
    const [leftStats, rightStats] = await Promise.all([
      stat(left),
      stat(right),
    ]);
    return leftStats.dev === rightStats.dev && leftStats.ino === rightStats.ino;
  } catch {
    return false;
  }
}

const buildSuccessVariant = {
  description: "Built PPTX deck",
  exitCode: 0,
  schema: builtSchema,
} as const;

const buildSuccessVariants = {
  built: buildSuccessVariant,
} as const;

const invalidOutputFailure = {
  description: "Output path overwrites deck source",
  exitCode: 2,
  schema: invalidOutputSchema,
} as const;

const sourceLoadFailedFailure = {
  description: "Failed to load deck source",
  exitCode: 1,
  schema: sourceLoadFailedSchema,
} as const;

const buildFailedFailure = {
  description: "Deck build failed with diagnostics",
  exitCode: 1,
  schema: buildFailedSchema,
} as const;

const writeFailedFailure = {
  description: "Failed to write PPTX output",
  exitCode: 1,
  schema: writeFailedSchema,
} as const;

const cancelledFailure = {
  description: "Deck build cancelled",
  exitCode: 130,
  schema: cancelledSchema,
} as const;

const buildFailureVariants = {
  invalidOutput: invalidOutputFailure,
  sourceLoadFailed: sourceLoadFailedFailure,
  buildFailed: buildFailedFailure,
  writeFailed: writeFailedFailure,
  cancelled: cancelledFailure,
} as const;

type BuildHandlerContext = {
  readonly input: z.infer<typeof buildInputSchema>;
  readonly dependencies: CliDependencies;
  readonly outcome: DataOutcome<"build", typeof buildSuccessVariants> &
    FailureOutcome<"build", typeof buildFailureVariants>;
};

async function buildDeckCommand({
  input,
  dependencies,
  outcome,
}: BuildHandlerContext) {
  if (dependencies.signal?.aborted) {
    return outcome.failure.cancelled({});
  }

  const entry = resolve(dependencies.cwd, input.entry);
  const output = resolve(
    dependencies.cwd,
    input.out ?? `${basename(entry, extname(entry))}.pptx`,
  );
  if (await isSameExistingFile(entry, output)) {
    return outcome.failure.invalidOutput({ entry, output });
  }
  const target =
    input.target === undefined ? undefined : targetByName(input.target);

  let result: Awaited<ReturnType<typeof buildFromSource>>;
  try {
    result = await buildFromSource(entry, {
      ...(input.export !== undefined && { exportName: input.export }),
      ...(target !== undefined && { target }),
      ...(dependencies.signal !== undefined && {
        signal: dependencies.signal,
      }),
      ...(input.embedFonts !== undefined && {
        embedFonts: input.embedFonts,
      }),
    });
  } catch (error) {
    if (dependencies.signal?.aborted && error === dependencies.signal.reason) {
      return outcome.failure.cancelled({});
    }
    if (error instanceof BuildError) {
      return outcome.failure.buildFailed({
        diagnostics: error.diagnostics.map(({ path, code, message }) => ({
          path,
          code,
          message,
        })),
      });
    }
    if (error instanceof SourceLoadError) {
      return outcome.failure.sourceLoadFailed({
        message: error.message,
      });
    }
    throw error;
  }

  try {
    await writeArtifact(output, result.pptx);
  } catch (writeError) {
    if (
      dependencies.signal?.aborted &&
      writeError === dependencies.signal.reason
    ) {
      return outcome.failure.cancelled({});
    }
    return outcome.failure.writeFailed({
      output,
      message: errorMessage(writeError),
    });
  }

  const warnings = result.report.diagnostics
    .filter((diagnostic) => diagnostic.severity === "warning")
    .map(({ path, code, message }) => ({ path, code, message }));

  return outcome.data.built({
    output,
    slides: result.report.slides.length,
    warnings,
  });
}

const define = defineCli<CliDependencies>();

const cliContract = define({
  root: "pptxdk",
  help: helpCapability({
    shortAlias: "-h",
    headings: {
      usage: "Usage",
      commands: "Commands",
      arguments: "Arguments",
      options: "Options",
      schema: "Schema",
    },
    wording: {
      commandPlaceholder: "command",
      choices: "(choices: {choices})",
      default: "(default: {value})",
    },
  }),
  version: versionCapability({
    value: text.line(cliVersion),
    shortAlias: "-V",
    description: "Print CLI version",
  }),
  schema: schemaCapability({
    description: text.line("Print machine-readable CLI contract schema"),
  }),
  usageFailureExitCode: 2,
  commands: {
    pptxdk: {
      kind: "rootGroup",
      name: "pptxdk",
      description: "Build PowerPoint decks from source modules.",
    },
    ...define.command("build")({
      kind: "command",
      parent: "pptxdk",
      name: "build",
      description: "Build a PowerPoint deck from a source module.",
      fields: {
        entry: {
          kind: "positional",
          description: "Deck source entrypoint",
        },
        out: {
          kind: "valueOption",
          longOption: "--out",
          shortAlias: "-o",
          description:
            "PPTX output path (defaults to <entry>.pptx in current directory)",
        },
        export: {
          kind: "valueOption",
          longOption: "--export",
          shortAlias: "-e",
          description:
            "Named deck or component exported from entry (defaults to default)",
        },
        target: {
          kind: "valueOption",
          longOption: "--target",
          description: "Target name exposed by library",
        },
        embedFonts: {
          kind: "valueOption",
          longOption: "--embed-fonts",
          description: "Embed full fonts or subset of used glyphs",
        },
      },
      input: buildInputSchema,
      outputs: define.outputModes("build", buildInputSchema, (mode) => ({
        default: "text",
        modes: {
          text: mode.atomic({
            description: "Human-readable text output",
            format: "text",
            success: {
              kind: "data",
              variants: {
                built: {
                  ...buildSuccessVariant,
                  text: ({ output, slides, warnings }) =>
                    text.lines(
                      toContractTextLines([
                        ...warnings.map(
                          (warning) =>
                            `Warning ${warning.path} [${warning.code}]: ${warning.message}`,
                        ),
                        `Built deck: ${output}`,
                        `Slides: ${slides}`,
                      ]),
                    ),
                },
              },
            },
            failures: {
              invalidOutput: {
                ...invalidOutputFailure,
                text: () =>
                  text.line("Output path cannot overwrite deck source"),
              },
              sourceLoadFailed: {
                ...sourceLoadFailedFailure,
                text: ({ message }) =>
                  text.lines(
                    toContractTextLines([
                      `Failed to load deck source: ${message}`,
                    ]),
                  ),
              },
              buildFailed: {
                ...buildFailedFailure,
                text: ({ diagnostics }) =>
                  text.lines(
                    toContractTextLines([
                      "Build failed:",
                      ...diagnostics.map(
                        (diagnostic) =>
                          `  ${diagnostic.path} [${diagnostic.code}]: ${diagnostic.message}`,
                      ),
                    ]),
                  ),
              },
              writeFailed: {
                ...writeFailedFailure,
                text: ({ output, message }) =>
                  text.lines(
                    toContractTextLines([
                      `Failed to write PPTX output to ${output}: ${message}`,
                    ]),
                  ),
              },
              cancelled: {
                ...cancelledFailure,
                text: () => text.line("Deck build cancelled"),
              },
            },
            handler: (context) => buildDeckCommand(context),
          }),
          structured: mode.atomic({
            description: "Structured JSON output",
            format: "structured",
            success: {
              kind: "data",
              variants: buildSuccessVariants,
            },
            failures: buildFailureVariants,
            handler: (context) => buildDeckCommand(context),
          }),
        },
      })),
    }),
  },
});

/**
 * Run a single CLI invocation and project its outcome through `runtime.write`.
 *
 * Returns the contract outcome exit code, `0` when `stdout` closes with `EPIPE`,
 * or `70` after attempting to report an unexpected runtime/output failure to `stderr`.
 */
export async function runCli(runtime: CliRuntime): Promise<number> {
  try {
    const result = await executeCli(cliContract, {
      invocation: parseCliInvocation(cliContract, runtime.argv.slice(2)),
      dependencies: {
        cwd: runtime.cwd ?? process.cwd(),
        ...(runtime.signal !== undefined && { signal: runtime.signal }),
      },
      write: runtime.write,
    });
    return result.exitCode;
  } catch (error) {
    if (
      error instanceof CliWriteError &&
      error.destination === "stdout" &&
      isNodeBrokenPipe(error)
    ) {
      return 0;
    }
    try {
      await runtime.write({
        destination: "stderr",
        chunk:
          "Command execution failed; please check the runtime environment or report this issue.\n",
      });
    } catch {
      // Diagnostic stream failure must not alter the unexpected failure exit code.
    }
    return 70;
  }
}
