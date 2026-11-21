import { randomUUID } from "node:crypto";
import { mkdir, rename, rm, writeFile } from "node:fs/promises";
import { basename, dirname, extname, resolve } from "node:path";

import {
  CliWriteError,
  defineCli,
  executeCli,
  helpCapability,
  outputCapability,
  parseCliInvocation,
  text,
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

const inputSchema = z.object({
  entry: z.string().min(1),
  out: z.string().min(1).optional(),
  export: z.string().min(1).optional(),
  target: z.enum(targetChoices).optional(),
  embedFonts: z.enum(["full", "subset"]).optional(),
});

const exportedSchema = z.object({
  output: z.string(),
  slides: z.number().int().nonnegative(),
  warnings: z.array(diagnosticItemSchema),
});

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

const cliContract = defineCli<CliDependencies>()({
  root: "reactPptx",
  help: helpCapability({
    shortAlias: "-h",
    headings: {
      usage: "Usage",
      arguments: "Arguments",
      options: "Options",
    },
    wording: {
      choices: "(choices: {choices})",
      default: "(default: {value})",
    },
  }),
  output: outputCapability({ defaultFormat: "text" }),
  usageFailureExitCode: 2,
  commands: {
    reactPptx: {
      kind: "rootCommand",
      name: "react-pptx",
      description: "Build a PowerPoint deck from a React source module.",
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
            "Named component exported from entry (defaults to default)",
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
      input: inputSchema,
      success: {
        kind: "data",
        variants: {
          exported: {
            description: "Exported PPTX deck",
            exitCode: 0,
            schema: exportedSchema,
            text: ({ output, slides, warnings }) =>
              text.lines(
                toContractTextLines([
                  ...warnings.map(
                    (warning) =>
                      `Warning ${warning.path} [${warning.code}]: ${warning.message}`,
                  ),
                  `Exported deck: ${output}`,
                  `Slides: ${slides}`,
                ]),
              ),
          },
        },
      },
      failures: {
        invalidOutput: {
          description: "Output path overwrites deck source",
          exitCode: 2,
          schema: z.object({
            entry: z.string(),
            output: z.string(),
          }),
          text: () => text.line("Output path cannot overwrite deck source"),
        },
        cancelled: {
          description: "Deck export cancelled",
          exitCode: 130,
          schema: z.object({}),
          text: () => text.line("Deck export cancelled"),
        },
        buildFailed: {
          description: "Deck build failed with diagnostics",
          exitCode: 1,
          schema: z.object({
            diagnostics: z.array(diagnosticItemSchema),
          }),
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
        sourceLoadFailed: {
          description: "Failed to load deck source",
          exitCode: 1,
          schema: z.object({
            message: z.string(),
          }),
          text: ({ message }) =>
            text.lines(
              toContractTextLines([`Failed to load deck source: ${message}`]),
            ),
        },
        exportFailed: {
          description: "Deck export failed",
          exitCode: 1,
          schema: z.object({
            message: z.string(),
          }),
          text: ({ message }) =>
            text.lines(toContractTextLines([`Deck export failed: ${message}`])),
        },
      },
      async handler({ input, dependencies, outcome }) {
        if (dependencies.signal?.aborted) {
          return outcome.failure.cancelled({});
        }

        const entry = resolve(dependencies.cwd, input.entry);
        const output = resolve(
          dependencies.cwd,
          input.out ?? `${basename(entry, extname(entry))}.pptx`,
        );
        if (entry === output) {
          return outcome.failure.invalidOutput({ entry, output });
        }

        const target =
          input.target === undefined ? undefined : targetByName(input.target);

        try {
          const result = await buildFromSource(entry, {
            ...(input.export !== undefined && { exportName: input.export }),
            ...(target !== undefined && { target }),
            ...(dependencies.signal !== undefined && {
              signal: dependencies.signal,
            }),
            ...(input.embedFonts !== undefined && {
              embedFonts: input.embedFonts,
            }),
          });
          await writeArtifact(output, result.pptx);

          const warnings = result.report.diagnostics
            .filter((diagnostic) => diagnostic.severity === "warning")
            .map(({ path, code, message }) => ({ path, code, message }));

          return outcome.data.exported({
            output,
            slides: result.report.slides.length,
            warnings,
          });
        } catch (error) {
          if (
            dependencies.signal?.aborted &&
            error === dependencies.signal.reason
          ) {
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
          return outcome.failure.exportFailed({
            message: errorMessage(error),
          });
        }
      },
    },
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
