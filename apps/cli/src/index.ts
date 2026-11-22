import { randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import { lstat, mkdir, rename, rm, stat, writeFile } from "node:fs/promises";
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
import {
  BuildError,
  fontEmbeddingModeSchema,
  targetNameSchema,
  targets,
  type Diagnostic,
  type Target,
} from "pptxdk";
import {
  buildFromSource,
  previewFromSource,
  SourceLoadError,
} from "pptxdk/node";
import { PreviewPageRangeError, previewPageNumberSchema } from "pptxdk/preview";
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
const { version: cliVersion } = z
  .object({ version: z.string().min(1) })
  .parse(
    JSON.parse(
      readFileSync(new URL("../package.json", import.meta.url), "utf8"),
    ),
  );

const diagnosticItemSchema = z.object({
  path: z.string(),
  code: z.string(),
  message: z.string(),
});

const previewPageItemSchema = z.object({
  page: previewPageNumberSchema,
  id: z.string(),
  output: z.string(),
});

const previewDiagnosticItemSchema = z.object({
  page: previewPageNumberSchema,
  path: z.string(),
  message: z.string(),
});

export const buildInputSchema = z.object({
  entry: z.string().min(1),
  out: z.string().min(1).optional(),
  export: z.string().min(1).optional(),
  target: targetNameSchema.optional(),
  embedFonts: fontEmbeddingModeSchema.optional(),
});
export type BuildInput = z.infer<typeof buildInputSchema>;

function parsePagesSpec(raw: string): number[] | undefined {
  const parts = raw.split(",");
  if (parts.length === 0) return undefined;

  const selected = new Set<number>();
  for (const part of parts) {
    // Reject empty items ("", "1,", ",1", "1,,2"), whitespace, signs, decimals, or non-digits.
    if (/^\d+$/u.test(part)) {
      const parsed = previewPageNumberSchema.safeParse(Number(part));
      if (!parsed.success) return undefined;
      selected.add(parsed.data);
      continue;
    }
    const rangeMatch = /^(\d+)-(\d+)$/u.exec(part);
    if (rangeMatch !== null) {
      const start = previewPageNumberSchema.safeParse(Number(rangeMatch[1]));
      const end = previewPageNumberSchema.safeParse(Number(rangeMatch[2]));
      if (!start.success || !end.success || start.data > end.data) {
        return undefined;
      }
      for (let page = start.data; page <= end.data; page += 1) {
        selected.add(page);
      }
      continue;
    }
    return undefined;
  }
  return [...selected].toSorted((a, b) => a - b);
}

const pagesOptionSchema = z.codec(
  z.string().min(1),
  z.array(previewPageNumberSchema).min(1),
  {
    decode(raw, ctx) {
      const parsed = parsePagesSpec(raw);
      if (parsed === undefined) {
        ctx.issues.push({
          code: "custom",
          input: raw,
          message:
            "Pages must be comma-separated 1-based page numbers or ascending closed ranges (for example, 1,3-5)",
        });
        return z.NEVER;
      }
      return parsed;
    },
    encode(pages) {
      return pages.join(",");
    },
  },
);

export const previewInputSchema = z.object({
  entry: z.string().min(1),
  outDir: z.string().min(1).optional(),
  export: z.string().min(1).optional(),
  pages: pagesOptionSchema.optional(),
});
export type PreviewInput = z.infer<typeof previewInputSchema>;

export const builtSchema = z.object({
  output: z.string(),
  slides: z.number().int().nonnegative(),
  warnings: z.array(diagnosticItemSchema),
});
export type BuiltOutput = z.infer<typeof builtSchema>;

export const previewedSchema = z.object({
  outputDir: z.string(),
  pages: z.array(previewPageItemSchema),
  warnings: z.array(diagnosticItemSchema),
});
export type PreviewedOutput = z.infer<typeof previewedSchema>;

const invalidOutputSchema = z.object({
  entry: z.string(),
  output: z.string(),
});

const pageOutOfRangeSchema = z.object({
  requestedPages: z.array(previewPageNumberSchema),
  totalSlides: z.number().int().nonnegative(),
});

const sourceLoadFailedSchema = z.object({
  message: z.string(),
});

const buildFailedSchema = z.object({
  diagnostics: z.array(diagnosticItemSchema),
});

const previewFailedSchema = z.object({
  outputDir: z.string(),
  pages: z.array(previewPageItemSchema),
  warnings: z.array(diagnosticItemSchema),
  diagnostics: z.array(previewDiagnosticItemSchema),
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

function errnoCode(error: unknown): string | undefined {
  return typeof error === "object" &&
    error !== null &&
    "code" in error &&
    typeof (error as { readonly code?: unknown }).code === "string"
    ? (error as { readonly code: string }).code
    : undefined;
}

function isCancelledError(error: unknown, signal?: AbortSignal): boolean {
  return Boolean(signal?.aborted && error === signal.reason);
}

function extractWarnings(diagnostics: readonly Diagnostic[]) {
  return diagnostics
    .filter((diagnostic) => diagnostic.severity === "warning")
    .map(({ path, code, message }) => ({ path, code, message }));
}

function toBuildFailureDiagnostics(diagnostics: readonly Diagnostic[]) {
  return diagnostics.map(({ path, code, message }) => ({
    path,
    code,
    message,
  }));
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

type PreviewOutputDirValidation =
  | { readonly kind: "valid" }
  | { readonly kind: "invalidOutput" }
  | { readonly kind: "error"; readonly error: unknown };

// Validate outputDir read-only before previewFromSource so source or build failures never create an empty directory.
async function validatePreviewOutputDir(
  entry: string,
  outDir: string,
): Promise<PreviewOutputDirValidation> {
  if (outDir === entry) return { kind: "invalidOutput" };
  try {
    const stats = await stat(outDir);
    return stats.isDirectory() ? { kind: "valid" } : { kind: "invalidOutput" };
  } catch (error) {
    const code = errnoCode(error);
    if (code === "ENOTDIR" || code === "ELOOP") {
      return { kind: "invalidOutput" };
    }
    if (code !== "ENOENT") {
      return { kind: "error", error };
    }
    try {
      await lstat(outDir);
      return { kind: "invalidOutput" };
    } catch (lstatError) {
      const lstatCode = errnoCode(lstatError);
      if (lstatCode === "ENOENT") return { kind: "valid" };
      if (lstatCode === "ENOTDIR" || lstatCode === "ELOOP") {
        return { kind: "invalidOutput" };
      }
      return { kind: "error", error: lstatError };
    }
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

const previewSuccessVariant = {
  description: "Previewed deck pages as PNG images",
  exitCode: 0,
  schema: previewedSchema,
} as const;

const previewSuccessVariants = {
  previewed: previewSuccessVariant,
} as const;

const previewInvalidOutputFailure = {
  description: "Output directory is invalid or points to a non-directory path",
  exitCode: 2,
  schema: invalidOutputSchema,
} as const;

const previewPageOutOfRangeFailure = {
  description: "Requested preview pages exceed total slide count",
  exitCode: 2,
  schema: pageOutOfRangeSchema,
} as const;

const previewFailedFailure = {
  description: "Deck preview failed on one or more pages",
  exitCode: 1,
  schema: previewFailedSchema,
} as const;

const previewWriteFailedFailure = {
  description: "Failed to write preview image",
  exitCode: 1,
  schema: writeFailedSchema,
} as const;

const previewCancelledFailure = {
  description: "Deck preview cancelled",
  exitCode: 130,
  schema: cancelledSchema,
} as const;

const previewFailureVariants = {
  invalidOutput: previewInvalidOutputFailure,
  pageOutOfRange: previewPageOutOfRangeFailure,
  sourceLoadFailed: sourceLoadFailedFailure,
  buildFailed: buildFailedFailure,
  previewFailed: previewFailedFailure,
  writeFailed: previewWriteFailedFailure,
  cancelled: previewCancelledFailure,
} as const;

type BuildHandlerContext = {
  readonly input: z.infer<typeof buildInputSchema>;
  readonly dependencies: CliDependencies;
  readonly outcome: DataOutcome<"build", typeof buildSuccessVariants> &
    FailureOutcome<"build", typeof buildFailureVariants>;
};

type PreviewHandlerContext = {
  readonly input: z.infer<typeof previewInputSchema>;
  readonly dependencies: CliDependencies;
  readonly outcome: DataOutcome<"preview", typeof previewSuccessVariants> &
    FailureOutcome<"preview", typeof previewFailureVariants>;
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
    if (isCancelledError(error, dependencies.signal)) {
      return outcome.failure.cancelled({});
    }
    if (error instanceof BuildError) {
      return outcome.failure.buildFailed({
        diagnostics: toBuildFailureDiagnostics(error.diagnostics),
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
    if (isCancelledError(writeError, dependencies.signal)) {
      return outcome.failure.cancelled({});
    }
    return outcome.failure.writeFailed({
      output,
      message: errorMessage(writeError),
    });
  }

  return outcome.data.built({
    output,
    slides: result.report.slides.length,
    warnings: extractWarnings(result.report.diagnostics),
  });
}

async function previewDeckCommand({
  input,
  dependencies,
  outcome,
}: PreviewHandlerContext) {
  if (dependencies.signal?.aborted) {
    return outcome.failure.cancelled({});
  }

  const entry = resolve(dependencies.cwd, input.entry);
  const outputDir = resolve(
    dependencies.cwd,
    input.outDir ?? `${basename(entry, extname(entry))}-preview`,
  );
  const outputDirCheck = await validatePreviewOutputDir(entry, outputDir);
  if (outputDirCheck.kind === "invalidOutput") {
    return outcome.failure.invalidOutput({ entry, output: outputDir });
  }
  if (outputDirCheck.kind === "error") {
    return outcome.failure.writeFailed({
      output: outputDir,
      message: errorMessage(outputDirCheck.error),
    });
  }

  let result: Awaited<ReturnType<typeof previewFromSource>>;
  try {
    result = await previewFromSource(entry, {
      ...(input.export !== undefined && { exportName: input.export }),
      ...(input.pages !== undefined && { pages: input.pages }),
      ...(dependencies.signal !== undefined && {
        signal: dependencies.signal,
      }),
    });
  } catch (error) {
    if (isCancelledError(error, dependencies.signal)) {
      return outcome.failure.cancelled({});
    }
    if (error instanceof BuildError) {
      return outcome.failure.buildFailed({
        diagnostics: toBuildFailureDiagnostics(error.diagnostics),
      });
    }
    if (input.pages !== undefined && error instanceof PreviewPageRangeError) {
      return outcome.failure.pageOutOfRange({
        requestedPages: input.pages,
        totalSlides: error.totalSlides,
      });
    }
    if (error instanceof SourceLoadError) {
      return outcome.failure.sourceLoadFailed({
        message: error.message,
      });
    }
    throw error;
  }

  const warnings = extractWarnings(result.report.diagnostics);
  // Base zero-padding width on total slides in the resolved deck rather than painted pages.
  const padWidth = Math.max(2, String(result.report.slides.length).length);
  const writtenPages: z.infer<typeof previewPageItemSchema>[] = [];
  const failedDiagnostics: z.infer<typeof previewDiagnosticItemSchema>[] = [];

  for (const pageResult of result.pages) {
    if (dependencies.signal?.aborted) {
      return outcome.failure.cancelled({});
    }

    const errors = pageResult.diagnostics.filter(
      (diagnostic) => diagnostic.severity === "error",
    );
    if (errors.length > 0 || pageResult.png === undefined) {
      if (errors.length > 0) {
        failedDiagnostics.push(
          ...errors.map(({ page, path, message }) => ({
            page,
            path,
            message,
          })),
        );
      } else {
        failedDiagnostics.push({
          page: pageResult.page,
          path: `deck/slide[${pageResult.id}]`,
          message: "Preview did not produce a PNG image",
        });
      }
      continue;
    }

    const pageOutput = resolve(
      outputDir,
      `page-${String(pageResult.page).padStart(padWidth, "0")}.png`,
    );
    try {
      await writeArtifact(pageOutput, pageResult.png);
    } catch (writeError) {
      if (isCancelledError(writeError, dependencies.signal)) {
        return outcome.failure.cancelled({});
      }
      return outcome.failure.writeFailed({
        output: pageOutput,
        message: errorMessage(writeError),
      });
    }

    writtenPages.push({
      page: pageResult.page,
      id: pageResult.id,
      output: pageOutput,
    });
  }

  if (failedDiagnostics.length > 0) {
    return outcome.failure.previewFailed({
      outputDir,
      pages: writtenPages,
      warnings,
      diagnostics: failedDiagnostics,
    });
  }

  return outcome.data.previewed({
    outputDir,
    pages: writtenPages,
    warnings,
  });
}

const sharedTextFailures = {
  sourceLoadFailed: {
    ...sourceLoadFailedFailure,
    text: ({ message }: z.infer<typeof sourceLoadFailedSchema>) =>
      text.lines(
        toContractTextLines([`Failed to load deck source: ${message}`]),
      ),
  },
  buildFailed: {
    ...buildFailedFailure,
    text: ({ diagnostics }: z.infer<typeof buildFailedSchema>) =>
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
} as const;

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
      description: "Build and preview PowerPoint decks from source modules.",
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
              ...sharedTextFailures,
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
    ...define.command("preview")({
      kind: "command",
      parent: "pptxdk",
      name: "preview",
      description: "Render static PNG page previews from a source module.",
      fields: {
        entry: {
          kind: "positional",
          description: "Deck source entrypoint",
        },
        outDir: {
          kind: "valueOption",
          longOption: "--out-dir",
          shortAlias: "-o",
          description:
            "Directory for preview PNGs (defaults to <entry>-preview in current directory)",
        },
        export: {
          kind: "valueOption",
          longOption: "--export",
          shortAlias: "-e",
          description:
            "Named deck or component exported from entry (defaults to default)",
        },
        pages: {
          kind: "valueOption",
          longOption: "--pages",
          shortAlias: "-p",
          description:
            "1-based page numbers or closed ranges to preview (for example, 1,3-5)",
        },
      },
      input: previewInputSchema,
      outputs: define.outputModes("preview", previewInputSchema, (mode) => ({
        default: "text",
        modes: {
          text: mode.atomic({
            description: "Human-readable text output",
            format: "text",
            success: {
              kind: "data",
              variants: {
                previewed: {
                  ...previewSuccessVariant,
                  text: ({ outputDir, pages, warnings }) =>
                    text.lines(
                      toContractTextLines([
                        ...warnings.map(
                          (warning) =>
                            `Warning ${warning.path} [${warning.code}]: ${warning.message}`,
                        ),
                        `Previewed ${pages.length} page(s) in ${outputDir}`,
                        ...pages.map(
                          (page) =>
                            `  Page ${page.page} (${page.id}): ${page.output}`,
                        ),
                      ]),
                    ),
                },
              },
            },
            failures: {
              invalidOutput: {
                ...previewInvalidOutputFailure,
                text: ({ output }) =>
                  text.lines(
                    toContractTextLines([
                      `Output directory cannot be a non-directory path: ${output}`,
                    ]),
                  ),
              },
              pageOutOfRange: {
                ...previewPageOutOfRangeFailure,
                text: ({ requestedPages, totalSlides }) =>
                  text.lines(
                    toContractTextLines([
                      `Requested pages (${requestedPages.join(", ")}) exceed total slides (${totalSlides})`,
                    ]),
                  ),
              },
              ...sharedTextFailures,
              previewFailed: {
                ...previewFailedFailure,
                text: ({ pages, warnings, diagnostics }) =>
                  text.lines(
                    toContractTextLines([
                      ...warnings.map(
                        (warning) =>
                          `Warning ${warning.path} [${warning.code}]: ${warning.message}`,
                      ),
                      "Preview failed:",
                      ...diagnostics.map(
                        (diagnostic) =>
                          `  Page ${diagnostic.page} ${diagnostic.path}: ${diagnostic.message}`,
                      ),
                      ...(pages.length > 0
                        ? [
                            "Written pages:",
                            ...pages.map(
                              (page) =>
                                `  Page ${page.page} (${page.id}): ${page.output}`,
                            ),
                          ]
                        : []),
                    ]),
                  ),
              },
              writeFailed: {
                ...previewWriteFailedFailure,
                text: ({ output, message }) =>
                  text.lines(
                    toContractTextLines([
                      `Failed to write preview image to ${output}: ${message}`,
                    ]),
                  ),
              },
              cancelled: {
                ...previewCancelledFailure,
                text: () => text.line("Deck preview cancelled"),
              },
            },
            handler: (context) => previewDeckCommand(context),
          }),
          structured: mode.atomic({
            description: "Structured JSON output",
            format: "structured",
            success: {
              kind: "data",
              variants: previewSuccessVariants,
            },
            failures: previewFailureVariants,
            handler: (context) => previewDeckCommand(context),
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
