import { randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import { lstat, mkdir, rename, rm, stat, writeFile } from "node:fs/promises";
import { findPackageJSON } from "node:module";
import {
  basename,
  dirname,
  extname,
  posix,
  relative,
  resolve,
} from "node:path";

import {
  CliWriteError,
  defineCli,
  executeCli,
  helpCapability,
  parseCliInvocation,
  schemaCapability,
  text,
  versionCapability,
  type WriteCliOutput,
} from "@boringparts/cli-contract";
import { isNodeBrokenPipe } from "@boringparts/cli-contract/node";
import {
  BuildError,
  CodegenError,
  ImportError,
  buildDiagnosticSchema,
  codegen,
  codegenDiagnosticSchema,
  codegenTargetSchema,
  fatalImportIssueSchema,
  fontEmbeddingModeSchema,
  importDiagnosticSchema,
  previewPlaceholderSchema,
  targetNameSchema,
  targets,
  type CodegenFileContent,
  type FatalImportIssue,
  type ImportDiagnostic,
  type ImportLocation,
  type ImportResult,
  type PreviewPlaceholder,
  type Target,
} from "@officedk/pptxdk";
import {
  buildFromSource,
  importPptxFile,
  previewFromSource,
  SourceLoadError,
} from "@officedk/pptxdk/node";
import {
  PreviewPageRangeError,
  previewPageNumberSchema,
} from "@officedk/pptxdk/preview";
import { previewErrorOptionSchemas } from "@pptxdk/core/build";
import { strToU8, zipSync } from "fflate";
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
  readonly codegen?: typeof codegen | undefined;
};

// Source and bundled modules have different depths. Both report the installed SDK version.
const { version: cliVersion } = z
  .object({ version: z.string().min(1) })
  .parse(
    JSON.parse(readFileSync(findPackageJSON("./", import.meta.url)!, "utf8")),
  );

const previewPageItemSchema = z.object({
  page: previewPageNumberSchema,
  id: z.string(),
  output: z.string(),
});

const previewErrorWithPageSchema = z.discriminatedUnion("code", [
  previewErrorOptionSchemas[0].extend({ page: previewPageNumberSchema }),
  previewErrorOptionSchemas[1].extend({ page: previewPageNumberSchema }),
  previewErrorOptionSchemas[2].extend({ page: previewPageNumberSchema }),
  previewErrorOptionSchemas[3].extend({ page: previewPageNumberSchema }),
  previewErrorOptionSchemas[4].extend({ page: previewPageNumberSchema }),
]);

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

export const codegenInputSchema = z.object({
  entry: z.string().min(1),
  target: codegenTargetSchema,
  outDir: z.string().min(1).optional(),
  outZip: z.string().min(1).optional(),
});
export type CodegenInput = z.infer<typeof codegenInputSchema>;

export const builtSchema = z.object({
  output: z.string(),
  slides: z.number().int().nonnegative(),
});
export type BuiltOutput = z.infer<typeof builtSchema>;

export const previewedSchema = z.object({
  outputDir: z.string(),
  pages: z.array(previewPageItemSchema),
  placeholders: z.array(previewPlaceholderSchema),
});
export type PreviewedOutput = z.infer<typeof previewedSchema>;

export const generatedSourceSchema = z.object({
  target: codegenTargetSchema,
  delivery: z.enum(["directory", "zip"]),
  output: z.string(),
  entry: z.string(),
  files: z.array(z.string()),
  slides: z.number().int().nonnegative(),
  diagnostics: z.array(importDiagnosticSchema),
});
export type GeneratedSourceOutput = z.infer<typeof generatedSourceSchema>;

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
  diagnostics: z.array(buildDiagnosticSchema),
});

const previewFailedSchema = z.object({
  outputDir: z.string(),
  pages: z.array(previewPageItemSchema),
  placeholders: z.array(previewPlaceholderSchema),
  errors: z.array(previewErrorWithPageSchema),
});

const inputReadFailedSchema = z.object({
  entry: z.string(),
  message: z.string(),
});

const importFailedSchema = z.object({
  entry: z.string(),
  issue: fatalImportIssueSchema,
  message: z.string(),
});

const codegenFailedSchema = z.object({
  entry: z.string(),
  diagnostics: z.array(codegenDiagnosticSchema),
  importDiagnostics: z.array(importDiagnosticSchema),
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
  if (!signal?.aborted) return false;
  return (
    error === signal.reason ||
    errnoCode(error) === "ABORT_ERR" ||
    (error instanceof Error && error.name === "AbortError")
  );
}

/**
 * Normalize dynamic text into the non-empty, single-line, NUL-free tuple
 * required by `@boringparts/cli-contract`'s `text.lines(...)` invariant.
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

type OutputDirValidation =
  | { readonly kind: "valid" }
  | { readonly kind: "invalidOutput" }
  | { readonly kind: "error"; readonly error: unknown };

// Validate outputDir read-only before running source/import work so upstream failures never create an empty directory.
async function validateOutputDir(
  entry: string,
  outDir: string,
): Promise<OutputDirValidation> {
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

function formatPreviewPlaceholderLines(
  placeholders: readonly PreviewPlaceholder[],
): string[] {
  return placeholders.map(
    (placeholder) => `  Placeholder ${placeholder.path}: ${placeholder.url}`,
  );
}

function presentSourceLoadFailed({
  message,
}: z.infer<typeof sourceLoadFailedSchema>) {
  return text.lines(
    toContractTextLines([`Failed to load deck source: ${message}`]),
  );
}

function formatBuildDiagnostic(
  diagnostic: z.infer<typeof buildDiagnosticSchema>,
): string {
  const fieldSuffix = diagnostic.field ? ` (${diagnostic.field})` : "";
  const base = `  ${diagnostic.path}${fieldSuffix} [${diagnostic.code}]: ${diagnostic.message}`;
  const details: string[] = [];
  if ("font" in diagnostic) {
    const fontIdentity = [
      diagnostic.font.family,
      diagnostic.font.weight,
      diagnostic.font.style,
    ]
      .filter((part) => part !== undefined)
      .join(" ");
    if (fontIdentity.length > 0) {
      details.push(`font: ${fontIdentity}`);
    }
  }
  if ("characters" in diagnostic && diagnostic.characters.length > 0) {
    details.push(`missing characters: ${diagnostic.characters.join(", ")}`);
  }
  return details.length > 0 ? `${base} (${details.join("; ")})` : base;
}

function presentBuildFailed({
  diagnostics,
}: z.infer<typeof buildFailedSchema>) {
  return text.lines(
    toContractTextLines([
      "Build failed:",
      ...diagnostics.map(formatBuildDiagnostic),
    ]),
  );
}

// Fixed timestamp guarantees byte-identical ZIP output across repeated invocations for the same input.
const fixedZipMtime = new Date("2020-01-01T00:00:00Z");

type NormalizedCodegenFile = {
  readonly path: string;
  readonly bytes: Uint8Array;
};

function normalizeRelativePosixPath(rawPath: string): string | undefined {
  if (
    rawPath.length === 0 ||
    rawPath.includes("\\") ||
    rawPath.includes("\0") ||
    rawPath.startsWith("/")
  ) {
    return undefined;
  }
  const normalized = posix.normalize(rawPath);
  if (
    normalized === "." ||
    normalized === ".." ||
    normalized.startsWith("../") ||
    posix.isAbsolute(normalized)
  ) {
    return undefined;
  }
  return normalized;
}

function normalizeCodegenFiles(
  files: Readonly<Record<string, CodegenFileContent>>,
): readonly NormalizedCodegenFile[] {
  const normalizedEntries: NormalizedCodegenFile[] = [];
  for (const [rawPath, content] of Object.entries(files)) {
    const normalizedPath = normalizeRelativePosixPath(rawPath);
    if (normalizedPath === undefined) {
      throw new Error(`Invalid generated relative file path: ${rawPath}`);
    }
    normalizedEntries.push({
      path: normalizedPath,
      bytes: typeof content === "string" ? strToU8(content) : content,
    });
  }
  return normalizedEntries.toSorted((left, right) =>
    left.path < right.path ? -1 : left.path > right.path ? 1 : 0,
  );
}

function isWithinDirectory(directory: string, candidate: string): boolean {
  const rel = relative(directory, candidate);
  return (
    rel.length > 0 &&
    rel !== ".." &&
    !rel.startsWith("../") &&
    !rel.startsWith("..\\")
  );
}

function formatImportLocation(location: ImportLocation): string {
  const details: string[] = [location.part];
  if (location.slideIndex !== undefined) {
    details.push(`slide ${location.slideIndex + 1}`);
  }
  if (location.objectName !== undefined && location.objectName.length > 0) {
    details.push(`object ${JSON.stringify(location.objectName)}`);
  }
  if (location.xpath !== undefined) {
    details.push(`at ${location.xpath}`);
  }
  return details.join(" ");
}

function formatImportDiagnostic(diagnostic: ImportDiagnostic): string {
  const locationText = formatImportLocation(diagnostic.location);
  const extras: string[] = [];
  switch (diagnostic.code) {
    case "corrupt-part":
      break;
    case "missing-target-part":
      extras.push(
        `relationship=${diagnostic.relationshipId}`,
        `target=${diagnostic.targetPart}`,
      );
      break;
    case "broken-relationship":
      extras.push(`relationship=${diagnostic.relationshipId}`);
      break;
    case "invalid-xml-value":
      extras.push(`attributeOrElement=${diagnostic.attributeOrElement}`);
      if (diagnostic.rawValue !== undefined) {
        extras.push(`value=${JSON.stringify(diagnostic.rawValue)}`);
      }
      break;
    case "unsupported-mechanism":
      extras.push(`element=${diagnostic.namespaceUri}:${diagnostic.localName}`);
      if (diagnostic.mechanismHint !== undefined) {
        extras.push(`hint=${diagnostic.mechanismHint}`);
      }
      if (diagnostic.loss !== undefined) {
        extras.push(
          `loss=${diagnostic.loss.lost}->${diagnostic.loss.retainedAs}`,
        );
      }
      break;
    case "compatibility-fallback":
      extras.push(
        `preferred=${diagnostic.preferredMechanism}`,
        `fallback=${diagnostic.selectedFallback}`,
        `loss=${diagnostic.loss.lost}->${diagnostic.loss.retainedAs}`,
      );
      break;
    case "unconsumed-xml":
      extras.push(
        `xml=${diagnostic.kind}:${diagnostic.namespaceUri}:${diagnostic.localName}`,
      );
      break;
  }
  if ("cascaded" in diagnostic && diagnostic.cascaded !== undefined) {
    extras.push(`cascaded=${diagnostic.cascaded.length}`);
  }
  const detailSuffix = extras.length > 0 ? ` (${extras.join(", ")})` : "";
  return `Import diagnostic [${diagnostic.code}] (${diagnostic.disposition}) ${locationText}${detailSuffix}`;
}

function formatFatalImportIssue(issue: FatalImportIssue): string {
  switch (issue.code) {
    case "corrupt-archive":
      return "Corrupt PPTX archive";
    case "missing-entry": {
      const expected =
        issue.expectedRelationshipType === undefined
          ? ""
          : ` (expected ${issue.expectedRelationshipType})`;
      return `Missing entry at ${formatImportLocation(issue.location)}${expected}`;
    }
    case "malformed-xml": {
      const detail = issue.detail === undefined ? "" : `: ${issue.detail}`;
      return `Malformed XML at ${formatImportLocation(issue.location)}${detail}`;
    }
    case "broken-entry-relationship": {
      const details = [
        issue.relationshipId === undefined
          ? undefined
          : `relationship=${issue.relationshipId}`,
        issue.target === undefined ? undefined : `target=${issue.target}`,
      ].filter((item) => item !== undefined);
      const suffix = details.length > 0 ? ` (${details.join(", ")})` : "";
      return `Broken entry relationship at ${formatImportLocation(issue.location)}${suffix}`;
    }
    case "invalid-presentation-root": {
      const rootName = [issue.namespaceUri, issue.localName]
        .filter((item) => item !== undefined)
        .join(":");
      const suffix = rootName.length > 0 ? ` (${rootName})` : "";
      return `Invalid presentation root at ${formatImportLocation(issue.location)}${suffix}`;
    }
    case "unsupported-format-baseline":
      return `Unsupported format baseline ${issue.baseline} (${issue.detectedNamespace}) at ${formatImportLocation(issue.location)}`;
  }
}

const define = defineCli<CliDependencies>();

const buildCommand = define.command("build")({
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
  outputs: ({ atomic }) => {
    const buildDeck = atomic({
      success: {
        kind: "data",
        variants: {
          built: {
            description: "Built PPTX deck",
            exitCode: 0,
            schema: builtSchema,
          },
        },
      },
      failures: {
        invalidOutput: {
          description: "Output path overwrites deck source",
          exitCode: 2,
          schema: invalidOutputSchema,
        },
        sourceLoadFailed: {
          description: "Failed to load deck source",
          exitCode: 1,
          schema: sourceLoadFailedSchema,
        },
        buildFailed: {
          description: "Deck build failed with diagnostics",
          exitCode: 1,
          schema: buildFailedSchema,
        },
        writeFailed: {
          description: "Failed to write PPTX output",
          exitCode: 1,
          schema: writeFailedSchema,
        },
        cancelled: {
          description: "Deck build cancelled",
          exitCode: 130,
          schema: cancelledSchema,
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
              diagnostics: [...error.diagnostics],
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
        });
      },
    });

    return {
      default: "text",
      modes: {
        text: buildDeck.text({
          description: "Human-readable text output",
          success: {
            built: ({ output, slides }) =>
              text.lines(
                toContractTextLines([
                  `Built deck: ${output}`,
                  `Slides: ${slides}`,
                ]),
              ),
          },
          failures: {
            invalidOutput: () =>
              text.line("Output path cannot overwrite deck source"),
            sourceLoadFailed: presentSourceLoadFailed,
            buildFailed: presentBuildFailed,
            writeFailed: ({ output, message }) =>
              text.lines(
                toContractTextLines([
                  `Failed to write PPTX output to ${output}: ${message}`,
                ]),
              ),
            cancelled: () => text.line("Deck build cancelled"),
          },
        }),
        structured: buildDeck.structured({
          description: "Structured JSON output",
        }),
      },
    };
  },
});

const previewCommand = define.command("preview")({
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
  outputs: ({ atomic }) => {
    const previewDeck = atomic({
      success: {
        kind: "data",
        variants: {
          previewed: {
            description: "Previewed deck pages as PNG images",
            exitCode: 0,
            schema: previewedSchema,
          },
        },
      },
      failures: {
        invalidOutput: {
          description:
            "Output directory is invalid or points to a non-directory path",
          exitCode: 2,
          schema: invalidOutputSchema,
        },
        pageOutOfRange: {
          description: "Requested preview pages exceed total slide count",
          exitCode: 2,
          schema: pageOutOfRangeSchema,
        },
        sourceLoadFailed: {
          description: "Failed to load deck source",
          exitCode: 1,
          schema: sourceLoadFailedSchema,
        },
        buildFailed: {
          description: "Deck build failed with diagnostics",
          exitCode: 1,
          schema: buildFailedSchema,
        },
        previewFailed: {
          description: "Deck preview failed on one or more pages",
          exitCode: 1,
          schema: previewFailedSchema,
        },
        writeFailed: {
          description: "Failed to write preview image",
          exitCode: 1,
          schema: writeFailedSchema,
        },
        cancelled: {
          description: "Deck preview cancelled",
          exitCode: 130,
          schema: cancelledSchema,
        },
      },
      async handler({ input, dependencies, outcome }) {
        if (dependencies.signal?.aborted) {
          return outcome.failure.cancelled({});
        }

        const entry = resolve(dependencies.cwd, input.entry);
        const outputDir = resolve(
          dependencies.cwd,
          input.outDir ?? `${basename(entry, extname(entry))}-preview`,
        );
        const outputDirCheck = await validateOutputDir(entry, outputDir);
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
              diagnostics: [...error.diagnostics],
            });
          }
          if (
            input.pages !== undefined &&
            error instanceof PreviewPageRangeError
          ) {
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

        // Base zero-padding width on total slides in the resolved deck rather than painted pages.
        const padWidth = Math.max(
          2,
          String(result.report.slides.length).length,
        );
        const writtenPages: z.infer<typeof previewPageItemSchema>[] = [];
        const placeholders: PreviewPlaceholder[] = [];
        const failedErrors: z.infer<typeof previewErrorWithPageSchema>[] = [];

        for (const pageResult of result.pages) {
          if (dependencies.signal?.aborted) {
            return outcome.failure.cancelled({});
          }

          if (pageResult.errors.length > 0 || pageResult.png === undefined) {
            if (pageResult.errors.length > 0) {
              failedErrors.push(
                ...pageResult.errors.map((error) => ({
                  ...error,
                  page: pageResult.page,
                })),
              );
            } else {
              failedErrors.push({
                code: "preview-unsupported-capability",
                page: pageResult.page,
                path: `deck/slide[${pageResult.id}]`,
                message: "Preview did not produce a PNG image",
                capability: "preview",
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
          placeholders.push(...pageResult.placeholders);
        }

        if (failedErrors.length > 0) {
          return outcome.failure.previewFailed({
            outputDir,
            pages: writtenPages,
            placeholders,
            errors: failedErrors,
          });
        }

        return outcome.data.previewed({
          outputDir,
          pages: writtenPages,
          placeholders,
        });
      },
    });

    return {
      default: "text",
      modes: {
        text: previewDeck.text({
          description: "Human-readable text output",
          success: {
            previewed: ({ outputDir, pages, placeholders }) =>
              text.lines(
                toContractTextLines([
                  `Previewed ${pages.length} page(s) in ${outputDir}`,
                  ...pages.map(
                    (page) =>
                      `  Page ${page.page} (${page.id}): ${page.output}`,
                  ),
                  ...(placeholders.length > 0
                    ? [
                        "Placeholders:",
                        ...formatPreviewPlaceholderLines(placeholders),
                      ]
                    : []),
                ]),
              ),
          },
          failures: {
            invalidOutput: ({ output }) =>
              text.lines(
                toContractTextLines([
                  `Output directory cannot be a non-directory path: ${output}`,
                ]),
              ),
            pageOutOfRange: ({ requestedPages, totalSlides }) =>
              text.lines(
                toContractTextLines([
                  `Requested pages (${requestedPages.join(", ")}) exceed total slides (${totalSlides})`,
                ]),
              ),
            sourceLoadFailed: presentSourceLoadFailed,
            buildFailed: presentBuildFailed,
            previewFailed: ({ pages, placeholders, errors }) =>
              text.lines(
                toContractTextLines([
                  "Preview failed:",
                  ...errors.map(
                    (error) =>
                      `  Page ${error.page} ${error.path}${error.field ? ` (${error.field})` : ""} [${error.code}]: ${error.message}`,
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
                  ...(placeholders.length > 0
                    ? [
                        "Placeholders:",
                        ...formatPreviewPlaceholderLines(placeholders),
                      ]
                    : []),
                ]),
              ),
            writeFailed: ({ output, message }) =>
              text.lines(
                toContractTextLines([
                  `Failed to write preview image to ${output}: ${message}`,
                ]),
              ),
            cancelled: () => text.line("Deck preview cancelled"),
          },
        }),
        structured: previewDeck.structured({
          description: "Structured JSON output",
        }),
      },
    };
  },
});

const codegenCommand = define.command("codegen")({
  kind: "command",
  parent: "pptxdk",
  name: "codegen",
  description:
    "Generate rebuildable vanilla TypeScript or React TSX source from a PPTX file.",
  fields: {
    entry: {
      kind: "positional",
      description: "Input PPTX file path",
    },
    target: {
      kind: "valueOption",
      longOption: "--target",
      shortAlias: "-t",
      description: "Author source syntax target",
    },
    outDir: {
      kind: "valueOption",
      longOption: "--out-dir",
      description:
        "Directory for generated source and assets (defaults to <entry>-source in current directory)",
    },
    outZip: {
      kind: "valueOption",
      longOption: "--out-zip",
      description: "ZIP archive path for generated source and assets",
    },
  },
  usageConstraints: [{ kind: "count", fields: ["outDir", "outZip"], max: 1 }],
  input: codegenInputSchema,
  outputs: ({ atomic }) => {
    const codegenDeck = atomic({
      success: {
        kind: "data",
        variants: {
          generated: {
            description: "Generated rebuildable source and assets",
            exitCode: 0,
            schema: generatedSourceSchema,
          },
        },
      },
      failures: {
        invalidOutput: {
          description:
            "Output path is invalid, non-directory, or overwrites input PPTX",
          exitCode: 2,
          schema: invalidOutputSchema,
        },
        inputReadFailed: {
          description: "Failed to read input PPTX file",
          exitCode: 1,
          schema: inputReadFailedSchema,
        },
        importFailed: {
          description: "PPTX import failed with fatal issue",
          exitCode: 1,
          schema: importFailedSchema,
        },
        codegenFailed: {
          description: "Source generation failed for editable declarations",
          exitCode: 1,
          schema: codegenFailedSchema,
        },
        writeFailed: {
          description: "Failed to write generated source or ZIP archive",
          exitCode: 1,
          schema: writeFailedSchema,
        },
        cancelled: {
          description: "Source generation cancelled",
          exitCode: 130,
          schema: cancelledSchema,
        },
      },
      async handler({ input, dependencies, outcome }) {
        if (dependencies.signal?.aborted) {
          return outcome.failure.cancelled({});
        }

        const entry = resolve(dependencies.cwd, input.entry);
        const delivery = input.outZip === undefined ? "directory" : "zip";
        const output =
          input.outZip !== undefined
            ? resolve(dependencies.cwd, input.outZip)
            : resolve(
                dependencies.cwd,
                input.outDir ?? `${basename(entry, extname(entry))}-source`,
              );

        if (delivery === "zip") {
          if (await isSameExistingFile(entry, output)) {
            return outcome.failure.invalidOutput({ entry, output });
          }
        } else {
          const dirCheck = await validateOutputDir(entry, output);
          if (dirCheck.kind === "invalidOutput") {
            return outcome.failure.invalidOutput({ entry, output });
          }
          if (dirCheck.kind === "error") {
            return outcome.failure.writeFailed({
              output,
              message: errorMessage(dirCheck.error),
            });
          }
        }

        let imported: ImportResult;
        try {
          imported = await importPptxFile(entry, {
            ...(dependencies.signal !== undefined && {
              signal: dependencies.signal,
            }),
          });
        } catch (error) {
          if (isCancelledError(error, dependencies.signal)) {
            return outcome.failure.cancelled({});
          }
          if (error instanceof ImportError) {
            return outcome.failure.importFailed({
              entry,
              issue: error.issue,
              message: error.message,
            });
          }
          return outcome.failure.inputReadFailed({
            entry,
            message: errorMessage(error),
          });
        }

        if (dependencies.signal?.aborted) {
          return outcome.failure.cancelled({});
        }

        const importDiagnostics = [...imported.report.diagnostics];
        const runCodegen = dependencies.codegen ?? codegen;
        let generated: ReturnType<typeof codegen>;
        try {
          generated = runCodegen(imported.document, { target: input.target });
        } catch (error) {
          if (isCancelledError(error, dependencies.signal)) {
            return outcome.failure.cancelled({});
          }
          if (error instanceof CodegenError) {
            return outcome.failure.codegenFailed({
              entry,
              diagnostics: [...error.diagnostics],
              importDiagnostics,
            });
          }
          throw error;
        }

        if (dependencies.signal?.aborted) {
          return outcome.failure.cancelled({});
        }

        const normalizedFiles = normalizeCodegenFiles(generated.files);
        const filePaths = normalizedFiles.map((file) => file.path);

        if (delivery === "zip") {
          const zipEntries: Record<string, Uint8Array> = {};
          for (const file of normalizedFiles) {
            zipEntries[file.path] = file.bytes;
          }
          try {
            const zipBytes = zipSync(zipEntries, {
              level: 6,
              mtime: fixedZipMtime,
            });
            await writeArtifact(output, zipBytes);
          } catch (writeError) {
            if (isCancelledError(writeError, dependencies.signal)) {
              return outcome.failure.cancelled({});
            }
            return outcome.failure.writeFailed({
              output,
              message: errorMessage(writeError),
            });
          }
        } else {
          const plannedTargets: {
            readonly targetPath: string;
            readonly bytes: Uint8Array;
          }[] = [];
          for (const file of normalizedFiles) {
            const targetPath = resolve(output, ...file.path.split("/"));
            if (
              !isWithinDirectory(output, targetPath) ||
              (await isSameExistingFile(entry, targetPath))
            ) {
              return outcome.failure.invalidOutput({
                entry,
                output: targetPath,
              });
            }
            plannedTargets.push({ targetPath, bytes: file.bytes });
          }

          for (const item of plannedTargets) {
            if (dependencies.signal?.aborted) {
              return outcome.failure.cancelled({});
            }
            try {
              await writeArtifact(item.targetPath, item.bytes);
            } catch (writeError) {
              if (isCancelledError(writeError, dependencies.signal)) {
                return outcome.failure.cancelled({});
              }
              return outcome.failure.writeFailed({
                output: item.targetPath,
                message: errorMessage(writeError),
              });
            }
          }
        }

        return outcome.data.generated({
          target: generated.target,
          delivery,
          output,
          entry: generated.entry,
          files: filePaths,
          slides: imported.report.slideCount,
          diagnostics: importDiagnostics,
        });
      },
    });

    return {
      default: "text",
      modes: {
        text: codegenDeck.text({
          description: "Human-readable text output",
          success: {
            generated: ({
              target,
              delivery,
              output,
              entry,
              files,
              diagnostics,
            }) =>
              text.lines(
                toContractTextLines([
                  ...diagnostics.map(formatImportDiagnostic),
                  `Generated ${target} source (${delivery}): ${output}`,
                  `Entry: ${entry}`,
                  `Files: ${files.length}`,
                ]),
              ),
          },
          failures: {
            invalidOutput: ({ output }) =>
              text.lines(
                toContractTextLines([
                  `Output directory cannot be a non-directory path or overwrite input: ${output}`,
                ]),
              ),
            inputReadFailed: ({ entry, message }) =>
              text.lines(
                toContractTextLines([
                  `Failed to read PPTX input ${entry}: ${message}`,
                ]),
              ),
            importFailed: ({ entry, issue, message }) =>
              text.lines(
                toContractTextLines([
                  `Failed to import PPTX ${entry} [${issue.code}]: ${formatFatalImportIssue(issue)} (${message})`,
                ]),
              ),
            codegenFailed: ({ diagnostics, importDiagnostics }) =>
              text.lines(
                toContractTextLines([
                  ...importDiagnostics.map(formatImportDiagnostic),
                  "Source generation failed:",
                  ...diagnostics.map(
                    (diagnostic) =>
                      `  ${diagnostic.path}${diagnostic.field !== undefined ? ` (${diagnostic.field})` : ""} [${diagnostic.code}]: ${diagnostic.message}`,
                  ),
                ]),
              ),
            writeFailed: ({ output, message }) =>
              text.lines(
                toContractTextLines([
                  `Failed to write generated output to ${output}: ${message}`,
                ]),
              ),
            cancelled: () => text.line("Source generation cancelled"),
          },
        }),
        structured: codegenDeck.structured({
          description: "Structured JSON output",
        }),
      },
    };
  },
});

export const cliContract = define({
  root: "pptxdk",
  help: helpCapability({
    shortAlias: "-h",
    headings: {
      usage: "Usage",
      commands: "Commands",
      arguments: "Arguments",
      options: "Options",
      constraints: "Constraints",
      schema: "Schema",
    },
    wording: {
      commandPlaceholder: "command",
      outputModePlaceholder: "mode",
      choices: "{usage} (choices: {choices})",
      default: "{usage} (default: {value})",
      choicesAndDefault: "{usage} (choices: {choices}, default: {value})",
      list: "{left}, {right}",
      countAtMost: "At most {max} of {fields} may be specified",
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
      description: "Build, preview, and generate source from PowerPoint decks.",
    },
    ...buildCommand,
    ...previewCommand,
    ...codegenCommand,
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
