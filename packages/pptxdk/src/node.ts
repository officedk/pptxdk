import { readFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { registerHooks } from "node:module";
import { extname, sep } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { toBuildDiagnostic } from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";
import { validateOptions } from "@pptxdk/emit";
import { transformSync, type Loader } from "esbuild";
import type { ComponentType, ReactNode } from "react";
import { z } from "zod";

import { resolveDocumentInput } from "./compose.ts";
import {
  BuildError,
  ImportError,
  build,
  buildDiagnosticSchema,
  buildOptionsSchema,
  importPptx,
  previewErrorSchema,
  previewPlaceholderSchema,
  resolve,
  resolveOptionsSchema,
  type BuildDiagnostic,
  type BuildOptions,
  type BuildResult,
  type FatalImportIssue,
  type ImportDiagnostic,
  type ImportOptions,
  type ImportReport,
  type ImportResult,
  type PreviewError,
  type PreviewPlaceholder,
  type Resolution,
  type ResolveOptions,
} from "./index.ts";
import type { PreviewOptions, PreviewResult } from "./preview.ts";

const sourceLoaders: Readonly<Record<string, Loader>> = {
  ".ts": "ts",
  ".tsx": "tsx",
  ".mts": "ts",
  ".jsx": "jsx",
};

export const exportNameSchema = z.string().min(1);

export const buildFromSourceOptionsSchema = buildOptionsSchema
  .unwrap()
  .extend({
    /** Use the named export of the entry module as the deck; defaults to "default". */
    exportName: exportNameSchema.optional(),
  })
  .readonly();
export type BuildFromSourceOptions = z.infer<
  typeof buildFromSourceOptionsSchema
>;

export const previewFromSourceOptionsSchema = z
  .strictObject({
    pages: z.array(z.int().positive()).readonly().optional(),
    signal: z.instanceof(AbortSignal).optional(),
    /** Use the named export of the entry module as the deck; defaults to "default". */
    exportName: exportNameSchema.optional(),
  })
  .readonly();
export type PreviewFromSourceOptions = z.infer<
  typeof previewFromSourceOptionsSchema
>;

export const resolveFromSourceOptionsSchema = resolveOptionsSchema
  .unwrap()
  .extend({
    /** Use the named export of the entry module as the deck; defaults to "default". */
    exportName: exportNameSchema.optional(),
  })
  .readonly();
export type ResolveFromSourceOptions = z.infer<
  typeof resolveFromSourceOptionsSchema
>;

export type {
  BuildDiagnostic,
  BuildOptions,
  BuildResult,
  FatalImportIssue,
  ImportDiagnostic,
  ImportOptions,
  ImportReport,
  ImportResult,
  PreviewError,
  PreviewOptions,
  PreviewPlaceholder,
  PreviewResult,
  Resolution,
  ResolveOptions,
};
export {
  ImportError,
  buildDiagnosticSchema,
  previewErrorSchema,
  previewPlaceholderSchema,
  resolveOptionsSchema,
};

/**
 * Reads a Transitional PPTX file from disk and imports it into an editable CommonNode document
 * and structured `ImportReport`, detaching completely from the file buffer after import.
 */
export async function importPptxFile(
  entry: string | URL,
  options: ImportOptions = {},
): Promise<ImportResult> {
  if (options.signal?.aborted) {
    throw options.signal.reason;
  }
  const bytes = await readFile(
    entry,
    options.signal === undefined ? undefined : { signal: options.signal },
  );
  return importPptx(new Uint8Array(bytes), options);
}

/** Error thrown when loading author deck source fails. */
export class SourceLoadError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "SourceLoadError";
  }
}

function registerSourceLoader() {
  // Deferred imports also require source transformation, so hold the loader hook throughout the source operation.
  return registerHooks({
    load(url, context, nextLoad) {
      if (!url.startsWith("file:")) return nextLoad(url, context);

      const file = fileURLToPath(url);
      if (file.split(sep).includes("node_modules")) {
        return nextLoad(url, context);
      }

      const loader = sourceLoaders[extname(file)];
      if (loader === undefined) return nextLoad(url, context);

      const result = transformSync(readFileSync(file, "utf8"), {
        loader,
        format: "esm",
        jsx: "automatic",
        jsxImportSource: "react",
        sourcemap: "inline",
        sourcefile: file,
        target: "node24",
      });
      return { format: "module", source: result.code, shortCircuit: true };
    },
  });
}

function sourceUrl(entry: string | URL): URL {
  if (entry instanceof URL) {
    if (entry.protocol !== "file:") {
      throw new SourceLoadError(
        "Deck entry must be a local file path or file URL",
      );
    }
    return entry;
  }
  return pathToFileURL(entry);
}

type SettledPromise<T> =
  | { readonly kind: "resolved"; readonly value: T }
  | { readonly kind: "cancelled" }
  | { readonly kind: "rejected"; readonly cause: unknown };

function settleWithSignal<T>(
  promise: PromiseLike<T>,
  signal?: AbortSignal,
): Promise<SettledPromise<T>> {
  return new Promise((resolve) => {
    let settled = false;
    const settle = (result: SettledPromise<T>) => {
      if (settled) return;
      settled = true;
      signal?.removeEventListener("abort", onAbort);
      resolve(result);
    };
    const onAbort = () => settle({ kind: "cancelled" });

    if (signal?.aborted) {
      onAbort();
      return;
    }
    signal?.addEventListener("abort", onAbort, { once: true });
    if (signal?.aborted) {
      onAbort();
      return;
    }

    // Intercept late settlement after abort to prevent unhandled rejection in the host process.
    void Promise.resolve(promise).then(
      (value) => settle({ kind: "resolved", value }),
      (cause: unknown) => settle({ kind: "rejected", cause }),
    );
  });
}

function isPromiseLike(value: unknown): value is PromiseLike<unknown> {
  return (
    typeof value === "object" &&
    value !== null &&
    "then" in value &&
    typeof (value as { readonly then?: unknown }).then === "function"
  );
}

async function awaitSourcePromise<T>(
  promise: PromiseLike<T>,
  filePath: string,
  signal?: AbortSignal,
): Promise<T> {
  const settled = await settleWithSignal(promise, signal);
  if (settled.kind === "cancelled") throw signal?.reason;
  if (settled.kind === "rejected") {
    if (signal?.aborted && settled.cause === signal.reason) {
      throw settled.cause;
    }
    const details =
      settled.cause instanceof Error
        ? (settled.cause.stack ?? settled.cause.message)
        : String(settled.cause);
    throw new SourceLoadError(
      `Failed to load deck source at ${filePath}:\n${details}`,
      { cause: settled.cause },
    );
  }
  return settled.value;
}

async function loadDeck(
  url: URL,
  exportName: string,
  signal?: AbortSignal,
): Promise<ReactNode | CommonNode> {
  const filePath = fileURLToPath(url);
  const deckModule = await awaitSourcePromise<Record<string, unknown>>(
    import(url.href),
    filePath,
    signal,
  );

  if (
    !Object.hasOwn(deckModule, exportName) ||
    deckModule[exportName] === undefined
  ) {
    throw new SourceLoadError(
      `Deck source at ${filePath} does not export ${JSON.stringify(exportName)}`,
    );
  }

  const exported = deckModule[exportName];
  if (exported instanceof CommonNode) {
    return exported;
  }

  if (isPromiseLike(exported)) {
    const resolved = await awaitSourcePromise(exported, filePath, signal);
    if (resolved instanceof CommonNode) {
      return resolved;
    }
    throw new SourceLoadError(
      `Deck source at ${filePath} export ${JSON.stringify(exportName)} resolved to an unsupported value; expected an editable document node`,
    );
  }

  if (typeof exported === "function") {
    const { createElement } = await import("react");
    return createElement(exported as ComponentType);
  }
  if (
    typeof exported === "object" &&
    exported !== null &&
    !Array.isArray(exported)
  ) {
    const { isValidElement } = await import("react");
    if (isValidElement(exported)) return exported;
  }

  throw new SourceLoadError(
    `Deck source at ${filePath} must export an editable document node, Promise of document node, React element, or zero-arg React component named ${JSON.stringify(exportName)}`,
  );
}

/**
 * Run a formal build from a Node.js local ESM source entry. TS, TSX, MTS, and JSX are loaded by this entry;
 * the entry and its source dependencies are resolved according to Node file locations. Cancellation signals
 * also cover source import, async deck promises, and expansion, preserving the original cancellation reason.
 * The return value is identical to `build`.
 */
export async function buildFromSource(
  entry: string | URL,
  options: BuildFromSourceOptions = {},
): Promise<BuildResult> {
  const url = sourceUrl(entry);
  const parsed = buildFromSourceOptionsSchema.safeParse(options);
  if (!parsed.success) {
    if (
      typeof options !== "object" ||
      options === null ||
      Array.isArray(options)
    ) {
      validateOptions(options as BuildOptions);
    }
    const { exportName: _exportName, ...buildOptions } = options as Record<
      string,
      unknown
    >;
    validateOptions(buildOptions);
    throw new BuildError([
      toBuildDiagnostic(
        "build.exportName",
        "invalid-options",
        "exportName must be a non-empty string",
        { field: "exportName" },
      ),
    ]);
  }
  const { exportName = "default", ...buildOptions } = parsed.data;
  if (buildOptions.signal?.aborted) throw buildOptions.signal.reason;

  const hooks = registerSourceLoader();
  try {
    const deck = await loadDeck(url, exportName, buildOptions.signal);
    return await build(deck, buildOptions);
  } finally {
    hooks.deregister();
  }
}

/**
 * Run a formal static preview from a Node.js local ESM source entry. Shares source loading, export selection,
 * and cancellation semantics with `buildFromSource`.
 * The return value is identical to `preview`.
 */
export async function previewFromSource(
  entry: string | URL,
  options: PreviewFromSourceOptions = {},
): Promise<PreviewResult> {
  const url = sourceUrl(entry);
  if (
    typeof options === "object" &&
    options !== null &&
    "signal" in options &&
    options.signal instanceof AbortSignal &&
    options.signal.aborted
  ) {
    throw options.signal.reason;
  }
  const parsed = previewFromSourceOptionsSchema.safeParse(options);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    if (issue?.path[0] === "pages") {
      throw new RangeError("Preview page numbers must be positive integers");
    }
    if (issue?.path[0] === "exportName") {
      throw new TypeError(
        "Preview options exportName must be a non-empty string",
      );
    }
    if (issue?.path[0] === "signal") {
      throw new TypeError("Preview options signal must be an AbortSignal");
    }
    if (issue?.path.length === 0 && issue.code === "unrecognized_keys") {
      throw new TypeError(
        `Preview options contain unknown parameter ${issue.keys[0]}`,
      );
    }
    throw new TypeError("Preview options must be an object");
  }
  const { exportName = "default", ...previewOptions } = parsed.data;

  const { preview } = await import("./preview.ts");
  const hooks = registerSourceLoader();
  try {
    const deck = await loadDeck(url, exportName, previewOptions.signal);
    return await preview(deck, previewOptions);
  } finally {
    hooks.deregister();
  }
}

/**
 * Run a read-only resolve from a Node.js local ESM source entry. Shares source loading, export selection,
 * and cancellation semantics with `buildFromSource` and `previewFromSource`, without emitting PPTX or PNGs.
 * The return value is identical to `resolve`.
 */
export async function resolveFromSource(
  entry: string | URL,
  options: ResolveFromSourceOptions = {},
): Promise<Resolution> {
  const url = sourceUrl(entry);
  const parsed = resolveFromSourceOptionsSchema.safeParse(options);
  if (!parsed.success) {
    if (
      typeof options !== "object" ||
      options === null ||
      Array.isArray(options)
    ) {
      throw new BuildError([
        toBuildDiagnostic(
          "resolve",
          "invalid-options",
          "Resolve options must be an object",
        ),
      ]);
    }
    const issue = parsed.error.issues[0];
    if (issue?.path.length === 0 && issue.code === "unrecognized_keys") {
      throw new BuildError([
        toBuildDiagnostic(
          "resolve",
          "unknown-parameter",
          `Resolve options contain unknown parameter ${issue.keys[0]}`,
          { prop: String(issue.keys[0]), scope: "options" },
        ),
      ]);
    }
    if (issue?.path[0] === "signal") {
      throw new BuildError([
        toBuildDiagnostic(
          "resolve.signal",
          "invalid-options",
          "Resolve options signal must be an AbortSignal",
          { field: "signal" },
        ),
      ]);
    }
    if (issue?.path[0] === "exportName") {
      throw new BuildError([
        toBuildDiagnostic(
          "resolve.exportName",
          "invalid-options",
          "exportName must be a non-empty string",
          { field: "exportName" },
        ),
      ]);
    }
    throw new BuildError([
      toBuildDiagnostic(
        "resolve",
        "invalid-options",
        "Resolve options must be an object",
      ),
    ]);
  }
  const { exportName = "default", ...resolveOptions } = parsed.data;
  if (resolveOptions.signal?.aborted) throw resolveOptions.signal.reason;

  const hooks = registerSourceLoader();
  try {
    const deck = await loadDeck(url, exportName, resolveOptions.signal);
    return await resolveDocumentInput(resolve, deck, resolveOptions);
  } finally {
    hooks.deregister();
  }
}
