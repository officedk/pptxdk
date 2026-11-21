import { readFileSync } from "node:fs";
import { registerHooks } from "node:module";
import { extname, sep } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { transformSync, type Loader } from "esbuild";
import {
  createElement,
  isValidElement,
  type ComponentType,
  type ReactNode,
} from "react";

import { build } from "./index.ts";
import type { BuildOptions, BuildResult } from "./index.ts";

const sourceLoaders: Readonly<Record<string, Loader>> = {
  ".ts": "ts",
  ".tsx": "tsx",
  ".mts": "ts",
  ".jsx": "jsx",
};

export type BuildFromSourceOptions = BuildOptions & {
  /** Use the named export of the entry module as the deck; defaults to "default". */
  readonly exportName?: string;
};

/** Error thrown when loading author deck source fails. */
export class SourceLoadError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "SourceLoadError";
  }
}

function registerSourceLoader() {
  // Deferred imports also require source transformation, so hold the loader hook throughout the source build.
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

type SourceModuleLoad =
  | { readonly kind: "loaded"; readonly module: Record<string, unknown> }
  | { readonly kind: "cancelled" }
  | { readonly kind: "failed"; readonly cause: unknown };

function importSourceModule(
  url: URL,
  signal?: AbortSignal,
): Promise<SourceModuleLoad> {
  return new Promise((resolve) => {
    let settled = false;
    const settle = (result: SourceModuleLoad) => {
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

    // Imports cannot be aborted once initiated; intercept late settlement to prevent unhandled rejection.
    void import(url.href).then(
      (module: Record<string, unknown>) => settle({ kind: "loaded", module }),
      (cause: unknown) => settle({ kind: "failed", cause }),
    );
  });
}

async function loadDeck(
  url: URL,
  exportName: string,
  signal?: AbortSignal,
): Promise<ReactNode> {
  const loaded = await importSourceModule(url, signal);
  if (loaded.kind === "cancelled") throw signal?.reason;

  if (loaded.kind === "failed") {
    const { cause } = loaded;
    const details =
      cause instanceof Error ? (cause.stack ?? cause.message) : String(cause);
    throw new SourceLoadError(
      `Failed to load deck source at ${fileURLToPath(url)}:\n${details}`,
      {
        cause,
      },
    );
  }

  const deckModule = loaded.module;
  const exported = deckModule[exportName];
  if (isValidElement(exported)) return exported;
  if (typeof exported === "function") {
    return createElement(exported as ComponentType);
  }

  throw new SourceLoadError(
    `Deck source at ${fileURLToPath(url)} must export a React element or component named ${JSON.stringify(exportName)}`,
  );
}

/**
 * Run a formal build from a Node.js local ESM source entry. TS, TSX, MTS, and JSX are loaded by this entry;
 * the entry and its source dependencies are resolved according to Node file locations. Cancellation signals
 * also cover source import and preserve the original cancellation reason.
 * The return value is identical to `build`.
 */
export async function buildFromSource(
  entry: string | URL,
  options: BuildFromSourceOptions = {},
): Promise<BuildResult> {
  const { exportName = "default", ...buildOptions } = options;
  const url = sourceUrl(entry);
  if (buildOptions.signal?.aborted) return build(null, buildOptions);

  const hooks = registerSourceLoader();
  try {
    const deck = await loadDeck(url, exportName, buildOptions.signal);
    return await build(deck, buildOptions);
  } finally {
    hooks.deregister();
  }
}
