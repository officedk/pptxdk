import type { ReactNode } from "react";

import { BuildError } from "#/diagnostics.ts";
import {
  createCompletedDeclarationBuild,
  validateOptions,
} from "#/formal-build.ts";
import { expandHostTree } from "#/host/reconciler.ts";
import type {
  BuildOptions,
  BuildResult,
  CapabilityModule,
  Diagnostic,
} from "#/types.ts";

function diagnostic(path: string, code: string, message: string): Diagnostic {
  return { severity: "error", path, code, message };
}

/**
 * Creates the asynchronous React authoring entry.
 *
 * It runs author components to completion, waits for declared content, cleans up the React root,
 * and then hands the resulting declaration tree to the same formal build used by plain authoring
 * entries. Cancellation, unhandled errors, and cleanup failures keep the existing rejection
 * semantics.
 */
export function createBuilder(modules: readonly CapabilityModule[]) {
  const formalBuild = createCompletedDeclarationBuild(modules);

  return async function build(
    element: ReactNode,
    options: BuildOptions = {},
  ): Promise<BuildResult> {
    // Validate before React runs so invalid options never start author components.
    validateOptions(options);
    try {
      const declaration = await expandHostTree(element, options.signal);
      // Expansion locks its first terminal state before unmount cleanup runs. The completed
      // declaration build therefore does not re-check the signal, while preserving the original
      // options for capability parse and compile.
      return formalBuild(declaration, options);
    } catch (error) {
      if (options.signal?.aborted && error === options.signal.reason) {
        throw error;
      }
      if (error instanceof BuildError) throw error;
      throw new BuildError([
        diagnostic(
          "deck",
          "react-render",
          error instanceof Error ? error.message : String(error),
        ),
      ]);
    }
  };
}
