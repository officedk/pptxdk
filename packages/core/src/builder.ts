import type { ReactNode } from "react";

import { assertAppendableAuthorIds } from "#/append.ts";
import { BuildError } from "#/diagnostics.ts";
import type { CommonContainer, CommonNode } from "#/document.ts";
import {
  createCompletedDeclarationBuild,
  validateOptions,
} from "#/formal-build.ts";
import { expandReactDocument, expandReactFragment } from "#/host/reconciler.ts";
import type {
  BuildOptions,
  BuildResult,
  CapabilityModule,
  Diagnostic,
} from "#/types.ts";

function diagnostic(path: string, code: string, message: string): Diagnostic {
  return { severity: "error", path, code, message };
}

function reactFailure(error: unknown, path = "deck"): BuildError {
  return new BuildError([
    diagnostic(
      path,
      "react-render",
      error instanceof Error ? error.message : String(error),
    ),
  ]);
}

export type ExpandDocumentOptions = {
  /**
   * Cooperative cancellation for one React expansion; no default timeout when unspecified.
   * Cancellation before or during expansion rejects with the original reason.
   */
  readonly signal?: AbortSignal;
};

/** Options for one JSX fragment append; cancellation covers expansion and cleanup only. */
export type AppendJSXOptions = ExpandDocumentOptions;

/** Completed editable document handed off by React expansion. */
export type ExpandedDocument = CommonContainer<"deck">;

/**
 * Public React handoff entry: runs author components to completion, waits for declared content,
 * cleans up the React root, and returns the completed editable document.
 *
 * The returned nodes no longer synchronize with the finished root. Repeat expansion creates a new
 * document instance and does not merge edits made to a previous result. Unhandled render errors,
 * cancellation, and cleanup failures reject; no partial document is returned.
 */
export async function expandDocument(
  element: ReactNode,
  options: ExpandDocumentOptions = {},
): Promise<ExpandedDocument> {
  const signal = options.signal;
  try {
    return await expandReactDocument(element, signal);
  } catch (error) {
    if (signal?.aborted && error === signal.reason) throw error;
    if (error instanceof BuildError) throw error;
    throw reactFailure(error);
  }
}

function appendFailure(error: unknown): BuildError {
  return new BuildError([
    diagnostic(
      "fragment",
      "append",
      error instanceof Error ? error.message : String(error),
    ),
  ]);
}

/**
 * Public React append entry: expands one standalone JSX fragment, waits for final content and
 * cleanup, then appends its top-level editable nodes to `target` as one batch.
 *
 * The fragment runs in its own root, so React context comes from the fragment or an explicit
 * Provider inside it; the target's finished root is not revived. Expansion, unhandled errors,
 * cancellation, cleanup failures, and id conflicts leave the target unchanged. Id conflicts are
 * checked against the target's current structure at commit time, so edits made while the fragment
 * expands are respected. The returned nodes are the exact instances now owned by the target.
 */
export async function appendJSX(
  target: CommonContainer,
  element: ReactNode,
  options: AppendJSXOptions = {},
): Promise<readonly CommonNode[]> {
  const signal = options.signal;
  let nodes: readonly CommonNode[];
  try {
    nodes = await expandReactFragment(element, signal);
  } catch (error) {
    if (signal?.aborted && error === signal.reason) throw error;
    if (error instanceof BuildError) throw error;
    throw reactFailure(error, "fragment");
  }
  try {
    assertAppendableAuthorIds(target, nodes);
    target.append(...nodes);
    return nodes;
  } catch (error) {
    if (error instanceof BuildError) throw error;
    throw appendFailure(error);
  }
}

/**
 * Creates the asynchronous one-call JSX authoring entry.
 *
 * It composes the public React handoff with the same formal build used by plain authoring entries.
 * Cancellation, unhandled errors, and cleanup failures keep the existing rejection semantics, and
 * the original build options continue to reach capability parse and compile contexts.
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
      const document = await expandDocument(
        element,
        options.signal === undefined ? {} : { signal: options.signal },
      );
      // Expansion locks its first terminal state before unmount cleanup runs. The completed
      // declaration build therefore does not re-check the signal, while preserving the original
      // options for capability parse and compile.
      return formalBuild({ children: [document] }, options);
    } catch (error) {
      if (options.signal?.aborted && error === options.signal.reason) {
        throw error;
      }
      if (error instanceof BuildError) throw error;
      throw reactFailure(error);
    }
  };
}
