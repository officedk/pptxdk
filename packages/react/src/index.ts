import {
  BuildError,
  toBuildDiagnostic,
  type BuildDiagnostic,
} from "@pptxdk/core";
import type { CommonContainer, CommonNode } from "@pptxdk/core/document";
import type { ReactNode } from "react";

import { assertAppendableAuthorIds } from "#/append.ts";
import { expandReactDocument, expandReactFragment } from "#/reconciler.ts";

function diagnostic(
  path: string,
  code: string,
  message: string,
): BuildDiagnostic {
  return toBuildDiagnostic(path, code, message);
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
  readonly signal?: AbortSignal | undefined;
};

/** Options for one standalone JSX fragment expansion; cancellation covers expansion and cleanup only. */
export type ExpandFragmentOptions = ExpandDocumentOptions;

/** Options for one JSX fragment append; cancellation covers expansion and cleanup only. */
export type AppendJSXOptions = ExpandFragmentOptions;

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
 * Public React fragment handoff entry: runs a standalone JSX fragment to completion, waits for
 * declared async content and root cleanup, and returns frozen ownerless top-level editable nodes.
 *
 * The returned nodes no longer synchronize with the finished root and have `parent === undefined`.
 * Top-level non-whitespace text, unhandled render errors, cleanup failures, and cancellation reject
 * without returning partial results.
 */
export async function expandFragment(
  element: ReactNode,
  options: ExpandFragmentOptions = {},
): Promise<readonly CommonNode[]> {
  const signal = options.signal;
  try {
    return await expandReactFragment(element, signal);
  } catch (error) {
    if (signal?.aborted && error === signal.reason) throw error;
    if (error instanceof BuildError) throw error;
    throw reactFailure(error, "fragment");
  }
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
    nodes = await expandFragment(element, options);
  } catch (error) {
    if (signal?.aborted && error === signal.reason) throw error;
    if (error instanceof BuildError) {
      const hasExpandFailure = error.diagnostics.some(
        (item) => item.code === "render-failed" && item.phase === "expand",
      );
      if (!hasExpandFailure) throw error;
      throw new BuildError(
        error.diagnostics.map((item) =>
          item.code === "render-failed" && item.phase === "expand"
            ? { ...item, phase: "append" }
            : item,
        ),
      );
    }
    throw appendFailure(error);
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

export { Deck, Section, Slide, SlideLayout, SlideMaster } from "#/elements.ts";
export type {
  DeckProps,
  ElementProps,
  SectionProps,
  SlideLayoutProps,
  SlideMasterProps,
  SlideProps,
} from "#/elements.ts";
export { Mount } from "#/mount.ts";
export type { MountProps } from "#/mount.ts";
export { fromVanilla } from "#/from-vanilla.ts";
export type { VanillaSlotResult } from "#/from-vanilla.ts";
