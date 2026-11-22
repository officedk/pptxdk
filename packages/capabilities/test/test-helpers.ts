import { documentCapabilities } from "@pptxdk/capabilities/all";
import type {
  BuildOptions,
  BuildResult,
  Resolution,
  ResolveOptions,
} from "@pptxdk/core/build";
import { createResolver } from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";
import { emit, emittedBuildQuery, validateOptions } from "@pptxdk/emit";
import {
  paint,
  validatePaintOptions,
  type PreviewOptions,
  type PreviewResult,
} from "@pptxdk/preview";
import { expandDocument } from "@pptxdk/react";
import type { ReactNode } from "react";

const resolveDefault = createResolver(documentCapabilities);

/**
 * Synchronously resolves an editable CommonNode document with the default capability set.
 */
export function resolve(
  document: CommonNode,
  options: ResolveOptions = {},
): Resolution {
  return resolveDefault(document, options);
}

/**
 * Stage-composed build helper for capability tests across both JSX and CommonNode entries.
 */
export async function build(
  input: ReactNode | CommonNode,
  options: BuildOptions = {},
): Promise<BuildResult> {
  validateOptions(options);
  const resolveOptions: ResolveOptions =
    options.signal === undefined ? {} : { signal: options.signal };
  const document =
    input instanceof CommonNode
      ? input
      : await expandDocument(
          input,
          options.signal === undefined ? {} : { signal: options.signal },
        );
  const resolution = resolveDefault(document, resolveOptions);
  const emitted = emit(resolution, options);
  return Object.freeze({
    pptx: emitted.pptx,
    report: emitted.report,
    query: emittedBuildQuery(emitted.report, resolution.query),
  });
}

export async function buildDocument(
  document: CommonNode,
  options: BuildOptions = {},
): Promise<BuildResult> {
  return build(document, options);
}

/**
 * Stage-composed static preview helper for capability tests across both JSX and CommonNode entries.
 */
export async function preview(
  input: ReactNode | CommonNode,
  options: PreviewOptions = {},
): Promise<PreviewResult> {
  const validatedOptions = validatePaintOptions(options);
  if (validatedOptions.signal?.aborted) throw validatedOptions.signal.reason;
  const signalOption =
    validatedOptions.signal === undefined
      ? {}
      : { signal: validatedOptions.signal };
  const document =
    input instanceof CommonNode
      ? input
      : await expandDocument(input, signalOption);
  const resolution = resolveDefault(document, signalOption);
  const { pages } = paint(resolution, validatedOptions);
  return Object.freeze({
    report: resolution.report,
    query: resolution.query,
    pages,
  });
}
