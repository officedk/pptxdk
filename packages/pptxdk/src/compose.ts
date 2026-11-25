import {
  createResolver,
  type BuildOptions,
  type BuildResult,
  type CapabilityModule,
  type Resolution,
  type ResolveOptions,
} from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";
import { emit, emittedBuildQuery, validateOptions } from "@pptxdk/emit";
import type { ReactNode } from "react";

/**
 * Resolves either an editable CommonNode synchronously or a React JSX element via on-demand
 * expansion using the supplied resolver.
 */
export async function resolveDocumentInput(
  resolveDeclaration: (
    document: CommonNode,
    options?: ResolveOptions,
  ) => Resolution,
  input: ReactNode | CommonNode,
  options: ResolveOptions = {},
): Promise<Resolution> {
  if (input instanceof CommonNode) {
    // Resolve synchronously before yielding so later edits to the mutable document do not leak into this build.
    return resolveDeclaration(input, options);
  }
  // Load the React reconciler adapter only when JSX input requires expansion.
  const { expandDocument } = await import("@pptxdk/react");
  const document = await expandDocument(input, options);
  return resolveDeclaration(document, options);
}

/** Emits a PPTX archive and BuildResult from an immutable Resolution snapshot. */
export function emitBuildResult(
  resolution: Resolution,
  options: BuildOptions = {},
): BuildResult {
  const emitted = emit(resolution, options);
  return Object.freeze({
    pptx: emitted.pptx,
    report: emitted.report,
    query: emittedBuildQuery(emitted.report, resolution.query),
  });
}

/**
 * Composes synchronous `resolve` and async-compatible `build` (supporting both JSX and CommonNode)
 * over one capability module set for entrypoints that assemble capabilities directly.
 */
export function createStagePipeline(modules: readonly CapabilityModule[]) {
  const resolveDeclaration = createResolver(modules);

  function resolve(
    document: CommonNode,
    options: ResolveOptions = {},
  ): Resolution {
    return resolveDeclaration(document, options);
  }

  async function build(
    input: ReactNode | CommonNode,
    options: BuildOptions = {},
  ): Promise<BuildResult> {
    // Validate build/target options upfront so invalid options fail before React expansion or resolve runs.
    validateOptions(options);
    const resolveOptions: ResolveOptions =
      options.signal === undefined ? {} : { signal: options.signal };
    const resolution = await resolveDocumentInput(
      resolveDeclaration,
      input,
      resolveOptions,
    );
    return emitBuildResult(resolution, options);
  }

  return { resolve, build };
}
