import {
  BuildError,
  createResolver,
  toBuildDiagnostic,
  type BuildOptions,
  type BuildResult,
  type CapabilityModule,
  type Resolution,
  type ResolveOptions,
} from "@pptxdk/core/build";
import { CommonNode } from "@pptxdk/core/document";
import { emit, emittedBuildQuery, validateOptions } from "@pptxdk/emit";
import type { ReactNode } from "react";

function isJsxLikeInput(value: unknown): boolean {
  if (
    typeof value !== "object" ||
    value === null ||
    Array.isArray(value) ||
    value instanceof CommonNode
  ) {
    return false;
  }
  if (
    "$$typeof" in value &&
    typeof (value as { readonly $$typeof?: unknown }).$$typeof === "symbol"
  ) {
    return true;
  }
  // Plain JSX-like object `{ type, props }` without a DeclarationContainer `children` array.
  return (
    "type" in value &&
    "props" in value &&
    !(
      "children" in value &&
      Array.isArray((value as { readonly children?: unknown }).children)
    )
  );
}

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
  const query = emittedBuildQuery(emitted.report, resolution);
  return Object.freeze({
    pptx: emitted.pptx,
    report: emitted.report,
    slides: query.slides,
    slide: (id: string) => query.slide(id),
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
    if (isJsxLikeInput(document)) {
      throw new BuildError([
        toBuildDiagnostic(
          "deck",
          "invalid-structure",
          "Root resolve() synchronously accepts only an expanded CommonNode document, not a JSX/React element; await expandDocument(element) first, call resolveFromSource() from @officedk/pptxdk/node, or use async resolve() from @officedk/pptxdk/build",
        ),
      ]);
    }
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
