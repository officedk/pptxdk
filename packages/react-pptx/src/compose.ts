import {
  createResolver,
  type BuildOptions,
  type BuildResult,
  type CapabilityModule,
  type Resolution,
  type ResolveOptions,
} from "@react-pptx/core/build";
import { CommonNode } from "@react-pptx/core/document";
import { emit, emittedBuildQuery, validateOptions } from "@react-pptx/emit";
import { expandDocument } from "@react-pptx/react";
import type { ReactNode } from "react";

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
    if (input instanceof CommonNode) {
      // Resolve synchronously before yielding so later edits to the mutable document do not leak into this build.
      const resolution = resolve(input, resolveOptions);
      const emitted = emit(resolution, options);
      return Object.freeze({
        pptx: emitted.pptx,
        report: emitted.report,
        query: emittedBuildQuery(emitted.report, resolution.query),
      });
    }
    const document = await expandDocument(
      input,
      options.signal === undefined ? {} : { signal: options.signal },
    );
    const resolution = resolve(document, resolveOptions);
    const emitted = emit(resolution, options);
    return Object.freeze({
      pptx: emitted.pptx,
      report: emitted.report,
      query: emittedBuildQuery(emitted.report, resolution.query),
    });
  }

  return { resolve, build };
}
