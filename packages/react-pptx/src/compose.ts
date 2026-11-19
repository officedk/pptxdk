import { expandDocument } from "@react-pptx/core";
import {
  createResolver,
  emit,
  validateOptions,
  type BuildOptions,
  type BuildResult,
  type CapabilityModule,
  type Resolution,
  type ResolveOptions,
} from "@react-pptx/core/build";
import { CommonNode } from "@react-pptx/core/document";
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
    if (input instanceof CommonNode) {
      // Resolve synchronously before yielding so later edits to the mutable document do not leak into this build.
      const resolution = resolve(input, options);
      return emit(
        resolution,
        options.signal === undefined ? {} : { signal: options.signal },
      );
    }
    // Validate options before React runs so invalid options never start author components.
    validateOptions(options);
    const document = await expandDocument(
      input,
      options.signal === undefined ? {} : { signal: options.signal },
    );
    const resolution = resolve(document, options);
    return emit(
      resolution,
      options.signal === undefined ? {} : { signal: options.signal },
    );
  }

  return { resolve, build };
}
