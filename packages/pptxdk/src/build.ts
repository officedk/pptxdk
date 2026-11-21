import {
  BuildError,
  CapabilityError,
  createResolver,
  formatBaseline,
  targets,
  type BuildOptions,
  type BuildResult,
  type CapabilityModule,
  type Resolution,
  type ResolveOptions,
} from "@pptxdk/core/build";
import type { CommonNode } from "@pptxdk/core/document";
import { emit, validateOptions } from "@pptxdk/emit";
import type { ReactNode } from "react";

import { emitBuildResult, resolveDocumentInput } from "./compose.ts";

/** Official capability module names accepted by `pptxdk/build`. */
export type OfficialCapabilityName =
  | "box"
  | "chart"
  | "group"
  | "image"
  | "motion"
  | "notes"
  | "shape"
  | "slide-background"
  | "table"
  | "text"
  | "transition";

/** Complete list of official capability names in canonical order. */
export const officialCapabilityNames = Object.freeze([
  "box",
  "chart",
  "group",
  "image",
  "motion",
  "notes",
  "shape",
  "slide-background",
  "table",
  "text",
  "transition",
] as const satisfies readonly OfficialCapabilityName[]);

const capabilityLoaders: Readonly<
  Record<OfficialCapabilityName, () => Promise<CapabilityModule>>
> = {
  box: async () => (await import("@pptxdk/capabilities/box")).boxCapability,
  chart: async () =>
    (await import("@pptxdk/capabilities/chart")).chartCapability,
  group: async () =>
    (await import("@pptxdk/capabilities/group")).groupCapability,
  image: async () =>
    (await import("@pptxdk/capabilities/image")).imageCapability,
  motion: async () =>
    (await import("@pptxdk/capabilities/motion")).motionCapability,
  notes: async () =>
    (await import("@pptxdk/capabilities/notes")).notesCapability,
  shape: async () =>
    (await import("@pptxdk/capabilities/shape")).shapeCapability,
  "slide-background": async () =>
    (await import("@pptxdk/capabilities/slide-background"))
      .slideBackgroundCapability,
  table: async () =>
    (await import("@pptxdk/capabilities/table")).tableCapability,
  text: async () => (await import("@pptxdk/capabilities/text")).textCapability,
  transition: async () =>
    (await import("@pptxdk/capabilities/transition")).transitionCapability,
};

export type TrimmedResolveOptions = ResolveOptions & {
  readonly capabilities: readonly OfficialCapabilityName[];
};

export type TrimmedBuildOptions = BuildOptions & {
  readonly capabilities: readonly OfficialCapabilityName[];
};

function assertOptionsObject(
  value: unknown,
  path: string,
): asserts value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new BuildError([
      {
        severity: "error",
        code: "invalid-options",
        path,
        message: "Build options must be an object",
      },
    ]);
  }
}

function validateCapabilityNames(
  names: unknown,
): readonly OfficialCapabilityName[] {
  if (!Array.isArray(names)) {
    throw new BuildError([
      {
        severity: "error",
        code: "invalid-capability-selection",
        path: "build.capabilities",
        message: "capabilities must be an array of official capability names",
      },
    ]);
  }
  const unique = new Set<OfficialCapabilityName>();
  for (let index = 0; index < names.length; index += 1) {
    const name: unknown = names[index];
    if (typeof name !== "string" || !Object.hasOwn(capabilityLoaders, name)) {
      throw new BuildError([
        {
          severity: "error",
          code: "unknown-capability",
          path: `build.capabilities[${index}]`,
          message: `Unknown official capability ${JSON.stringify(name)}`,
        },
      ]);
    }
    unique.add(name as OfficialCapabilityName);
  }
  return [...unique];
}

/**
 * Dynamically loads only the requested official capability modules in deduplicated first-seen order.
 */
export async function selectCapabilities(
  names: readonly OfficialCapabilityName[],
  options: { readonly signal?: AbortSignal } = {},
): Promise<readonly CapabilityModule[]> {
  const orderedNames = validateCapabilityNames(names);
  if (options.signal?.aborted) {
    throw options.signal.reason;
  }
  const modules = await Promise.all(
    orderedNames.map((name) => capabilityLoaders[name]()),
  );
  if (options.signal?.aborted) {
    throw options.signal.reason;
  }
  return Object.freeze(modules);
}

async function resolveSelected(
  input: ReactNode | CommonNode,
  capabilities: readonly OfficialCapabilityName[],
  signal?: AbortSignal,
): Promise<Resolution> {
  const signalOption = signal === undefined ? {} : { signal };
  const modules = await selectCapabilities(capabilities, signalOption);
  return resolveDocumentInput(createResolver(modules), input, signalOption);
}

/**
 * Loads only the selected official capabilities and resolves either a CommonNode document or a JSX
 * element into an immutable Resolution snapshot for `emit` and `paint`.
 */
export async function resolve(
  input: ReactNode | CommonNode,
  options: TrimmedResolveOptions,
): Promise<Resolution> {
  assertOptionsObject(options, "resolve");
  for (const key of Object.keys(options)) {
    if (key !== "capabilities" && key !== "signal") {
      throw new BuildError([
        {
          severity: "error",
          code: "unknown-parameter",
          path: "resolve",
          message: `Resolve options contain unknown parameter ${key}`,
        },
      ]);
    }
  }
  return resolveSelected(input, options.capabilities, options.signal);
}

/**
 * Loads only the selected official capabilities, resolves the document, and emits a PPTX archive.
 */
export async function build(
  input: ReactNode | CommonNode,
  options: TrimmedBuildOptions,
): Promise<BuildResult> {
  assertOptionsObject(options, "build");
  const { capabilities, ...buildOptions } = options;
  validateOptions(buildOptions);
  validateCapabilityNames(capabilities);
  const resolution = await resolveSelected(
    input,
    capabilities,
    buildOptions.signal,
  );
  return emitBuildResult(resolution, buildOptions);
}

export {
  BuildError,
  CapabilityError,
  createResolver,
  emit,
  formatBaseline,
  targets,
};
export type {
  BuildOptions,
  BuildReport,
  BuildResult,
  CapabilityModule,
  DeclarationChild,
  DeclarationContainer,
  DeclarationNode,
  DeclarationText,
  Diagnostic,
  EmitOptions,
  EmitResult,
  Extension,
  Resolution,
  ResolutionReport,
  ResolveOptions,
  Target,
} from "@pptxdk/core/build";
