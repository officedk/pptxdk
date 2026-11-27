import {
  CodegenError,
  codegenDiagnosticCodeSchema,
  codegenDiagnosticSchema,
  codegenOptionsSchema,
  codegenTargetSchema,
  createCodegen,
  type CodegenDiagnostic,
  type CodegenDiagnosticCode,
  type CodegenEntryName,
  type CodegenFileContent,
  type CodegenOptions,
  type CodegenResult,
  type CodegenTarget,
} from "@pptxdk/codegen";
import {
  BuildError,
  CapabilityError,
  buildDiagnosticFontUnavailableReasonSchema,
  buildDiagnosticSchema,
  buildOptionsSchema,
  buildOptionsShape,
  createResolver,
  emitOptionsSchema,
  extensionSchema,
  fontEmbeddingModeSchema,
  formatBaseline,
  formatBaselineSchema,
  previewErrorSchema,
  previewPlaceholderSchema,
  resolveOptionsSchema,
  resolveOptionsShape,
  targetNameSchema,
  targetSchema,
  targets,
  toBuildDiagnostic,
  type BuildOptions,
  type BuildResult,
  type CapabilityModule,
  type Resolution,
  type ResolveOptions,
} from "@pptxdk/core/build";
import type { CommonNode } from "@pptxdk/core/document";
import { emit, validateOptions } from "@pptxdk/emit";
import {
  ImportError,
  cascadedOmissionSchema,
  createImporter,
  editabilityLossSchema,
  fatalImportIssueSchema,
  importDiagnosticSchema,
  importLocationSchema,
  importOptionsSchema,
  importReportSchema,
} from "@pptxdk/import";
import type { ReactNode } from "react";
import { z } from "zod";

import { emitBuildResult, resolveDocumentInput } from "./compose.ts";

type AssertExact<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : never;

/** Complete list of official capability names in canonical order. */
export const officialCapabilityNames = Object.freeze([
  "box",
  "chart",
  "footer-placeholder",
  "group",
  "image",
  "masters",
  "motion",
  "notes",
  "placeholder",
  "shape",
  "slide-background",
  "table",
  "text",
  "transition",
] as const);

/** Official capability module names accepted by `@officedk/pptxdk/build`. */
export type OfficialCapabilityName =
  | "box"
  | "chart"
  | "footer-placeholder"
  | "group"
  | "image"
  | "masters"
  | "motion"
  | "notes"
  | "placeholder"
  | "shape"
  | "slide-background"
  | "table"
  | "text"
  | "transition";

const rawOfficialCapabilityNameSchema = z.enum(officialCapabilityNames);
export const officialCapabilityNameSchema: z.ZodType<OfficialCapabilityName> =
  rawOfficialCapabilityNameSchema satisfies z.ZodType<
    AssertExact<
      OfficialCapabilityName,
      z.infer<typeof rawOfficialCapabilityNameSchema>
    > extends true
      ? OfficialCapabilityName
      : never
  >;

const capabilitiesSelectionSchema = z
  .array(rawOfficialCapabilityNameSchema)
  .readonly();

const capabilityLoaders: Readonly<
  Record<OfficialCapabilityName, () => Promise<CapabilityModule>>
> = {
  box: async () => (await import("@pptxdk/capabilities/box")).boxCapability,
  chart: async () =>
    (await import("@pptxdk/capabilities/chart")).chartCapability,
  "footer-placeholder": async () =>
    (await import("@pptxdk/capabilities/footer-placeholder"))
      .footerPlaceholderCapability,
  group: async () =>
    (await import("@pptxdk/capabilities/group")).groupCapability,
  image: async () =>
    (await import("@pptxdk/capabilities/image")).imageCapability,
  masters: async () =>
    (await import("@pptxdk/capabilities/masters")).mastersCapability,
  motion: async () =>
    (await import("@pptxdk/capabilities/motion")).motionCapability,
  notes: async () =>
    (await import("@pptxdk/capabilities/notes")).notesCapability,
  placeholder: async () =>
    (await import("@pptxdk/capabilities/placeholder")).placeholderCapability,
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

const rawTrimmedResolveOptionsSchema = z
  .strictObject({
    ...resolveOptionsShape,
    capabilities: capabilitiesSelectionSchema,
  })
  .readonly();
export const trimmedResolveOptionsSchema: z.ZodType<TrimmedResolveOptions> =
  rawTrimmedResolveOptionsSchema satisfies z.ZodType<
    AssertExact<
      TrimmedResolveOptions,
      z.infer<typeof rawTrimmedResolveOptionsSchema>
    > extends true
      ? TrimmedResolveOptions
      : never
  >;

export type TrimmedBuildOptions = BuildOptions & {
  readonly capabilities: readonly OfficialCapabilityName[];
};

const rawTrimmedBuildOptionsSchema = z
  .strictObject({
    ...buildOptionsShape,
    capabilities: capabilitiesSelectionSchema,
  })
  .readonly();
export const trimmedBuildOptionsSchema: z.ZodType<TrimmedBuildOptions> =
  rawTrimmedBuildOptionsSchema satisfies z.ZodType<
    AssertExact<
      TrimmedBuildOptions,
      z.infer<typeof rawTrimmedBuildOptionsSchema>
    > extends true
      ? TrimmedBuildOptions
      : never
  >;

function validateCapabilityNames(
  names: unknown,
): readonly OfficialCapabilityName[] {
  const parsed = capabilitiesSelectionSchema.safeParse(names);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    if (issue !== undefined && issue.path.length > 0) {
      const index = Number(issue.path[0]);
      const badName = Array.isArray(names) ? names[index] : undefined;
      throw new BuildError([
        toBuildDiagnostic(
          `build.capabilities[${index}]`,
          "unknown-capability",
          `Unknown official capability ${JSON.stringify(badName)}`,
          {
            capability: typeof badName === "string" ? badName : String(badName),
            field: "capabilities",
          },
        ),
      ]);
    }
    throw new BuildError([
      toBuildDiagnostic(
        "build.capabilities",
        "invalid-capability-selection",
        "capabilities must be an array of official capability names",
        { field: "capabilities" },
      ),
    ]);
  }
  return [...new Set(parsed.data)];
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
  const parsed = trimmedResolveOptionsSchema.safeParse(options);
  if (!parsed.success) {
    const issue = parsed.error.issues[0];
    if (issue?.path[0] === "capabilities") {
      validateCapabilityNames(
        (options as { readonly capabilities?: unknown }).capabilities,
      );
    }
    if (issue?.path.length === 0 && issue.code === "unrecognized_keys") {
      throw new BuildError([
        toBuildDiagnostic(
          "resolve",
          "unknown-parameter",
          `Resolve options contain unknown parameter ${issue.keys[0]}`,
          { prop: String(issue.keys[0]), scope: "options" },
        ),
      ]);
    }
    if (issue?.path[0] === "signal") {
      throw new BuildError([
        toBuildDiagnostic(
          "resolve.signal",
          "invalid-options",
          "Resolve options signal must be an AbortSignal",
          { field: "signal" },
        ),
      ]);
    }
    throw new BuildError([
      toBuildDiagnostic(
        "resolve",
        "invalid-options",
        "Build options must be an object",
      ),
    ]);
  }
  const capabilities = validateCapabilityNames(parsed.data.capabilities);
  return resolveSelected(input, capabilities, parsed.data.signal);
}

/**
 * Loads only the selected official capabilities, resolves the document, and emits a PPTX archive.
 */
export async function build(
  input: ReactNode | CommonNode,
  options: TrimmedBuildOptions,
): Promise<BuildResult> {
  const parsed = trimmedBuildOptionsSchema.safeParse(options);
  if (!parsed.success) {
    if (
      typeof options !== "object" ||
      options === null ||
      Array.isArray(options)
    ) {
      throw new BuildError([
        toBuildDiagnostic(
          "build",
          "invalid-options",
          "Build options must be an object",
        ),
      ]);
    }
    const { capabilities, ...buildOptions } = options as Record<
      string,
      unknown
    >;
    validateOptions(buildOptions);
    validateCapabilityNames(capabilities);
    throw new BuildError([
      toBuildDiagnostic(
        "build",
        "invalid-options",
        "Build options are invalid",
      ),
    ]);
  }
  const capabilities = validateCapabilityNames(parsed.data.capabilities);
  const { capabilities: _capabilities, ...buildOptions } = parsed.data;
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
  CodegenError,
  ImportError,
  buildDiagnosticFontUnavailableReasonSchema,
  buildDiagnosticSchema,
  buildOptionsSchema,
  cascadedOmissionSchema,
  codegenDiagnosticCodeSchema,
  codegenDiagnosticSchema,
  codegenOptionsSchema,
  codegenTargetSchema,
  createCodegen,
  createImporter,
  createResolver,
  editabilityLossSchema,
  emit,
  emitOptionsSchema,
  extensionSchema,
  fatalImportIssueSchema,
  fontEmbeddingModeSchema,
  formatBaseline,
  formatBaselineSchema,
  importDiagnosticSchema,
  importLocationSchema,
  importOptionsSchema,
  importReportSchema,
  previewErrorSchema,
  previewPlaceholderSchema,
  resolveOptionsSchema,
  targetNameSchema,
  targetSchema,
  targets,
};
export type {
  CodegenDiagnostic,
  CodegenDiagnosticCode,
  CodegenEntryName,
  CodegenFileContent,
  CodegenOptions,
  CodegenResult,
  CodegenTarget,
};
export type {
  CascadedOmission,
  EditabilityLoss,
  FatalImportIssue,
  ImportDiagnostic,
  ImportLocation,
  ImportOptions,
  ImportReport,
  ImportResult,
} from "@pptxdk/import";
export type {
  BuildDiagnostic,
  BuildDiagnosticFontUnavailableReason,
  BuildOptions,
  BuildReport,
  BuildResult,
  CapabilityModule,
  DeclarationChild,
  DeclarationContainer,
  DeclarationNode,
  DeclarationText,
  EmitOptions,
  EmitResult,
  Extension,
  FontEmbeddingMode,
  FontStyleDescriptor,
  FontUsage,
  Locator,
  LocatorResource,
  ObjectTransform,
  PreviewError,
  PreviewPlaceholder,
  Rect,
  Resolution,
  ResolutionDimensionsSummary,
  ResolutionReport,
  ResolvedFooter,
  ResolvedFooterSlot,
  ResolvedImage,
  ResolvedLayout,
  ResolvedMaster,
  ResolvedObject,
  ResolvedPlaceholder,
  ResolvedSlide,
  ResolvedSlideLayout,
  ResolvedTable,
  ResolvedTableCell,
  ResolvedTableRow,
  ResolveOptions,
  ScreenPoint,
  SlideLocator,
  Target,
  TargetName,
  TextMetrics,
  UnmeasuredFontIssue,
} from "@pptxdk/core/build";
