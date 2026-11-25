import { documentCapabilities } from "@pptxdk/capabilities/all";
import {
  createResolver,
  type BuildQuery,
  type Resolution,
  type ResolutionReport,
} from "@pptxdk/core/build";
import type { PreviewTransform, PreviewWarning } from "@pptxdk/core/capability";
import type { CommonNode } from "@pptxdk/core/document";
import {
  PreviewPageRangeError,
  paint,
  paintOptionsSchema,
  previewDiagnosticSchema,
  previewOptionsSchema,
  previewPageNumberSchema,
  previewPagesSchema,
  previewSeveritySchema,
  screenGeometrySchema,
  screenPointSchema,
  validatePaintOptions,
  type PreviewOptions,
  type PreviewResult,
} from "@pptxdk/preview";
import type { ReactNode } from "react";

import { resolveDocumentInput } from "./compose.ts";

const resolveDocument = createResolver(documentCapabilities);

/**
 * Default static preview for both plain editable documents and JSX elements.
 *
 * Plain CommonNode documents run synchronously through resolve + paint without loading React or
 * react-reconciler; JSX inputs dynamically load React expansion on demand.
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
  const resolution = await resolveDocumentInput(
    resolveDocument,
    input,
    signalOption,
  );
  const { pages } = paint(resolution, validatedOptions);
  return Object.freeze({
    report: resolution.report,
    query: resolution.query,
    pages,
  });
}

export {
  PreviewPageRangeError,
  paint,
  paintOptionsSchema,
  previewDiagnosticSchema,
  previewOptionsSchema,
  previewPageNumberSchema,
  previewPagesSchema,
  previewSeveritySchema,
  screenGeometrySchema,
  screenPointSchema,
};
export type {
  BuildQuery,
  PreviewTransform,
  PreviewWarning,
  Resolution,
  ResolutionReport,
};
export type {
  PaintOptions,
  PaintResult,
  PreviewDiagnostic,
  PreviewNodeObservation,
  PreviewOptions,
  PreviewPageResult,
  PreviewResult,
  PreviewSeverity,
  ScreenGeometry,
  ScreenPoint,
} from "@pptxdk/preview";
