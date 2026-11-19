import { documentCapabilities } from "@react-pptx/capabilities/document";
import {
  createResolver,
  type BuildQuery,
  type Resolution,
  type ResolutionReport,
} from "@react-pptx/core/build";
import type {
  PreviewTransform,
  PreviewWarning,
} from "@react-pptx/core/capability";
import { CommonNode } from "@react-pptx/core/document";
import {
  paint,
  type PreviewOptions,
  type PreviewResult,
} from "@react-pptx/preview";
import type { ReactNode } from "react";

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
  if (options.pages?.some((page) => !Number.isInteger(page) || page < 1)) {
    throw new RangeError("Preview page numbers must be positive integers");
  }
  if (options.signal?.aborted) throw options.signal.reason;
  const signalOption =
    options.signal === undefined ? {} : { signal: options.signal };
  if (input instanceof CommonNode) {
    const resolution = resolveDocument(input, signalOption);
    const { pages } = paint(resolution, options);
    return Object.freeze({
      report: resolution.report,
      query: resolution.query,
      pages,
    });
  }
  const { expandDocument } = await import("@react-pptx/react");
  const document = await expandDocument(input, signalOption);
  const resolution = resolveDocument(document, signalOption);
  const { pages } = paint(resolution, options);
  return Object.freeze({
    report: resolution.report,
    query: resolution.query,
    pages,
  });
}

export { paint };
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
} from "@react-pptx/preview";
