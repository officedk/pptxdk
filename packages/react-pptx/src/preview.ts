import { documentCapabilities } from "@react-pptx/capabilities/document";
import { createDocumentPreview } from "@react-pptx/core/preview";

/**
 * Default formal static preview for plain editable documents.
 *
 * Runs without loading React or react-reconciler and reuses the exact same formal build pass
 * for PPTX, report, query, and per-page PNG output.
 */
export const preview = createDocumentPreview(documentCapabilities);

export type {
  PreviewDiagnostic,
  PreviewNodeObservation,
  PreviewOptions,
  PreviewPageResult,
  PreviewResult,
  PreviewSeverity,
  PreviewTransform,
  PreviewWarning,
  ScreenGeometry,
  ScreenPoint,
} from "@react-pptx/core/preview";
