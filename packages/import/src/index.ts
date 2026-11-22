export { readAlternateContent } from "#/alternate-content.ts";
export type { AlternateContentSelection } from "#/alternate-content.ts";
export {
  CascadedOmissionSchema,
  EditabilityLossSchema,
  FatalImportIssueSchema,
  ImportDiagnosticSchema,
  ImportError,
  ImportLocationSchema,
  ImportOptionsSchema,
  ImportReportSchema,
} from "#/diagnostics.ts";
export type {
  CascadedOmission,
  EditabilityLoss,
  FatalImportIssue,
  ImportDiagnostic,
  ImportLocation,
  ImportOptions,
  ImportReport,
  UnconsumedXmlItem,
} from "#/diagnostics.ts";
export { createImporter, MalformedXmlPartError } from "#/import.ts";
export type { ImportResult } from "#/import.ts";
export { allocateScopedAuthorId, importMastersAndLayouts } from "#/masters.ts";
export { defineCanvasMechanism, defineShowMechanism } from "#/mechanism.ts";
export type {
  DeckImportContext,
  DeckImportHandler,
  DeckImportOutcome,
  DeferredSlideLink,
  ImportMechanismHandler,
  ShapeTreeImportContext,
  ShapeTreeImportHandler,
  ShapeTreeImportOutcome,
  SlideImportContext,
  SlideImportHandler,
  SlideImportOutcome,
} from "#/mechanism.ts";
export {
  A_NS,
  CHART_NS,
  CONTENT_TYPES_NS,
  DIAGRAM_NS,
  EMU_PER_PX,
  MC_NS,
  OPC_REL_NS,
  P_NS,
  REL_DIAGRAM_COLORS,
  REL_DIAGRAM_DATA,
  REL_DIAGRAM_DRAWING,
  REL_DIAGRAM_LAYOUT,
  REL_DIAGRAM_QUICK_STYLE,
  REL_HYPERLINK,
  REL_IMAGE,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  REL_THEME,
  R_NS,
  STRICT_NS_PREFIX,
  STRICT_P_NS,
  STRICT_REL_OFFICE_DOCUMENT,
  TABLE_NS_URI,
  XMLNS_NS,
  XML_NS,
  canonicalQName,
  emuToPx,
} from "#/namespaces.ts";
export {
  OpcRelationshipSchema,
  openPptxPackage,
  parseRelationshipsXml,
  relsPathFor,
  resolvePartPath,
} from "#/opc.ts";
export type {
  OpcRelationship,
  ParsedRelationshipEntry,
  PptxPackage,
} from "#/opc.ts";
export {
  buildRelationshipIndex,
  inspectUnsupportedGraphicFrame,
  peekShapeTreeObjectIdentity,
  resolveExternalRelationshipTarget,
  resolveInternalRelationshipTarget,
} from "#/resources.ts";
export type {
  RelationshipIndex,
  ResolvedExternalResource,
  ResolvedInternalResource,
} from "#/resources.ts";
export {
  ExtentEmuFragmentSchema,
  HexRgb6Schema,
  OffsetEmuFragmentSchema,
  OoxmlBoolSchema,
  OoxmlIntSchema,
  OoxmlNonNegativeIntSchema,
  OoxmlPositiveIntSchema,
  TransformAttributesFragmentSchema,
} from "#/xml-fragments.ts";
export type {
  ExtentEmuFragment,
  OffsetEmuFragment,
  TransformAttributesFragment,
} from "#/xml-fragments.ts";
export { XmlElementCursor, parseXmlPart } from "#/xml-reader.ts";
