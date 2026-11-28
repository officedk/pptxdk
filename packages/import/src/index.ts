export { readAlternateContent } from "#/alternate-content.ts";
export type { AlternateContentSelection } from "#/alternate-content.ts";
export {
  arrowTypeSchema,
  arrowTypes,
  readBlipFill,
  readDrawingColor,
  readEffectList,
  readGradientFill,
  readLineBorder,
  readPlaceholderSpPrVisuals,
  readTransformXml,
} from "#/visual.ts";
export type {
  ArrowType,
  ImportedLineBorder,
  ImportedTransform,
  PlaceholderSpPrVisualOptions,
} from "#/visual.ts";
export {
  ImportError,
  cascadedOmissionSchema,
  editabilityLossSchema,
  fatalImportIssueSchema,
  importDiagnosticSchema,
  importLocationSchema,
  importOptionsSchema,
  importReportSchema,
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
export {
  allocateScopedAuthorId,
  appendUnconsumedDiagnostics,
  consumeIdentityShapeTreeBoilerplate,
  importMastersAndLayouts,
} from "#/masters.ts";
export { inspectNotesAndHandoutMasters } from "#/notes-masters.ts";
export { importPresentationSections } from "#/sections.ts";
export {
  isNonDefaultColorMap,
  readColorMapElement,
  readColorMapOverride,
  readThemeOverridePart,
  readThemePart,
} from "#/themes.ts";
export {
  readEmbeddedFontList,
  readPresentationThemeFontScheme,
  unpackEotFontData,
} from "#/fonts.ts";
export type {
  EmbeddedFontRecord,
  EmbeddedFontSlot,
  EmbeddedFontSlotRecord,
  PresentationFontSchemeInfo,
  UnpackedEotResult,
} from "#/fonts.ts";
export { defineCanvasMechanism, defineShowMechanism } from "#/mechanism.ts";
export type {
  DeckImportContext,
  DeckImportHandler,
  DeckImportOutcome,
  DeferredSlideLink,
  ImportedPlaceholderSpPr,
  ImportMechanismHandler,
  OmittedSlideObjectInfo,
  PlaceholderVisualReader,
  ScopeConnectorResolutionContext,
  ShapeTreeImportContext,
  ShapeTreeImportHandler,
  ShapeTreeImportOutcome,
  SlideImportContext,
  SlideImportHandler,
  SlideImportOutcome,
  TextInlineImportContext,
  TextInlineImportHandler,
} from "#/mechanism.ts";
export {
  A_NS,
  CHART_NS,
  CONTENT_TYPES_NS,
  C_NS,
  DIAGRAM_NS,
  EMU_PER_PX,
  MC_NS,
  OPC_REL_NS,
  P14_CREATION_ID_EXT_URI,
  P14_NS,
  P14_SECTION_LIST_EXT_URI,
  P159_NS,
  P_NS,
  REL_CHART,
  REL_DIAGRAM_COLORS,
  REL_DIAGRAM_DATA,
  REL_DIAGRAM_DRAWING,
  REL_DIAGRAM_LAYOUT,
  REL_DIAGRAM_QUICK_STYLE,
  REL_FONT,
  REL_HANDOUT_MASTER,
  REL_HYPERLINK,
  REL_IMAGE,
  REL_NOTES_MASTER,
  REL_NOTES_SLIDE,
  REL_OFFICE_DOCUMENT,
  REL_SLIDE,
  REL_SLIDE_LAYOUT,
  REL_SLIDE_MASTER,
  REL_THEME,
  REL_THEME_OVERRIDE,
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
  opcRelationshipSchema,
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
  extentEmuFragmentSchema,
  hexRgb6Schema,
  offsetEmuFragmentSchema,
  ooxmlBoolSchema,
  ooxmlIntSchema,
  ooxmlNonNegativeIntSchema,
  ooxmlPositiveIntSchema,
  transformAttributesFragmentSchema,
} from "#/xml-fragments.ts";
export type {
  ExtentEmuFragment,
  OffsetEmuFragment,
  TransformAttributesFragment,
} from "#/xml-fragments.ts";
export { XmlElementCursor, parseXmlPart } from "#/xml-reader.ts";
