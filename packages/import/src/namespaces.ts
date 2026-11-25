export const XMLNS_NS = "http://www.w3.org/2000/xmlns/";
export const XML_NS = "http://www.w3.org/XML/1998/namespace";
export const OPC_REL_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
export const CONTENT_TYPES_NS =
  "http://schemas.openxmlformats.org/package/2006/content-types";
export const MC_NS =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";

// Transitional OOXML (Supported Baseline)
export const P_NS =
  "http://schemas.openxmlformats.org/presentationml/2006/main";
export const P14_NS =
  "http://schemas.microsoft.com/office/powerpoint/2010/main";
export const P159_NS =
  "http://schemas.microsoft.com/office/powerpoint/2015/09/main";
export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const REL_OFFICE_DOCUMENT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
export const REL_CHART =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
export const REL_SLIDE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
export const REL_SLIDE_LAYOUT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
export const REL_SLIDE_MASTER =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
export const REL_THEME =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
export const REL_NOTES_MASTER =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
export const REL_NOTES_SLIDE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
export const REL_HANDOUT_MASTER =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster";
export const REL_THEME_OVERRIDE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride";
export const REL_FONT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
export const REL_IMAGE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
export const REL_HYPERLINK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
export const REL_DIAGRAM_DATA =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
export const REL_DIAGRAM_LAYOUT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout";
export const REL_DIAGRAM_QUICK_STYLE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle";
export const REL_DIAGRAM_COLORS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors";
export const REL_DIAGRAM_DRAWING =
  "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";

export const DIAGRAM_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/diagram";
export const CHART_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const TABLE_NS_URI =
  "http://schemas.openxmlformats.org/drawingml/2006/table";
export const P14_SECTION_LIST_EXT_URI =
  "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
export const P14_CREATION_ID_EXT_URI = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}";

// ISO/IEC 29500 Strict OOXML (Detect & Reject with unsupported-format-baseline)
export const STRICT_NS_PREFIX = "http://purl.oclc.org/ooxml/";
export const STRICT_REL_OFFICE_DOCUMENT =
  "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
export const STRICT_P_NS = "http://purl.oclc.org/ooxml/presentationml/main";

export const EMU_PER_PX = 9525;

/** Converts DrawingML English Metric Units (EMU) to logical pixels (96 DPI). */
export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

const canonicalPrefixes: ReadonlyMap<string, string> = new Map([
  [P_NS, "p"],
  [P14_NS, "p14"],
  [P159_NS, "p159"],
  [A_NS, "a"],
  [C_NS, "c"],
  [R_NS, "r"],
  [MC_NS, "mc"],
  [DIAGRAM_NS, "dgm"],
  [CHART_NS, "c"],
  [XML_NS, "xml"],
]);

/** Formats an element or attribute name with a canonical prefix for deterministic diagnostic XPaths. */
export function canonicalQName(
  namespaceUri: string | null | undefined,
  localName: string,
): string {
  if (!namespaceUri) return localName;
  const prefix = canonicalPrefixes.get(namespaceUri);
  if (prefix !== undefined) return `${prefix}:${localName}`;
  return `{${namespaceUri}}${localName}`;
}
