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
export const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const REL_OFFICE_DOCUMENT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
export const REL_SLIDE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
export const REL_SLIDE_LAYOUT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
export const REL_SLIDE_MASTER =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
export const REL_THEME =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

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
  [A_NS, "a"],
  [R_NS, "r"],
  [MC_NS, "mc"],
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
