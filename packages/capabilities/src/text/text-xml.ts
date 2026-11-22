import type { NativePlaceholderRef } from "@pptxdk/core/capability";

/** Serializes a `<p:ph .../>` native placeholder reference inside `<p:nvPr>`. */
export function placeholderRefXml(nativeRef: NativePlaceholderRef): string {
  const attrs: string[] = [];
  if (nativeRef.explicitType !== false) attrs.push(`type="${nativeRef.type}"`);
  if (nativeRef.sz !== undefined) attrs.push(`sz="${nativeRef.sz}"`);
  if (nativeRef.orient !== undefined)
    attrs.push(`orient="${nativeRef.orient}"`);
  if (nativeRef.idx !== undefined) attrs.push(`idx="${nativeRef.idx}"`);
  if (nativeRef.hasCustomPrompt === true) attrs.push('hasCustomPrompt="1"');
  return `<p:ph${attrs.length === 0 ? "" : ` ${attrs.join(" ")}`}/>`;
}

/** Internal PresentationML text box / text placeholder wrapper for the text capability. */
export function textBoxXml(
  id: number,
  escapedName: string,
  _box: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  },
  textBody: string,
  properties: string,
  linkXml: string,
  placeholder?: {
    readonly nativeRef: NativePlaceholderRef;
    readonly spPrXml?: string;
  },
): string {
  if (placeholder !== undefined) {
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${placeholderRefXml(placeholder.nativeRef)}</p:nvPr></p:nvSpPr>${placeholder.spPrXml ?? properties}${textBody}</p:sp>`;
  }
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapedName}">${linkXml}</p:cNvPr><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>${properties}${textBody}</p:sp>`;
}
