import type { NativePlaceholderRef } from "@pptxdk/core/capability";
import { placeholderRefXml } from "@pptxdk/emit";

export { placeholderRefXml };

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
