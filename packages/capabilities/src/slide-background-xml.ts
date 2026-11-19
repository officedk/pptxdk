import type { ThemeContext } from "@react-pptx/core/capability";
import { xmlEscape } from "@react-pptx/emit";

import { fillXml } from "./visual.ts";

/** Internal PresentationML slide background wrapper reusing the shared visual fill writer. */
export function slideBackgroundFillXml(
  theme: ThemeContext,
  style: Readonly<Record<string, unknown>>,
): string {
  return `<p:bg><p:bgPr>${fillXml(theme, style, undefined, undefined)}<a:effectLst/></p:bgPr></p:bg>`;
}

/** Internal PresentationML slide background wrapper for a registered image relationship. */
export function slideBackgroundImageXml(relationshipId: string): string {
  return `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="${xmlEscape(relationshipId)}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
}
