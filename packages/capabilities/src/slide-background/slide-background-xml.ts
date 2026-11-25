import type { ResolvedColor } from "@pptxdk/core/capability";
import { resolvedColorXml, xmlEscape } from "@pptxdk/emit";

import {
  gradientFillXml,
  type ResolvedGradientSpec,
} from "../shared/visual.ts";

/** Internal PresentationML slide background wrapper for a resolved solid color. */
export function slideBackgroundSolidXml(color: ResolvedColor): string {
  return `<p:bg><p:bgPr><a:solidFill>${resolvedColorXml(color)}</a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
}

/** Internal PresentationML slide background wrapper for a resolved gradient. */
export function slideBackgroundGradientXml(
  gradient: ResolvedGradientSpec,
): string {
  return `<p:bg><p:bgPr>${gradientFillXml(gradient)}<a:effectLst/></p:bgPr></p:bg>`;
}

/** Internal PresentationML slide background wrapper for a registered image relationship. */
export function slideBackgroundImageXml(relationshipId: string): string {
  return `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="${xmlEscape(relationshipId)}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
}

/** Internal PresentationML slide background wrapper for a `<p:bgRef>` theme format-scheme reference. */
export function slideBackgroundRefXml(
  idx: 1001 | 1002 | 1003,
  color: ResolvedColor,
): string {
  return `<p:bg><p:bgRef idx="${idx}">${resolvedColorXml(color)}</p:bgRef></p:bg>`;
}
