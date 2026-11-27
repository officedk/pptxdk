import type {
  CapabilityResource,
  ResolvedColor,
} from "@pptxdk/core/capability";
import { resolvedColorXml, xmlEscape } from "@pptxdk/emit";

import {
  emitImageFillXml,
  gradientFillXml,
  type ResolvedGradientSpec,
  type ResolvedImageFill,
  type ShapeBox,
} from "../shared/visual.ts";

/** Internal PresentationML slide background wrapper for explicit `"none"` clearing. */
export function slideBackgroundNoneXml(): string {
  return "<p:bg><p:bgPr><a:noFill/><a:effectLst/></p:bgPr></p:bg>";
}

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

/** Internal PresentationML slide background wrapper for a resolved image fill. */
export function slideBackgroundImageXml(
  fill: ResolvedImageFill,
  registerResource: (resource: CapabilityResource) => string,
  box: ShapeBox,
): string {
  if (
    fill.source.kind === "embedded" &&
    fill.fit === "fill" &&
    fill.opacity === undefined
  ) {
    const relationshipId = registerResource({
      bytes: fill.source.bytes,
      extension: fill.source.extension,
      contentType: fill.source.contentType,
      relationshipType:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    });
    return `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="${xmlEscape(relationshipId)}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
  }
  const blipFill = emitImageFillXml(fill, registerResource, box).replace(
    /^<a:blipFill rotWithShape="1">/u,
    "<a:blipFill>",
  );
  return `<p:bg><p:bgPr>${blipFill}<a:effectLst/></p:bgPr></p:bg>`;
}

/** Internal PresentationML slide background wrapper for a `<p:bgRef>` theme format-scheme reference. */
export function slideBackgroundRefXml(
  idx: 1001 | 1002 | 1003,
  color: ResolvedColor,
): string {
  return `<p:bg><p:bgRef idx="${idx}">${resolvedColorXml(color)}</p:bgRef></p:bg>`;
}
